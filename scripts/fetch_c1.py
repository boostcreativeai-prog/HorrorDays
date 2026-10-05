"""Scarica da 18Tickets i Riepiloghi mensili C1 di HORROR DAYS e li salva in data/c1/.
Percorso sull'intranet di 18Tickets:
1. login su /intranet/ (campi employee[email] e employee[password]); la sessione scade presto,
   quindi il login si rifà a ogni esecuzione
2. dalla home /intranet/ si clicca "Riepiloghi di servizio (C1, C2)": la sezione si carica via AJAX
   (aprire /siae_reports/index direttamente dà HTTP 406). Si imposta la data 01/MM/AAAA nel campo
   "from", si preme "Mostra stato dei report per la data selezionata" (anche questo via AJAX),
   poi si sceglie l'evento nel menu film_ids e si preme "Stampa c1 mensile"
3. il report viene generato in modo asincrono: quando è pronto compare il popup
   "Il tuo report è pronto!" con il link "clicca qui" a un PDF firmato su DigitalOcean Spaces,
   scaricabile senza cookie. Se ci mette molto compare prima il popup "Attenzione" (chiede una mail):
   lo si ignora e si continua ad aspettare. Il link si cerca sia nella pagina sia nelle risposte di rete
   (XHR e websocket), e ogni mese usa una pagina nuova con un secondo tentativo se il primo non arriva
Configurazione solo da variabili d'ambiente (secret e variabili di GitHub):
  T18_USER, T18_PASSWORD   credenziali (obbligatorie)
  T18_LOGIN_URL            default https://drivein.18tickets.it/intranet/
  C1_MONTHS                default "10/2026,11/2026"
  C1_FILM_IDS              default "193028" (HORROR DAYS)
  C1_EVENTO                default "HORROR DAYS" (controllo sul contenuto del PDF)
Se qualcosa va storto lo script esce con errore: il workflow si ferma, il sito resta
all'ultimo aggiornamento valido e nella pagina dell'esecuzione trovi screenshot, HTML della pagina
e registro delle richieste di rete (fetch_error.png, fetch_error.html, fetch_network.log).
"""
import os, re, sys, time, html, pathlib, subprocess, urllib.request
from urllib.parse import urlsplit
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

OUT = pathlib.Path("data/c1")
LOGIN_URL = os.environ.get("T18_LOGIN_URL", "").strip() or "https://drivein.18tickets.it/intranet/"
_u = urlsplit(LOGIN_URL)
BASE = f"{_u.scheme}://{_u.netloc}/intranet"
MONTHS = [m.strip() for m in os.environ.get("C1_MONTHS", "10/2026,11/2026").split(",") if m.strip()]
FILM_IDS = os.environ.get("C1_FILM_IDS", "").strip() or "193028"
EVENTO = os.environ.get("C1_EVENTO", "").strip() or "HORROR DAYS"
REPORT_TIMEOUT = 240  # secondi di attesa massima per ogni tentativo di generazione del report
TENTATIVI = 2
NETLOG = pathlib.Path("fetch_network.log")
SPACES_URL = re.compile(r"https://[^\s\"'<>\\]*digitaloceanspaces[^\s\"'<>\\]*")


class Rete:
    """Registra le richieste di rete e cerca nelle risposte i link ai PDF su DigitalOcean Spaces."""

    def __init__(self, context):
        self.pending = []  # risposte da leggere: il corpo si legge fuori dall'handler
        self.links = []
        self.log = NETLOG.open("w", encoding="utf-8")
        context.on("response", self._response)
        context.on("page", lambda pg: pg.on("websocket", self._websocket))

    def note(self, msg):
        self.log.write(f"{time.strftime('%H:%M:%S')} {msg}\n")
        self.log.flush()

    def _response(self, r):
        self.note(f"{r.request.method} {r.status} {r.request.resource_type} {r.url.split('?')[0]}")
        if r.request.resource_type in ("xhr", "fetch", "document", "script", "other"):
            self.pending.append(r)

    def _websocket(self, ws):
        self.note(f"WS {ws.url.split('?')[0]}")
        ws.on("framereceived", lambda payload: self._text(payload if isinstance(payload, str) else "", "ws"))

    def _text(self, text, origine):
        text = html.unescape(text.replace("\\/", "/").replace("\\u0026", "&"))
        for url in SPACES_URL.findall(text):
            if url not in self.links:
                self.links.append(url)
                self.note(f"LINK ({origine}) {url.split('?')[0]}")

    def drain(self):
        while self.pending:
            r = self.pending.pop(0)
            try:
                self._text(r.text(), r.url.split("?")[0])
            except Exception:
                pass  # redirect o corpo non disponibile

    def find(self, mm, yyyy, skip):
        self.drain()
        nuovi = [u for u in self.links if u not in skip and f"C1_{mm}_{yyyy}" in u]
        return nuovi[-1] if nuovi else None


def env(name):
    v = os.environ.get(name, "").strip()
    if not v:
        sys.exit(f"Variabile d'ambiente mancante: {name}")
    return v


def login(page):
    page.goto(LOGIN_URL, wait_until="domcontentloaded")
    page.locator('input[name="employee[email]"]').first.fill(env("T18_USER"))
    page.locator('input[name="employee[password]"]').first.fill(env("T18_PASSWORD"))
    form = page.locator('form[action*="sign_in"]').first
    submit = form.locator('button[type=submit], input[type=submit]')
    if submit.count():
        submit.first.click()
    else:
        page.locator('input[name="employee[password]"]').first.press("Enter")
    page.wait_for_load_state("networkidle")

    otp = page.locator('input[name*="otp" i]:visible, input[autocomplete="one-time-code"]:visible').count()
    page.goto(f"{BASE}/", wait_until="domcontentloaded")
    if "sign_in" in page.url or page.locator('input[name="employee[password]"]').count() \
            or "richiesta autorizzazione" in page.content():
        if otp:
            sys.exit("18Tickets chiede un codice di verifica (2FA): il login automatico non è possibile.")
        sys.exit("Login non riuscito: controlla T18_USER e T18_PASSWORD.")
    print("Login effettuato", flush=True)


def report_selector(mm, yyyy):
    return f'a[href*="digitaloceanspaces"][href*="C1_{mm}_{yyyy}"]'


def wait_report_link(page, rete, mm, yyyy, before, skip):
    """Aspetta un NUOVO link al PDF del mese giusto, nel popup 'clicca qui' o nelle risposte di rete.
    Restituisce l'URL firmato, oppure None se non arriva entro REPORT_TIMEOUT."""
    sel = report_selector(mm, yyyy)
    deadline = time.monotonic() + REPORT_TIMEOUT
    attenzione = False
    while time.monotonic() < deadline:
        if page.locator(sel).count() > before:
            return page.locator(sel).last.get_attribute("href")
        url = rete.find(mm, yyyy, skip)
        if url:
            return url
        if not attenzione and page.get_by_text("richiedendo più tempo del previsto").count():
            attenzione = True
            print(f"{mm}/{yyyy}: 18Tickets dice che il report richiede più tempo del previsto, continuo ad attendere...", flush=True)
            rete.note(f"{mm}/{yyyy}: popup Attenzione")
        page.wait_for_timeout(2000)
    return None


def request_c1(page, rete, mm, yyyy):
    """Avvia la generazione del C1 mensile filtrato sull'evento.
    Restituisce quanti link al PDF del mese c'erano nella pagina prima del clic e i link già visti
    in rete, oppure None se nel mese l'evento non compare (mese saltato)."""
    # la sezione Riepiloghi si apre solo dal link nella home (caricamento AJAX)
    page.goto(f"{BASE}/", wait_until="networkidle")
    page.get_by_role("link", name=re.compile(r"Riepiloghi di servizio", re.I)).first.click()
    page.locator("input#from").wait_for(state="attached", timeout=60_000)
    # data al primo del mese, poi "Mostra stato dei report per la data selezionata" (AJAX)
    page.locator("input#from").evaluate("(el, v) => { el.value = v; }", f"01/{mm}/{yyyy}")
    page.locator('form[action*="siae_reports/show"] input[type=submit]').first.click()
    # la sezione è aggiornata quando il menu eventi mostra il mese richiesto
    page.locator('select[name="film_ids"] option', has_text=f"{mm}/{yyyy}").first \
        .wait_for(state="attached", timeout=60_000)

    film_id = FILM_IDS.split(",")[0]
    if not page.locator(f'select[name="film_ids"] option[value="{film_id}"]').count():
        print(f"{mm}/{yyyy}: {EVENTO} non compare tra gli eventi del mese, salto.", flush=True)
        return None
    page.locator('select[name="film_ids"]').first.select_option(film_id)

    before = page.locator(report_selector(mm, yyyy)).count()
    rete.drain()
    skip = set(rete.links)
    btn = page.get_by_role("button", name=re.compile(r"^\s*Stampa c1 mensile\s*$", re.I))
    if not btn.count():
        sys.exit(f"{mm}/{yyyy}: pulsante 'Stampa c1 mensile' non trovato nella pagina dei Riepiloghi.")
    btn.first.click()
    print(f"{mm}/{yyyy}: richiesta C1 mensile inviata, attendo il report...", flush=True)
    return before, skip


def fetch_month(context, rete, mm, yyyy):
    """Chiede il C1 del mese in una pagina nuova, con un secondo tentativo. Restituisce l'URL del PDF,
    None se l'evento non c'è nel mese; se il report non arriva lascia la pagina aperta ed esce con errore."""
    for tentativo in range(1, TENTATIVI + 1):
        page = context.new_page()
        rete.note(f"--- {mm}/{yyyy} tentativo {tentativo}")
        req = request_c1(page, rete, mm, yyyy)
        if req is None:
            page.close()
            return None
        before, skip = req
        url = wait_report_link(page, rete, mm, yyyy, before, skip)
        if url:
            page.close()
            return url
        print(f"{mm}/{yyyy}: report non arrivato in {REPORT_TIMEOUT} s (tentativo {tentativo} di {TENTATIVI}).", flush=True)
        if tentativo < TENTATIVI:
            page.close()
    raise PWTimeout(f"il C1 di {mm}/{yyyy} non è arrivato dopo {TENTATIVI} tentativi")


def save_pdf(url, target):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=60) as r:
        data = r.read()
    if data[:4] != b"%PDF":
        sys.exit(f"Il file scaricato per {target.name} non è un PDF.")
    tmp = target.with_suffix(".tmp")
    tmp.write_bytes(data)
    text = subprocess.run(["pdftotext", "-layout", str(tmp), "-"], capture_output=True, text=True).stdout
    pages = text.count("QUADRO A")
    evento = len(re.findall(r"MANIFESTAZIONE\s+" + re.escape(EVENTO), text))
    if pages == 0 or evento != pages:
        tmp.unlink(missing_ok=True)
        sys.exit(f"{target.name}: {evento} pagine di {EVENTO} su {pages}. Il filtro evento non ha funzionato.")
    tmp.replace(target)
    print(f"Scaricato {target} ({pages} spettacoli)", flush=True)


def save_error(context):
    page = context.pages[-1] if context.pages else None
    if page is None:
        return None
    try:
        page.screenshot(path="fetch_error.png", full_page=True)
        pathlib.Path("fetch_error.html").write_text(page.content(), encoding="utf-8")
    except Exception as e:
        print(f"Impossibile salvare la diagnostica: {e}", flush=True)
    return page.url


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch()
        context = browser.new_context(accept_downloads=True, locale="it-IT")
        rete = Rete(context)
        try:
            page = context.new_page()
            login(page)
            page.close()
            for mese in MONTHS:
                mm, yyyy = mese.split("/")
                url = fetch_month(context, rete, mm, yyyy)
                if url:
                    save_pdf(url, OUT / f"Riepilogo_mensile_C1_{mm}_{yyyy}.pdf")
        except PWTimeout as e:
            where = save_error(context)
            sys.exit(f"Timeout su 18Tickets ({where}): {e}")
        except SystemExit:
            save_error(context)
            raise
        finally:
            rete.log.close()
            browser.close()


if __name__ == "__main__":
    main()
