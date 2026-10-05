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
   scaricabile senza cookie
Configurazione solo da variabili d'ambiente (secret e variabili di GitHub):
  T18_USER, T18_PASSWORD   credenziali (obbligatorie)
  T18_LOGIN_URL            default https://drivein.18tickets.it/intranet/
  C1_MONTHS                default "10/2026,11/2026"
  C1_FILM_IDS              default "193028" (HORROR DAYS)
  C1_EVENTO                default "HORROR DAYS" (controllo sul contenuto del PDF)
Se qualcosa va storto lo script esce con errore: il workflow si ferma, il sito resta
all'ultimo aggiornamento valido e nella pagina dell'esecuzione trovi lo screenshot.
"""
import os, re, sys, pathlib, subprocess, urllib.request
from urllib.parse import urlsplit
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

OUT = pathlib.Path("data/c1")
LOGIN_URL = os.environ.get("T18_LOGIN_URL", "").strip() or "https://drivein.18tickets.it/intranet/"
_u = urlsplit(LOGIN_URL)
BASE = f"{_u.scheme}://{_u.netloc}/intranet"
MONTHS = [m.strip() for m in os.environ.get("C1_MONTHS", "10/2026,11/2026").split(",") if m.strip()]
FILM_IDS = os.environ.get("C1_FILM_IDS", "").strip() or "193028"
EVENTO = os.environ.get("C1_EVENTO", "").strip() or "HORROR DAYS"
REPORT_TIMEOUT = 240_000  # ms di attesa massima per la generazione del report


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


def wait_report_link(page, mm, yyyy, before=0):
    """Aspetta un NUOVO link 'clicca qui' del popup e restituisce l'URL firmato del PDF del mese giusto."""
    sel = report_selector(mm, yyyy)
    page.wait_for_function(
        "([s, n]) => document.querySelectorAll(s).length > n",
        arg=[sel, before],
        timeout=REPORT_TIMEOUT,
    )
    return page.locator(sel).last.get_attribute("href")


def request_c1(page, mm, yyyy):
    """Avvia la generazione del C1 mensile filtrato sull'evento.
    Restituisce quanti link al PDF del mese c'erano prima del clic,
    oppure None se nel mese l'evento non compare (mese saltato)."""
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
    btn = page.get_by_role("button", name=re.compile(r"^\s*Stampa c1 mensile\s*$", re.I))
    if not btn.count():
        sys.exit(f"{mm}/{yyyy}: pulsante 'Stampa c1 mensile' non trovato nella pagina dei Riepiloghi.")
    btn.first.click()
    print(f"{mm}/{yyyy}: richiesta C1 mensile inviata, attendo il report...", flush=True)
    return before


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


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_context(accept_downloads=True, locale="it-IT").new_page()
        try:
            login(page)
            for mese in MONTHS:
                mm, yyyy = mese.split("/")
                before = request_c1(page, mm, yyyy)
                if before is None:
                    continue
                url = wait_report_link(page, mm, yyyy, before)
                save_pdf(url, OUT / f"Riepilogo_mensile_C1_{mm}_{yyyy}.pdf")
        except PWTimeout as e:
            page.screenshot(path="fetch_error.png", full_page=True)
            sys.exit(f"Timeout su 18Tickets ({page.url}): {e}")
        except SystemExit:
            page.screenshot(path="fetch_error.png", full_page=True)
            raise
        finally:
            browser.close()


if __name__ == "__main__":
    main()
