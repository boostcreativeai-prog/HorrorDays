# HORROR DAYS · Report vendite

Report vendite di HORROR DAYS (Drive In Pozzuoli, MIX MARK S.R.L.) aggiornato ogni ora
dai Riepiloghi mensili C1 SIAE della cassa 18Tickets e pubblicato su Vercel.

## Come funziona

Ogni ora, giorno e notte, GitHub Actions:

1. fa il login sull'intranet di 18Tickets, chiede il C1 mensile di ogni mese filtrato su HORROR DAYS e scarica il PDF quando è pronto (`scripts/fetch_c1.py`);
2. aggiorna `data/HorrorDays_Vendite.xlsx` e `site/data.json` (`scripts/build_report.py`), con lo storico conservato;
3. salva le modifiche nel repository: Vercel ripubblica il sito da solo.

GitHub avvia i giri orari con 15-20 minuti di ritardo. Per avere i dati del momento c'è il pulsante
**Aggiorna ora** sul sito (`api/aggiorna.js`): avvia subito il workflow e in 1-2 minuti la pagina mostra i nuovi dati
senza bisogno di ricaricarla.

I numeri sono quelli fiscali del C1: biglietti emessi e incasso lordo **al netto degli annullati**,
quindi solo acquisti andati a buon fine.

**Auto e persone.** Ogni biglietto è un'auto, e la capienza del C1 (200 per turno) è in auto. Le persone
dipendono dal prezzo: 24 € = 2 persone, 36 € = 3, 48 € = 4 (sempre 12 € a persona). Il report mostra auto e
persone per turno, serata e tipologia. Un prezzo diverso (es. un omaggio a 0 €) finisce tra le anomalie:
indica quante persone vale nella colonna "Persone per auto" del foglio Tariffe dell'Excel in `data/`.
La capienza è di **100 auto per turno** (`CAPIENZA_AUTO` in `scripts/build_report.py`), non i 200 del C1.

**Quando si vende.** Il C1 non riporta l'ora di ogni acquisto. A ogni aggiornamento lo script confronta ogni turno
con l'aggiornamento precedente e salva le differenze in `data/movimenti.json` (registro vendite): così si sa
tra quali due aggiornamenti è avvenuta ogni vendita. Da qui il sito e l'Excel (fogli Registro vendite e
Picchi vendite) calcolano giorni, fasce orarie e giorni della settimana con più vendite. Le auto vendute prima
della prima rilevazione (05/10/2026 15:55) non hanno un orario.

Se il download fallisce il workflow si ferma, il sito resta all'ultimo aggiornamento valido
e GitHub ti manda un'email. Nella pagina dell'esecuzione trovi lo screenshot dell'errore.

## Struttura

```
.github/workflows/aggiorna-report.yml   esecuzione oraria
scripts/fetch_c1.py                     login e download da 18Tickets
scripts/build_report.py                 lettura C1, Excel, dati del sito, controlli
data/c1/                                PDF C1 (uno per mese, sovrascritti a ogni giro)
data/HorrorDays_Vendite.xlsx            report Excel (il foglio Tariffe si può modificare)
data/movimenti.json                     registro vendite (differenze tra un aggiornamento e l'altro)
site/                                   sito pubblicato su Vercel
middleware.js                           password del sito (Basic Auth)
api/aggiorna.js                         pulsante "Aggiorna ora" (avvia il workflow)
```

## Configurazione (una volta sola)

1. **Repository.** Crea un repository **privato** su GitHub e carica questi file.
2. **Secret.** In Settings > Secrets and variables > Actions > Secrets aggiungi:
   - `T18_USER` e `T18_PASSWORD`: le credenziali dell'intranet 18Tickets
   - `T18_LOGIN_URL` (facoltativo): di default `https://drivein.18tickets.it/intranet/`
3. **Mesi ed evento.** Di default scarica `10/2026,11/2026` filtrati su HORROR DAYS (`film_ids` 193028).
   Per cambiarli crea le variabili `C1_MONTHS` e `C1_FILM_IDS` in Settings > Secrets and variables > Actions > Variables.
4. **Verifica.** Lo script controlla che ogni pagina del PDF sia di HORROR DAYS: se il filtro evento non funziona si ferma.
5. **Vercel.** New Project > importa il repository. `vercel.json` imposta già la cartella `site` come output, senza build.
6. **Protezione.** Il report contiene dati di vendita: `middleware.js` chiede utente e password (Basic Auth)
   per tutto il sito, `data.json` compreso. In Vercel > Settings > Environment Variables crea `REPORT_USER` e
   `REPORT_PASSWORD` (ambiente Production) e ripubblica. Senza le due variabili il sito risponde 503.
   In Settings > Environments > Production la branch deve essere `main`. Tieni privato anche il repository GitHub.
   Per il pulsante "Aggiorna ora" crea su GitHub un token fine-grained (Settings > Developer settings >
   Personal access tokens > Fine-grained tokens) limitato al repository HorrorDays con permesso
   **Actions: Read and write**, e salvalo in Vercel come variabile `GH_DISPATCH_TOKEN`. Senza token il pulsante non compare.
7. **Prova.** In Actions > Aggiorna report vendite > Run workflow lancia un primo giro a mano.

## Font e logo

Il sito usa i colori e i font di horrordays.it: Horror Days Display (`site/fonts/`, solo per il titolo:
ha solo maiuscole e niente cifre), Anton per le intestazioni, Inter per i testi e VT323 per i numeri.
Il logo viene caricato da horrordays.it.

## Uso in locale

```
pip install -r requirements.txt     # serve anche poppler-utils (pdftotext)
python scripts/build_report.py      # usa i PDF già presenti in data/c1
python -m http.server -d site       # apri http://localhost:8000
```

## Minuti di GitHub Actions

Ogni giro dura circa un minuto (GitHub ne conteggia 1-2): 24 giri al giorno sono circa 750-1.500 minuti al mese,
più i clic su "Aggiorna ora", dentro i 2.000 gratuiti dei repository privati. Per fermare gli aggiornamenti dopo il 1° novembre
disattiva il workflow in Actions.
