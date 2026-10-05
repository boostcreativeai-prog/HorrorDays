# HORROR DAYS · Report vendite

Report vendite di HORROR DAYS (Drive In Pozzuoli, MIX MARK S.R.L.) aggiornato ogni ora
dai Riepiloghi mensili C1 SIAE della cassa 18Tickets e pubblicato su Vercel.

## Come funziona

Ogni ora, dalle 9 alle 24, GitHub Actions:

1. fa il login sull'intranet di 18Tickets, chiede il C1 mensile di ogni mese filtrato su HORROR DAYS e scarica il PDF quando è pronto (`scripts/fetch_c1.py`);
2. aggiorna `data/HorrorDays_Vendite.xlsx` e `site/data.json` (`scripts/build_report.py`), con lo storico conservato;
3. salva le modifiche nel repository: Vercel ripubblica il sito da solo.

Se il download fallisce il workflow si ferma, il sito resta all'ultimo aggiornamento valido
e GitHub ti manda un'email. Nella pagina dell'esecuzione trovi lo screenshot dell'errore.

## Struttura

```
.github/workflows/aggiorna-report.yml   esecuzione oraria
scripts/fetch_c1.py                     login e download da 18Tickets
scripts/build_report.py                 lettura C1, Excel, dati del sito, controlli
data/c1/                                PDF C1 (uno per mese, sovrascritti a ogni giro)
data/HorrorDays_Vendite.xlsx            report Excel (il foglio Tariffe si può modificare)
site/                                   sito pubblicato su Vercel
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
6. **Protezione.** Il report contiene dati di vendita: in Vercel attiva Settings > Deployment Protection
   (Vercel Authentication), così lo vede solo chi è nel tuo team.
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

Ogni giro dura 2-4 minuti (il report C1 viene generato in modo asincrono): 17 giri al giorno sono circa 1.000-2.000 minuti al mese,
al limite dei 2.000 gratuiti dei repository privati: se servisse, riduci la fascia oraria nel workflow. Per fermare gli aggiornamenti dopo il 1° novembre
disattiva il workflow in Actions.
