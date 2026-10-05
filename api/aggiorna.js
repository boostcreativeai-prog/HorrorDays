// Pulsante "Aggiorna ora" del sito: avvia il workflow GitHub che scarica i C1 e ripubblica i dati.
// GET  /api/aggiorna  stato dell'ultimo aggiornamento (in corso o no)
// POST /api/aggiorna  avvia un aggiornamento, se non ce n'è già uno in corso
// Variabile Vercel: GH_DISPATCH_TOKEN, token GitHub fine-grained con permesso Actions: Read and write
// sul solo repository HorrorDays. La password del sito (middleware.js) protegge anche questa funzione.

const REPO = process.env.GH_REPO || "boostcreativeai-prog/HorrorDays";
const WORKFLOW = "aggiorna-report.yml";
const IN_CORSO = ["queued", "in_progress", "waiting", "pending", "requested"];

function json(body, status = 200, cdnSeconds = 0) {
  const headers = { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" };
  // lo stato è uguale per tutti: la CDN di Vercel lo tiene per pochi secondi, così tante pagine aperte
  // che controllano ogni 20 secondi non consumano il limite di chiamate all'API di GitHub
  if (cdnSeconds) headers["Vercel-CDN-Cache-Control"] = `max-age=${cdnSeconds}`;
  return new Response(JSON.stringify(body), { status, headers });
}

async function github(path, init = {}) {
  return fetch(`https://api.github.com/repos/${REPO}/actions/workflows/${WORKFLOW}${path}`, {
    ...init,
    headers: {
      Accept: "application/vnd.github+json",
      Authorization: `Bearer ${process.env.GH_DISPATCH_TOKEN}`,
      "X-GitHub-Api-Version": "2022-11-28",
      "User-Agent": "horrordays-report",
      ...(init.body ? { "Content-Type": "application/json" } : {}),
    },
  });
}

async function ultimo() {
  const r = await github("/runs?per_page=1");
  if (!r.ok) throw new Error(`GitHub ${r.status}`);
  const run = (await r.json()).workflow_runs?.[0];
  if (!run) return { inCorso: false };
  return {
    inCorso: IN_CORSO.includes(run.status),
    avviato: run.created_at,
    concluso: run.status === "completed" ? run.updated_at : null,
    esito: run.conclusion,
  };
}

export async function GET() {
  if (!process.env.GH_DISPATCH_TOKEN) return json({ errore: "Aggiornamento manuale non configurato." }, 503);
  try {
    return json(await ultimo(), 200, 10);
  } catch (e) {
    return json({ errore: e.message }, 502);
  }
}

export async function POST() {
  if (!process.env.GH_DISPATCH_TOKEN) return json({ errore: "Aggiornamento manuale non configurato." }, 503);
  try {
    const stato = await ultimo();
    if (stato.inCorso) return json({ ...stato, avviatoOra: false });
    const r = await github("/dispatches", { method: "POST", body: JSON.stringify({ ref: "main" }) });
    if (r.status !== 204) return json({ errore: `GitHub ${r.status}` }, 502);
    return json({ inCorso: true, avviato: new Date().toISOString(), avviatoOra: true }, 202);
  } catch (e) {
    return json({ errore: e.message }, 502);
  }
}
