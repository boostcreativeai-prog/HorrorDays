// Protegge tutto il sito (pagina, data.json, Excel, font) con Basic Auth.
// Credenziali nelle variabili d'ambiente di Vercel: REPORT_USER e REPORT_PASSWORD.
// Se mancano il sito risponde 503, così non resta mai aperto per errore.

export const config = { matcher: "/:path*" };

const NO_STORE = { "Cache-Control": "no-store", "Content-Type": "text/plain; charset=utf-8" };

function equal(a, b) {
  // confronto a tempo costante sui byte UTF-8
  const x = new TextEncoder().encode(a);
  const y = new TextEncoder().encode(b);
  let diff = x.length ^ y.length;
  for (let i = 0; i < Math.max(x.length, y.length); i++) diff |= (x[i] ?? 0) ^ (y[i] ?? 0);
  return diff === 0;
}

function credentials(request) {
  const header = request.headers.get("authorization") || "";
  const [scheme, encoded] = header.split(" ");
  if (!/^basic$/i.test(scheme || "") || !encoded) return null;
  try {
    const bytes = Uint8Array.from(atob(encoded), (c) => c.charCodeAt(0));
    const decoded = new TextDecoder().decode(bytes);
    const i = decoded.indexOf(":");
    return i < 0 ? null : [decoded.slice(0, i), decoded.slice(i + 1)];
  } catch {
    return null;
  }
}

export default function middleware(request) {
  const user = process.env.REPORT_USER;
  const password = process.env.REPORT_PASSWORD;
  if (!user || !password) {
    return new Response("Report non configurato.", { status: 503, headers: NO_STORE });
  }

  const given = credentials(request);
  // valuta sempre entrambi i confronti
  const okUser = given ? equal(given[0], user) : false;
  const okPassword = given ? equal(given[1], password) : false;
  if (okUser && okPassword) return; // prosegue verso i file di site/

  return new Response("Accesso riservato.", {
    status: 401,
    headers: { ...NO_STORE, "WWW-Authenticate": 'Basic realm="Horror Days", charset="UTF-8"' },
  });
}
