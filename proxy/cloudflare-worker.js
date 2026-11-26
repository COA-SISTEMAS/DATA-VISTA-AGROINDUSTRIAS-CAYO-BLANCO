/**
 * DATA VISTA - SharePoint public-link download bridge
 * Handles anonymous SharePoint redirects and their temporary cookies.
 * No Microsoft credentials are stored.
 */
const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
  "Access-Control-Max-Age": "86400"
};

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      ...CORS,
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store"
    }
  });
}

function allowedHost(host) {
  return host === "sharepoint.com" ||
    host.endsWith(".sharepoint.com") ||
    host.endsWith(".sharepointonline.com") ||
    host.endsWith(".1drv.com");
}

function readSetCookies(headers) {
  try {
    return headers.getAll("Set-Cookie");
  } catch {
    const value = headers.get("Set-Cookie");
    return value ? [value] : [];
  }
}

function saveCookies(headers, responseHost, jar) {
  for (const line of readSetCookies(headers)) {
    const parts = line.split(";").map(part => part.trim());
    const pair = parts[0];
    const equals = pair.indexOf("=");
    if (equals <= 0) continue;

    let domain = responseHost;
    let path = "/";
    for (const part of parts.slice(1)) {
      const lower = part.toLowerCase();
      if (lower.startsWith("domain=")) domain = part.slice(7).replace(/^\./, "").toLowerCase();
      if (lower.startsWith("path=")) path = part.slice(5) || "/";
    }

    jar.set(parts[0].slice(0, equals), {
      value: pair,
      domain,
      path
    });
  }
}

function cookieHeader(url, jar) {
  const host = url.hostname.toLowerCase();
  const path = url.pathname || "/";
  const values = [];

  for (const cookie of jar.values()) {
    const domainMatch = host === cookie.domain || host.endsWith("." + cookie.domain);
    const pathMatch = path.startsWith(cookie.path);
    if (domainMatch && pathMatch && url.protocol === "https:") {
      values.push(cookie.value);
    }
  }

  return values.join("; ");
}

async function fetchSharePointFile(source) {
  const jar = new Map();
  let current = new URL(source);
  current.searchParams.set("download", "1");

  for (let hop = 0; hop < 9; hop++) {
    if (current.protocol !== "https:" || !allowedHost(current.hostname)) {
      throw new Error("SharePoint redirigió a un dominio no permitido.");
    }

    const headers = new Headers({
      "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel, text/csv, application/octet-stream, */*",
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36"
    });

    const cookies = cookieHeader(current, jar);
    if (cookies) headers.set("Cookie", cookies);

    const response = await fetch(current.toString(), {
      method: "GET",
      redirect: "manual",
      headers
    });

    saveCookies(response.headers, current.hostname, jar);

    if ([301, 302, 303, 307, 308].includes(response.status)) {
      const location = response.headers.get("Location");
      if (!location) throw new Error("SharePoint devolvió una redirección sin destino.");
      current = new URL(location, current);
      continue;
    }

    return response;
  }

  throw new Error("SharePoint realizó demasiadas redirecciones.");
}

export default {
  async fetch(request) {
    if (request.method === "OPTIONS") {
      return new Response(null, { headers: CORS });
    }

    if (request.method !== "GET") {
      return json({ error: "Método no permitido." }, 405);
    }

    const raw = new URL(request.url).searchParams.get("url");
    if (!raw) return json({ error: "Falta el parámetro url." }, 400);

    let source;
    try {
      source = new URL(raw);
    } catch {
      return json({ error: "El enlace no es una URL válida." }, 400);
    }

    if (source.protocol !== "https:" ||
        !(source.hostname === "sharepoint.com" || source.hostname.endsWith(".sharepoint.com"))) {
      return json({ error: "Solo se permiten enlaces HTTPS de SharePoint." }, 400);
    }

    try {
      const upstream = await fetchSharePointFile(source);

      if (!upstream.ok) {
        return json({
          error: "SharePoint respondió HTTP " + upstream.status +
            ". El servicio siguió las redirecciones, pero SharePoint no entregó el archivo."
        }, 502);
      }

      const contentType = (upstream.headers.get("content-type") || "").toLowerCase();
      if (contentType.includes("text/html")) {
        return json({
          error: "SharePoint entregó una página web y no el archivo. Revisa que el enlace permita descargarlo."
        }, 502);
      }

      const bytes = await upstream.arrayBuffer();
      if (!bytes.byteLength) {
        return json({ error: "SharePoint devolvió un archivo vacío." }, 502);
      }
      if (bytes.byteLength > 50 * 1024 * 1024) {
        return json({ error: "El archivo supera el límite de 50 MB." }, 413);
      }

      const headers = new Headers(CORS);
      headers.set("Content-Type", contentType || "application/octet-stream");
      headers.set("Cache-Control", "no-store");
      headers.set("Content-Disposition", 'attachment; filename="datavista-source.xlsx"');

      return new Response(bytes, { status: 200, headers });
    } catch (error) {
      return json({
        error: "No se pudo descargar desde SharePoint: " +
          (error?.message || "error de red")
      }, 502);
    }
  }
};