/**
 * DATA VISTA - SharePoint public-link download bridge
 * Deploy as a Cloudflare Worker. No Microsoft credentials are stored.
 * Only accepts HTTPS URLs hosted on *.sharepoint.com.
 */
const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
  "Access-Control-Max-Age": "86400"
};
function json(data, status=200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { ...CORS, "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store" }
  });
}
export default {
  async fetch(request) {
    if (request.method === "OPTIONS") return new Response(null, { headers: CORS });
    if (request.method !== "GET") return json({ error: "Método no permitido." }, 405);
    const raw = new URL(request.url).searchParams.get("url");
    if (!raw) return json({ error: "Falta el parámetro url." }, 400);
    let source;
    try { source = new URL(raw); } catch { return json({ error: "El enlace no es una URL válida." }, 400); }
    if (source.protocol !== "https:" || !(source.hostname === "sharepoint.com" || source.hostname.endsWith(".sharepoint.com"))) {
      return json({ error: "Solo se permiten enlaces HTTPS de SharePoint." }, 400);
    }
    // Keep the original sharing token and request the downloadable workbook.
    source.searchParams.set("download", "1");
    try {
      const upstream = await fetch(source.toString(), {
        method: "GET",
        redirect: "follow",
        headers: { "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel, text/csv, application/octet-stream, */*" }
      });
      if (!upstream.ok) return json({ error: "SharePoint respondió HTTP " + upstream.status + ". Revisa que el vínculo permita descargar el archivo." }, 502);
      const contentType = (upstream.headers.get("content-type") || "").toLowerCase();
      if (contentType.includes("text/html")) return json({ error: "SharePoint entregó una página web y no el archivo. Verifica el permiso de descarga del vínculo." }, 502);
      const bytes = await upstream.arrayBuffer();
      if (!bytes.byteLength) return json({ error: "SharePoint devolvió un archivo vacío." }, 502);
      if (bytes.byteLength > 50 * 1024 * 1024) return json({ error: "El archivo supera el límite de 50 MB." }, 413);
      const headers = new Headers(CORS);
      headers.set("Content-Type", contentType || "application/octet-stream");
      headers.set("Content-Length", String(bytes.byteLength));
      headers.set("Cache-Control", "no-store");
      headers.set("Content-Disposition", 'attachment; filename="datavista-source.xlsx"');
      return new Response(bytes, { status: 200, headers });
    } catch (error) {
      return json({ error: "No se pudo descargar desde SharePoint: " + (error?.message || "error de red") }, 502);
    }
  }
};