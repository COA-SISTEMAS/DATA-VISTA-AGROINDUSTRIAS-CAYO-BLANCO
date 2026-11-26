# Conectar DATA VISTA con Excel de SharePoint sin iniciar sesión

DATA VISTA puede consultar un vínculo de SharePoint configurado como **Cualquier persona con el vínculo – Puede ver**, siempre que el vínculo permita descargar el archivo. No necesita credenciales de Microsoft ni registro de aplicación Entra.

## Por qué se necesita un intermediario

SharePoint puede bloquear solicitudes `fetch()` desde un sitio web externo mediante CORS. El navegador no puede evitar esa política. Este Worker descarga el archivo desde el servidor y lo devuelve a DATA VISTA con encabezados CORS.

## Despliegue gratuito con Cloudflare Workers

1. Crea o inicia sesión en una cuenta de Cloudflare.
2. En el panel, abre **Workers & Pages** y crea un Worker.
3. Abre el editor de código del Worker.
4. Reemplaza el código de ejemplo por todo el contenido de `proxy/cloudflare-worker.js`.
5. Guarda y despliega el Worker.
6. Copia la URL pública que Cloudflare asigna, por ejemplo `https://nombre-del-worker.usuario.workers.dev`.
7. En DATA VISTA, abre **Fuentes de datos → Enlace de Excel**.
8. Pega el enlace compartido de SharePoint y la URL del Worker en **Servicio de conexión de DATA VISTA**.
9. Pulsa **Conectar y analizar**. La URL del Worker queda guardada en ese navegador.

## Actualizar los datos

Continúa editando el mismo libro en Excel Online. Cuando quieras actualizar DATA VISTA, vuelve a pulsar **Conectar y analizar**. La aplicación vuelve a descargar la versión más reciente del libro y recalcula los análisis.

## Condiciones y seguridad

- El vínculo de SharePoint debe ser anónimo y permitir descargar el archivo. Un vínculo de vista previa con descarga bloqueada no sirve.
- El Worker solo acepta URLs HTTPS cuyo host termine en `.sharepoint.com`; no es un proxy web abierto.
- El Worker no almacena el libro ni guarda credenciales de Microsoft.
- El vínculo «Cualquier persona» puede ser reenviado. Úsalo únicamente con datos que puedas compartir públicamente.
- El Worker limita los archivos a 50 MB.
- El plan gratuito de Cloudflare Workers puede tener límites de uso.
- Si la organización bloquea las descargas anónimas, esta alternativa no funcionará y habrá que usar autenticación autorizada.

## Archivos

- Código del Worker: `proxy/cloudflare-worker.js`
