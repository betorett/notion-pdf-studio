# Notion PDF Studio

Programa local para exportar paginas de Notion a PDF, Word y Google Docs con seleccion multiple, filtros por propiedades, DBs guardadas, preview con saltos de pagina, formulas KaTeX, diagramas Mermaid y ajustes de margenes.

## Arranque

```powershell
npm start
```

Abre `http://localhost:4173`.

El puerto se puede cambiar con `.env` o variable de entorno:

```text
PORT=4173
```

## Instalador Windows

Genera un instalador ligero de escritorio con Node embebido y apertura en Chrome/Edge como app:

```powershell
npm run dist:win
```

El instalador autocontenido queda en:

```text
dist\NotionPDFStudio-Setup.ps1
```

Instala por usuario en `%LOCALAPPDATA%\Programs\Notion PDF Studio`, crea acceso directo en Escritorio y Menu Inicio, registra desinstalador y no requiere Node.js instalado en el equipo final.

## Conectar Notion

1. Crea una integracion interna en Notion y copia el token.
2. Comparte la base de datos con esa integracion desde Notion (`Add connections`).
3. Pega el token en la app.
4. Pega el ID o la URL de la database/data source.
5. Pulsa `Sincronizar`.

Tambien puedes crear un archivo `.env`:

```text
NOTION_TOKEN=secret_xxxxxxxxx
```

## Flujo de exportacion

1. Sincroniza una database/data source de Notion.
2. Anade filtros por propiedades.
3. Selecciona varias paginas.
4. Pulsa `Vista previa` para descargar los bloques y ver saltos de pagina estimados.
5. Ajusta papel, escala, margenes, headers, footers y propiedades.
6. Elige `Archivo unico` o `Archivos separados`.
7. Pulsa `Exportar PDF` y en el dialogo del navegador elige `Guardar como PDF`.
8. Pulsa `Exportar Word` para descargar `.docx`; si eliges salida separada, descarga un `.zip` con un `.docx` por pagina.
9. Pulsa `Google Docs` para generar el mismo `.docx` y abrir Google Drive; sube el archivo y abrelo con Google Docs.

## Alcance actual

- Usa la API publica de Notion con `data_sources` actual y fallback a databases antiguas.
- Renderiza bloques comunes: parrafos, headings, listas, todos, callouts, citas, divisores, codigo, tablas, formulas, Mermaid, imagenes y placeholders para media/unsupported.
- Exporta Word con tablas DOCX reales, Mermaid como imagen PNG y ecuaciones KaTeX convertidas a ecuaciones editables de Word cuando es posible.
- Permite ordenar, ocultar propiedades por DB y plegar el panel de ajustes de la vista previa.
- Resuelve nombres de paginas relacionadas cuando la integracion tiene permiso para leer la DB relacionada.
- Guarda DBs locales en `data/local-state.json`.
- No sube tu contenido a terceros; el servidor corre en local y llama directamente a la API de Notion.
- Expone `GET /api/health` para health checks de despliegue.

## Publicacion

Validacion:

```powershell
npm run check
```

Docker:

```powershell
docker build -t notion-pdf-studio .
docker run --rm -p 4173:4173 --env-file .env notion-pdf-studio
```

Para publicar en un servidor, define `PORT` y, si quieres token por entorno, `NOTION_TOKEN`. Si no defines token, el usuario puede pegarlo en la UI y se guarda localmente en `data/local-state.json`.

## Nota importante sobre fidelidad

Notion no expone una API publica para pedir "exportar como PDF exactamente igual al cliente de Notion". Esta app reconstruye el documento desde bloques y propiedades para aproximarse a la vista de Notion y dar control de PDF. Algunos bloques nuevos, embeds privados, relaciones/rollups largos o URLs de archivos temporales pueden requerir revision manual.

Para que una propiedad `relation` muestre nombres en lugar de IDs cortos, comparte tambien la database relacionada con la misma integracion de Notion.
