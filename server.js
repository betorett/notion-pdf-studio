"use strict";

const http = require("node:http");
const fsSync = require("node:fs");
const fs = require("node:fs/promises");
const path = require("node:path");
const { URL } = require("node:url");
const {
  AlignmentType,
  BorderStyle,
  Document,
  ExternalHyperlink,
  Footer,
  Header,
  HeadingLevel,
  ImageRun,
  ImportedXmlComponent,
  Packer,
  PageBreak,
  PageNumber,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType
} = require("docx");
const JSZip = require("jszip");
const katex = require("katex");
const { mml2omml } = require("@hungknguyen/mathml2omml");

const ROOT = __dirname;
const PUBLIC_DIR = path.join(ROOT, "public");
const DATA_DIR = path.join(ROOT, "data");
const STATE_FILE = path.join(DATA_DIR, "local-state.json");
const PORT = Number(process.env.PORT || 4173);
const NOTION_VERSION = "2026-03-11";
const LEGACY_NOTION_VERSION = "2022-06-28";

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".svg": "image/svg+xml; charset=utf-8",
  ".ico": "image/x-icon",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".zip": "application/zip"
};

const DEMO_SOURCE_ID = "demo-notion-workspace";

loadDotEnv();

const demoSchema = {
  Name: { id: "title", type: "title", title: {} },
  Subject: {
    id: "subject",
    type: "relation",
    relation: { data_source_id: "demo-subjects" }
  },
  Status: {
    id: "status",
    type: "status",
    status: {
      options: [
        { id: "draft", name: "Draft", color: "gray" },
        { id: "ready", name: "Ready", color: "green" },
        { id: "review", name: "Review", color: "blue" }
      ]
    }
  },
  Tags: {
    id: "tags",
    type: "multi_select",
    multi_select: {
      options: [
        { id: "docs", name: "Docs", color: "blue" },
        { id: "finance", name: "Finance", color: "yellow" },
        { id: "client", name: "Client", color: "purple" }
      ]
    }
  },
  Owner: { id: "owner", type: "rich_text", rich_text: {} },
  Updated: { id: "updated", type: "date", date: {} },
  Approved: { id: "approved", type: "checkbox", checkbox: {} }
};

const demoPageRows = [
  {
    id: "demo-roadmap",
    title: "Q2 Product Roadmap",
    url: "https://notion.so/demo-roadmap",
    icon: { type: "emoji", emoji: "📘" },
    cover: null,
    created_time: "2026-04-01T10:00:00.000Z",
    last_edited_time: "2026-04-22T16:24:00.000Z",
    properties: {
      Name: "Q2 Product Roadmap",
      Status: "Ready",
      Tags: ["Docs", "Client"],
      Subject: relationValue([{ id: "demo-subject-marketing", title: "Marketing" }]),
      Owner: "Vicent",
      Updated: "2026-04-22",
      Approved: true
    }
  },
  {
    id: "demo-finance",
    title: "Finance Operating Notes",
    url: "https://notion.so/demo-finance",
    icon: { type: "emoji", emoji: "📊" },
    cover: null,
    created_time: "2026-03-12T10:00:00.000Z",
    last_edited_time: "2026-04-19T09:10:00.000Z",
    properties: {
      Name: "Finance Operating Notes",
      Status: "Review",
      Tags: ["Finance", "Docs"],
      Subject: relationValue([{ id: "demo-subject-finance", title: "Finance" }]),
      Owner: "Mar",
      Updated: "2026-04-19",
      Approved: false
    }
  },
  {
    id: "demo-client-brief",
    title: "Client Delivery Brief",
    url: "https://notion.so/demo-client-brief",
    icon: { type: "emoji", emoji: "🧭" },
    cover: null,
    created_time: "2026-04-10T10:00:00.000Z",
    last_edited_time: "2026-04-25T12:05:00.000Z",
    properties: {
      Name: "Client Delivery Brief",
      Status: "Draft",
      Tags: ["Client"],
      Subject: relationValue([{ id: "demo-subject-ops", title: "Operations" }]),
      Owner: "Laia",
      Updated: "2026-04-25",
      Approved: false
    }
  }
];

const demoBundles = {
  "demo-roadmap": {
    ...demoPageRows[0],
    blocks: [
      paragraph("This page demonstrates the local renderer: headings, lists, callouts, todos, tables, code, quotes and media-style placeholders are shaped to feel close to Notion while staying print-friendly."),
      heading("heading_1", "Objectives"),
      bulleted("Ship a reliable multi-page PDF export flow."),
      bulleted("Keep property filtering close to Notion database views."),
      bulleted("Expose margin, scale and header/footer controls before printing."),
      callout("💡", "Tip", "Use the preview pane to inspect the estimated page breaks before opening the browser print dialog."),
      heading("heading_2", "Delivery stages"),
      numbered("Connect integration and resolve database/data source schema."),
      numbered("Filter and select pages in bulk."),
      numbered("Fetch block trees recursively and render to a Notion-like document."),
      numbered("Print or save the composed output as PDF."),
      divider(),
      heading("heading_2", "Launch checklist"),
      todo("Shared database with the Notion integration", true),
      todo("Verified page permissions and read-content capability", true),
      todo("Reviewed final PDF margins", false),
      quote("The goal is not a screenshot. It is a printable document that keeps the structure, rhythm and hierarchy of the original Notion page."),
      table([
        ["Area", "Owner", "Status"],
        ["Renderer", "Vicent", "Ready"],
        ["Filters", "Mar", "Review"],
        ["PDF settings", "Laia", "Ready"]
      ])
    ]
  },
  "demo-finance": {
    ...demoPageRows[1],
    blocks: [
      heading("heading_1", "Monthly close notes"),
      paragraph("Use this page as an example of text-heavy operational content. The preview will split longer sections across several PDF sheets when needed."),
      heading("heading_2", "Revenue checks"),
      bulleted("Compare invoices exported from Notion with the accounting system."),
      bulleted("Confirm payment terms and outstanding balances."),
      bulleted("Attach final PDF to the monthly reporting packet."),
      code("const margin = { top: 18, right: 16, bottom: 18, left: 16 };\nprintDocument({ paper: 'A4', margin, scale: 0.96 });", "javascript"),
      equationBlock("E = mc^2"),
      code("flowchart LR\n  A[Notion] --> B[Vista previa]\n  B --> C[PDF final]\n  C --> D[Archivo listo]", "mermaid"),
      callout("⚠️", "Attention", "Formula, rollup and relation properties can be truncated by the API. For mission-critical exports, inspect the page property item endpoint when needed."),
      paragraph("A final PDF should preserve hierarchy: title, property context, section headings, tables, and page breaks should remain predictable.")
    ]
  },
  "demo-client-brief": {
    ...demoPageRows[2],
    blocks: [
      paragraph("This short brief is useful for testing multi-select and status filters."),
      heading("heading_2", "Scope"),
      bulleted("Export selected client-facing pages into a single PDF."),
      bulleted("Keep headers, footers and page numbers optional."),
      heading("heading_2", "Risks"),
      todo("Review image blocks with expiring Notion file URLs", false),
      todo("Check unsupported blocks before sending externally", false),
      divider(),
      paragraph("The renderer marks unsupported Notion blocks instead of hiding them, so you can catch gaps before sending a PDF.")
    ]
  }
};

function paragraph(text) {
  return {
    id: cryptoId("paragraph", text),
    type: "paragraph",
    has_children: false,
    paragraph: { rich_text: richText(text), color: "default" }
  };
}

function heading(type, text) {
  return {
    id: cryptoId(type, text),
    type,
    has_children: false,
    [type]: { rich_text: richText(text), color: "default", is_toggleable: false }
  };
}

function bulleted(text) {
  return {
    id: cryptoId("bullet", text),
    type: "bulleted_list_item",
    has_children: false,
    bulleted_list_item: { rich_text: richText(text), color: "default" }
  };
}

function numbered(text) {
  return {
    id: cryptoId("number", text),
    type: "numbered_list_item",
    has_children: false,
    numbered_list_item: { rich_text: richText(text), color: "default" }
  };
}

function todo(text, checked) {
  return {
    id: cryptoId("todo", text),
    type: "to_do",
    has_children: false,
    to_do: { rich_text: richText(text), checked, color: "default" }
  };
}

function quote(text) {
  return {
    id: cryptoId("quote", text),
    type: "quote",
    has_children: false,
    quote: { rich_text: richText(text), color: "default" }
  };
}

function callout(icon, title, body) {
  return {
    id: cryptoId("callout", title),
    type: "callout",
    has_children: false,
    callout: {
      icon: { type: "emoji", emoji: icon },
      rich_text: richText(`${title}: ${body}`),
      color: "gray_background"
    }
  };
}

function code(text, language) {
  return {
    id: cryptoId("code", language),
    type: "code",
    has_children: false,
    code: { rich_text: richText(text), language, caption: [] }
  };
}

function equationBlock(expression) {
  return {
    id: cryptoId("equation", expression),
    type: "equation",
    has_children: false,
    equation: { expression }
  };
}

function table(rows) {
  return {
    id: cryptoId("table", rows.flat().join("-")),
    type: "table",
    has_children: true,
    table: { table_width: rows[0].length, has_column_header: true, has_row_header: false },
    children: rows.map((cells, index) => ({
      id: cryptoId("row", `${index}-${cells.join("-")}`),
      type: "table_row",
      has_children: false,
      table_row: { cells: cells.map((cell) => richText(cell)) }
    }))
  };
}

function divider() {
  return {
    id: cryptoId("divider", String(Date.now())),
    type: "divider",
    has_children: false,
    divider: {}
  };
}

function richText(content) {
  return [
    {
      type: "text",
      text: { content, link: null },
      annotations: {
        bold: false,
        italic: false,
        strikethrough: false,
        underline: false,
        code: false,
        color: "default"
      },
      plain_text: content,
      href: null
    }
  ];
}

function relationValue(items, hasMore = false) {
  return { kind: "relation", items, hasMore };
}

function cryptoId(prefix, value) {
  let hash = 0;
  for (const char of `${prefix}:${value}`) {
    hash = (hash << 5) - hash + char.charCodeAt(0);
    hash |= 0;
  }
  return `${prefix}-${Math.abs(hash)}`;
}

function loadDotEnv() {
  try {
    const text = fsSync.readFileSync(path.join(ROOT, ".env"), "utf8");
    for (const line of text.split(/\r?\n/)) {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith("#") || !trimmed.includes("=")) continue;
      const [key, ...rest] = trimmed.split("=");
      if (!process.env[key]) process.env[key] = rest.join("=").trim().replace(/^["']|["']$/g, "");
    }
  } catch {
    // .env is optional; UI-supplied tokens still work.
  }
}

async function ensureState() {
  await fs.mkdir(DATA_DIR, { recursive: true });
  try {
    const raw = await fs.readFile(STATE_FILE, "utf8");
    return JSON.parse(raw);
  } catch {
    const initial = {
      notionToken: "",
      savedDatabases: [],
      exportHistory: []
    };
    await writeState(initial);
    return initial;
  }
}

async function writeState(state) {
  await fs.mkdir(DATA_DIR, { recursive: true });
  await fs.writeFile(STATE_FILE, `${JSON.stringify(state, null, 2)}\n`, "utf8");
}

function defaultPrintSettings() {
  return {
    paper: "A4",
    orientation: "portrait",
    marginTop: 18,
    marginRight: 16,
    marginBottom: 18,
    marginLeft: 16,
    scale: 96,
    showHeader: true,
    showFooter: true,
    startEachPageOnNewSheet: true,
    showProperties: true
  };
}

function sendJson(res, status, payload) {
  const body = JSON.stringify(payload);
  res.writeHead(status, responseHeaders({
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(body)
  }));
  res.end(body);
}

function sendBinary(res, status, buffer, contentType, filename) {
  res.writeHead(status, responseHeaders({
    "Content-Type": contentType,
    "Content-Length": buffer.length,
    "Content-Disposition": `attachment; filename="${downloadFileName(filename)}"`
  }));
  res.end(buffer);
}

function sendError(res, error) {
  const status = Number(error.status || 500);
  sendJson(res, status, {
    error: error.message || "Unexpected server error",
    details: error.details || undefined
  });
}

async function readJson(req) {
  const chunks = [];
  let length = 0;
  for await (const chunk of req) {
    length += chunk.length;
    if (length > 50 * 1024 * 1024) {
      throw httpError(413, "Request body is too large.");
    }
    chunks.push(chunk);
  }
  if (!chunks.length) return {};
  try {
    return JSON.parse(Buffer.concat(chunks).toString("utf8"));
  } catch {
    throw httpError(400, "Invalid JSON body.");
  }
}

function httpError(status, message, details) {
  const error = new Error(message);
  error.status = status;
  error.details = details;
  return error;
}

function responseHeaders(extra = {}) {
  return {
    "X-Content-Type-Options": "nosniff",
    "Referrer-Policy": "strict-origin-when-cross-origin",
    "Permissions-Policy": "camera=(), microphone=(), geolocation=()",
    ...extra
  };
}

function getToken(body = {}) {
  return String(body.token || process.env.NOTION_TOKEN || process.env.NOTION_API_KEY || "").trim();
}

function redactState(state) {
  return {
    ...state,
    notionToken: String(state.notionToken || "")
  };
}

function normalizeNotionId(input) {
  const value = String(input || "").trim();
  if (!value) return "";
  if (value === DEMO_SOURCE_ID) return value;
  const clean = value.split("?")[0].split("#")[0];
  const dashed = clean.match(/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/g);
  if (dashed?.length) return dashed[dashed.length - 1];
  const compact = clean.match(/[0-9a-fA-F]{32}/g);
  if (compact?.length) {
    const id = compact[compact.length - 1];
    return `${id.slice(0, 8)}-${id.slice(8, 12)}-${id.slice(12, 16)}-${id.slice(16, 20)}-${id.slice(20)}`;
  }
  return value;
}

async function notionFetch(token, endpoint, options = {}) {
  if (!token) throw httpError(401, "Missing Notion token. Paste one in the app or set NOTION_TOKEN in .env.");
  const response = await fetch(`https://api.notion.com/v1${endpoint}`, {
    method: options.method || "GET",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      "Notion-Version": options.version || NOTION_VERSION
    },
    body: options.body ? JSON.stringify(options.body) : undefined
  });
  const text = await response.text();
  const payload = text ? safeJson(text) : {};
  if (!response.ok) {
    throw httpError(
      response.status,
      payload.message || `Notion API returned HTTP ${response.status}.`,
      payload
    );
  }
  return payload;
}

function safeJson(text) {
  try {
    return JSON.parse(text);
  } catch {
    return { raw: text };
  }
}

async function resolveSource(token, input) {
  const id = normalizeNotionId(input);
  if (id === DEMO_SOURCE_ID) {
    throw httpError(404, "Demo workspace is disabled in production mode.");
  }
  if (!id) throw httpError(400, "Provide a Notion database or data source ID.");

  const attempts = [];

  try {
    const dataSource = await notionFetch(token, `/data_sources/${id}`);
    return sourceFromDataSource(dataSource, "data_source");
  } catch (error) {
    attempts.push({ mode: "data_source", status: error.status, message: error.message });
  }

  try {
    const database = await notionFetch(token, `/databases/${id}`);
    const firstSource = Array.isArray(database.data_sources) ? database.data_sources[0] : null;
    if (firstSource?.id) {
      const dataSource = await notionFetch(token, `/data_sources/${firstSource.id}`);
      return sourceFromDataSource(dataSource, "database_parent", database);
    }
    if (database.properties) return sourceFromLegacyDatabase(database, NOTION_VERSION);
  } catch (error) {
    attempts.push({ mode: "database_parent", status: error.status, message: error.message });
  }

  try {
    const legacy = await notionFetch(token, `/databases/${id}`, { version: LEGACY_NOTION_VERSION });
    return sourceFromLegacyDatabase(legacy, LEGACY_NOTION_VERSION);
  } catch (error) {
    attempts.push({ mode: "legacy_database", status: error.status, message: error.message });
  }

  throw httpError(
    404,
    "Could not resolve that ID as a Notion data source or database. Make sure the integration has access.",
    attempts
  );
}

function sourceFromDataSource(dataSource, mode, parentDatabase = null) {
  return {
    id: dataSource.id,
    sourceId: dataSource.id,
    mode,
    name: richTextToPlain(dataSource.title) || richTextToPlain(parentDatabase?.title) || "Untitled data source",
    schema: dataSource.properties || {},
    queryEndpoint: `/data_sources/${dataSource.id}/query`,
    version: NOTION_VERSION
  };
}

function sourceFromLegacyDatabase(database, version) {
  return {
    id: database.id,
    sourceId: database.id,
    mode: "legacy_database",
    name: richTextToPlain(database.title) || "Untitled database",
    schema: database.properties || {},
    queryEndpoint: `/databases/${database.id}/query`,
    version
  };
}

async function queryPages(token, source, filters = [], sorts = [], maxPages = 200) {
  if (source.mode === "demo") {
    return sortDemoRows(filterDemoRows(filters), sorts);
  }
  const notionFilter = buildNotionFilter(filters, source.schema);
  const bodyBase = {
    page_size: 100
  };
  if (notionFilter) bodyBase.filter = notionFilter;
  const normalizedSorts = normalizeSorts(sorts);
  if (normalizedSorts.length) bodyBase.sorts = normalizedSorts;

  const results = [];
  const relationTitleCache = new Map();
  let cursor;
  do {
    const body = cursor ? { ...bodyBase, start_cursor: cursor } : bodyBase;
    const payload = await notionFetch(token, source.queryEndpoint, {
      method: "POST",
      version: source.version,
      body
    });
    for (const item of payload.results || []) {
      if (item.object === "page") results.push(await summarizePage(item, source.schema, token, relationTitleCache));
      if (results.length >= maxPages) break;
    }
    cursor = payload.has_more && results.length < maxPages ? payload.next_cursor : null;
  } while (cursor);

  return results;
}

function filterDemoRows(filters = []) {
  const active = filters.filter((filter) => filter.property && filter.operator);
  if (!active.length) return demoPageRows;
  return demoPageRows.filter((row) =>
    active.every((filter) => {
      const value = row.properties[filter.property];
      const comparable = propertyComparable(value);
      const wanted = String(filter.value || "").toLowerCase();
      switch (filter.operator) {
        case "contains":
          return comparable.includes(wanted);
        case "does_not_contain":
          return !comparable.includes(wanted);
        case "equals":
          return comparable === wanted;
        case "does_not_equal":
          return comparable !== wanted;
        case "before":
          return String(value || "") < String(filter.value || "");
        case "after":
          return String(value || "") > String(filter.value || "");
        case "on_or_before":
          return String(value || "") <= String(filter.value || "");
        case "on_or_after":
          return String(value || "") >= String(filter.value || "");
        case "is_empty":
          return isPropertyEmpty(value);
        case "is_not_empty":
          return !isPropertyEmpty(value);
        default:
          return true;
      }
    })
  );
}

function sortDemoRows(rows, sorts = []) {
  const active = sorts.filter((sort) => sort.property && sort.direction);
  if (!active.length) return rows;
  return [...rows].sort((a, b) => {
    for (const sort of active) {
      const left = sort.property === "__last_edited_time"
        ? a.last_edited_time
        : sort.property === "__created_time"
          ? a.created_time
          : propertyComparable(a.properties[sort.property]);
      const right = sort.property === "__last_edited_time"
        ? b.last_edited_time
        : sort.property === "__created_time"
          ? b.created_time
          : propertyComparable(b.properties[sort.property]);
      const direction = sort.direction === "ascending" ? 1 : -1;
      const result = String(left).localeCompare(String(right), undefined, { numeric: true, sensitivity: "base" });
      if (result) return result * direction;
    }
    return 0;
  });
}

function buildNotionFilter(filters = [], schema = {}) {
  const clauses = [];
  for (const filter of filters) {
    if (!filter || !filter.property || !filter.operator) continue;
    const prop = schema[filter.property];
    if (!prop?.type) continue;
    const clause = buildFilterClause(filter.property, prop.type, filter.operator, filter.value);
    if (clause) clauses.push(clause);
  }
  if (!clauses.length) return null;
  return clauses.length === 1 ? clauses[0] : { and: clauses };
}

function buildFilterClause(property, type, operator, rawValue) {
  const emptyOperators = new Set(["is_empty", "is_not_empty"]);
  const valueRequired = !emptyOperators.has(operator);
  if (valueRequired && rawValue === "") return null;
  const value = emptyOperators.has(operator) ? true : rawValue;

  if (type === "title" || type === "rich_text" || type === "url" || type === "email" || type === "phone_number") {
    const filterType = type === "url" || type === "email" ? "rich_text" : type;
    return { property, [filterType]: { [operator]: value } };
  }
  if (type === "select" || type === "status") {
    return { property, [type]: { [operator]: value } };
  }
  if (type === "multi_select") {
    return { property, multi_select: { [operator]: value } };
  }
  if (type === "checkbox") {
    const checkboxOperator = operator === "does_not_equal" ? "does_not_equal" : "equals";
    return { property, checkbox: { [checkboxOperator]: rawValue === true || rawValue === "true" } };
  }
  if (type === "date") {
    return { property, date: { [operator]: value } };
  }
  if (type === "number") {
    const number = Number(rawValue);
    if (Number.isNaN(number) && valueRequired) return null;
    return { property, number: { [operator]: emptyOperators.has(operator) ? true : number } };
  }
  if (type === "created_time" || type === "last_edited_time") {
    return { property, [type]: { [operator]: value } };
  }
  if (type === "files" || type === "relation" || type === "people") {
    if (emptyOperators.has(operator)) return { property, [type]: { [operator]: true } };
    if ((type === "relation" || type === "people") && ["contains", "does_not_contain"].includes(operator)) {
      return { property, [type]: { [operator]: rawValue } };
    }
  }
  return null;
}

function normalizeSorts(sorts = []) {
  return sorts
    .filter((sort) => sort && sort.property && sort.direction)
    .map((sort) => {
      const direction = sort.direction === "ascending" ? "ascending" : "descending";
      if (sort.property === "__created_time") return { timestamp: "created_time", direction };
      if (sort.property === "__last_edited_time") return { timestamp: "last_edited_time", direction };
      return { property: sort.property, direction };
    });
}

async function fetchPageBundle(token, id) {
  if (demoBundles[id]) return demoBundles[id];
  const pageId = normalizeNotionId(id);
  const page = await notionFetch(token, `/pages/${pageId}`);
  const blocks = await fetchBlockChildren(token, pageId);
  return {
    ...(await summarizePage(page, {}, token)),
    blocks
  };
}

async function fetchBlockChildren(token, blockId, depth = 0, seen = new Set()) {
  if (depth > 12 || seen.has(blockId)) return [];
  seen.add(blockId);
  const results = [];
  let cursor;
  do {
    const query = new URLSearchParams({ page_size: "100" });
    if (cursor) query.set("start_cursor", cursor);
    const payload = await notionFetch(token, `/blocks/${blockId}/children?${query.toString()}`);
    for (const block of payload.results || []) {
      const next = { ...block };
      if (block.has_children) {
        next.children = await fetchBlockChildren(token, block.id, depth + 1, seen);
      }
      results.push(next);
    }
    cursor = payload.has_more ? payload.next_cursor : null;
  } while (cursor);
  return results;
}

async function summarizePage(page, schema = {}, token = "", relationTitleCache = new Map()) {
  const properties = {};
  for (const [name, value] of Object.entries(page.properties || {})) {
    properties[name] = await propertyToPlain(value, token, relationTitleCache);
  }
  return {
    id: page.id,
    title: findPageTitle(page) || "Untitled",
    url: page.url || "",
    icon: page.icon || null,
    cover: page.cover || null,
    created_time: page.created_time || "",
    last_edited_time: page.last_edited_time || "",
    properties
  };
}

function findPageTitle(page) {
  for (const value of Object.values(page.properties || {})) {
    if (value.type === "title") return richTextToPlain(value.title);
  }
  return "";
}

async function propertyToPlain(value, token = "", relationTitleCache = new Map()) {
  if (!value || !value.type) return "";
  switch (value.type) {
    case "title":
    case "rich_text":
      return richTextToPlain(value[value.type]);
    case "select":
    case "status":
      return value[value.type]?.name || "";
    case "multi_select":
      return (value.multi_select || []).map((item) => item.name);
    case "date":
      return value.date?.start || "";
    case "checkbox":
      return Boolean(value.checkbox);
    case "number":
      return value.number ?? "";
    case "url":
    case "email":
    case "phone_number":
      return value[value.type] || "";
    case "people":
      return (value.people || []).map((person) => person.name || person.id);
    case "files":
      return (value.files || []).map((file) => ({
        name: file.name || file.type,
        url: file.type === "external" ? file.external?.url : file.file?.url
      }));
    case "relation":
      return relationValue(await Promise.all((value.relation || []).map(async (item) => ({
        id: item.id,
        title: await resolveRelatedPageTitle(token, item.id, relationTitleCache)
      }))), Boolean(value.has_more));
    case "formula":
      return formulaToPlain(value.formula);
    case "rollup":
      return rollupToPlain(value.rollup, token, relationTitleCache);
    case "created_time":
    case "last_edited_time":
      return value[value.type] || "";
    case "unique_id":
      return `${value.unique_id?.prefix || ""}${value.unique_id?.number ?? ""}`;
    default:
      return "";
  }
}

async function resolveRelatedPageTitle(token, pageId, cache) {
  if (!pageId) return "";
  if (cache.has(pageId)) return cache.get(pageId);
  const fallback = shortId(pageId);
  if (!token || pageId.startsWith("demo-")) {
    cache.set(pageId, fallback);
    return fallback;
  }
  try {
    const page = await notionFetch(token, `/pages/${normalizeNotionId(pageId)}`);
    const title = findPageTitle(page) || fallback;
    cache.set(pageId, title);
    return title;
  } catch {
    cache.set(pageId, `${fallback} · no access`);
    return cache.get(pageId);
  }
}

function formulaToPlain(formula) {
  if (!formula?.type) return "";
  if (formula.type === "date") return formula.date?.start || "";
  return normalizePlainValue(formula[formula.type]);
}

async function rollupToPlain(rollup, token = "", relationTitleCache = new Map()) {
  if (!rollup?.type) return "";
  if (rollup.type === "array") {
    const values = await Promise.all((rollup.array || []).map((item) => propertySyncToPlain(item, token, relationTitleCache)));
    return normalizePlainValue(values);
  }
  if (rollup.type === "date") return rollup.date?.start || "";
  return normalizePlainValue(rollup[rollup.type]);
}

async function propertySyncToPlain(value, token = "", relationTitleCache = new Map()) {
  if (!value?.type) return "";
  switch (value.type) {
    case "title":
    case "rich_text":
      return richTextToPlain(value[value.type]);
    case "select":
    case "status":
      return value[value.type]?.name || "";
    case "multi_select":
      return (value.multi_select || []).map((item) => item.name);
    case "relation":
      return Promise.all((value.relation || []).map((item) => resolveRelatedPageTitle(token, item.id, relationTitleCache)));
    case "formula":
      return formulaToPlain(value.formula);
    case "rollup":
      return rollupToPlain(value.rollup, token, relationTitleCache);
    case "date":
      return value.date?.start || "";
    case "checkbox":
      return Boolean(value.checkbox);
    case "number":
      return value.number ?? "";
    default:
      return normalizePlainValue(value[value.type]);
  }
}

function normalizePlainValue(value) {
  if (Array.isArray(value)) return value.flatMap((item) => normalizePlainValue(item));
  if (value == null) return "";
  if (typeof value === "object") return value.name || value.plain_text || value.id || JSON.stringify(value);
  return value;
}

function propertyComparable(value) {
  if (Array.isArray(value)) return value.map(propertyComparable).join(" ").toLowerCase();
  if (value?.kind === "relation") return value.items.map((item) => item.title || item.id).join(" ").toLowerCase();
  if (value && typeof value === "object") return Object.values(value).flat().map(propertyComparable).join(" ").toLowerCase();
  return String(value ?? "").toLowerCase();
}

function isPropertyEmpty(value) {
  if (value?.kind === "relation") return !value.items.length;
  return value == null || value === "" || (Array.isArray(value) && value.length === 0);
}

function shortId(id) {
  return String(id || "").replaceAll("-", "").slice(0, 8);
}

function richTextToPlain(richText) {
  return Array.isArray(richText) ? richText.map((part) => part.plain_text || "").join("") : "";
}

function downloadFileName(filename) {
  return String(filename || "notion-export")
    .replace(/[\\/:*?"<>|]+/g, "-")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 140) || "notion-export";
}

function slugFileName(filename) {
  return downloadFileName(filename)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9._ -]/g, "")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .slice(0, 90) || "notion-page";
}

function mmToTwip(value) {
  return Math.round(Number(value || 0) * 56.6929133858);
}

function paperSettings(settings = {}) {
  const paper = settings.paper === "Letter"
    ? { width: 215.9, height: 279.4 }
    : { width: 210, height: 297 };
  if (settings.orientation === "landscape") {
    return { width: paper.height, height: paper.width, orientation: "landscape" };
  }
  return { ...paper, orientation: "portrait" };
}

function normalizeWordSettings(settings = {}) {
  return {
    paper: settings.paper === "Letter" ? "Letter" : "A4",
    orientation: settings.orientation === "landscape" ? "landscape" : "portrait",
    marginTop: Number(settings.marginTop || 18),
    marginRight: Number(settings.marginRight || 16),
    marginBottom: Number(settings.marginBottom || 18),
    marginLeft: Number(settings.marginLeft || 16),
    showHeader: settings.showHeader !== false,
    showFooter: settings.showFooter !== false,
    showProperties: settings.showProperties !== false,
    startEachPageOnNewSheet: settings.startEachPageOnNewSheet !== false,
    exportMode: settings.exportMode === "separate" ? "separate" : "single"
  };
}

async function createDocxBuffer({ bundles, settings, hiddenProperties, mermaidAssets, title }) {
  const wordSettings = normalizeWordSettings(settings);
  const paper = paperSettings(wordSettings);
  const children = [];
  const hidden = new Set(Array.isArray(hiddenProperties) ? hiddenProperties : []);

  for (let index = 0; index < bundles.length; index += 1) {
    const bundle = bundles[index];
    if (index > 0 && wordSettings.startEachPageOnNewSheet) {
      children.push(new Paragraph({ children: [new PageBreak()] }));
    }
    children.push(...bundleToDocxChildren(bundle, { settings: wordSettings, hidden, mermaidAssets }));
  }

  const doc = new Document({
    creator: "Notion PDF Studio",
    title: title || "Exportación Notion",
    description: "Documento exportado desde Notion PDF Studio",
    styles: {
      default: {
        document: {
          run: { font: "Arial", size: 22, color: "25231F" },
          paragraph: { spacing: { after: 120 } }
        }
      },
      paragraphStyles: [
        { id: "NotionTitle", name: "Notion Title", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 44, bold: true, color: "25231F" }, paragraph: { spacing: { before: 120, after: 220 } } },
        { id: "NotionCode", name: "Notion Code", basedOn: "Normal", run: { font: "Consolas", size: 19, color: "25231F" }, paragraph: { shading: { type: ShadingType.CLEAR, fill: "F4F3F0" }, spacing: { before: 80, after: 140 } } }
      ]
    },
    sections: [{
      properties: {
        page: {
          size: {
            width: mmToTwip(paper.width),
            height: mmToTwip(paper.height),
            orientation: paper.orientation
          },
          margin: {
            top: mmToTwip(wordSettings.marginTop),
            right: mmToTwip(wordSettings.marginRight),
            bottom: mmToTwip(wordSettings.marginBottom),
            left: mmToTwip(wordSettings.marginLeft)
          }
        }
      },
      headers: wordSettings.showHeader ? {
        default: new Header({
          children: [new Paragraph({
            children: [new TextRun({ text: title || "Notion PDF Studio", color: "7B756C", size: 18 })],
            alignment: AlignmentType.RIGHT
          })]
        })
      } : undefined,
      footers: wordSettings.showFooter ? {
        default: new Footer({
          children: [new Paragraph({
            children: [
              new TextRun({ text: "Notion PDF Studio · ", color: "7B756C", size: 18 }),
              new TextRun({ children: [PageNumber.CURRENT], color: "7B756C", size: 18 })
            ],
            alignment: AlignmentType.RIGHT
          })]
        })
      } : undefined,
      children
    }]
  });
  return Packer.toBuffer(doc);
}

function bundleToDocxChildren(bundle, context) {
  const children = [
    new Paragraph({
      style: "NotionTitle",
      children: [new TextRun({ text: bundle.title || "Sin título", bold: true })]
    })
  ];
  if (context.settings.showProperties) {
    children.push(...propertiesToDocx(bundle.properties || {}, context.hidden));
  }
  for (const block of flattenNotionBlocks(bundle.blocks || [])) {
    children.push(...blockToDocx(block, context));
  }
  return children;
}

function propertiesToDocx(properties, hidden) {
  const rows = Object.entries(properties)
    .filter(([name, value]) => name !== "Name" && !hidden.has(name) && value !== "" && value != null)
    .slice(0, 14)
    .map(([name, value]) => new TableRow({
      children: [
        new TableCell({
          width: { size: 28, type: WidthType.PERCENTAGE },
          shading: { type: ShadingType.CLEAR, fill: "F8F7F4" },
          children: [new Paragraph({ children: [new TextRun({ text: name, bold: true, color: "6B665F" })] })]
        }),
        new TableCell({
          width: { size: 72, type: WidthType.PERCENTAGE },
          children: [new Paragraph({ children: [new TextRun({ text: propertyValueText(value) || "Vacío" })] })]
        })
      ]
    }));
  if (!rows.length) return [];
  return [
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      layout: TableLayoutType.AUTOFIT,
      borders: softTableBorders(),
      rows
    }),
    new Paragraph({ text: "" })
  ];
}

function flattenNotionBlocks(blocks, depth = 0) {
  const out = [];
  let numberedIndex = 0;
  for (const block of blocks) {
    if (block.type === "numbered_list_item") numberedIndex += 1;
    else numberedIndex = 0;
    out.push({ ...block, depth, listIndex: numberedIndex || undefined });
    if (block.children?.length && block.type !== "table") {
      out.push(...flattenNotionBlocks(block.children, depth + 1));
    }
  }
  return out;
}

function blockToDocx(block, context) {
  const type = block.type || "unsupported";
  const indent = { left: Math.min(Number(block.depth || 0) * 360, 1440) };
  if (["heading_1", "heading_2", "heading_3"].includes(type)) {
    const heading = type === "heading_1" ? HeadingLevel.HEADING_1 : type === "heading_2" ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3;
    return [new Paragraph({ heading, indent, children: richTextToDocx(block[type]?.rich_text || []) })];
  }
  if (type === "heading_4") {
    return [new Paragraph({ indent, children: [new TextRun({ text: richTextToPlain(block.heading_4?.rich_text || []), bold: true, size: 24 })] })];
  }
  if (type === "paragraph") {
    return [new Paragraph({ indent, children: richTextToDocx(block.paragraph?.rich_text || []) || [new TextRun("")] })];
  }
  if (type === "bulleted_list_item") {
    return [new Paragraph({ indent: { ...indent, hanging: 220 }, children: [new TextRun({ text: "•  " }), ...richTextToDocx(block.bulleted_list_item?.rich_text || [])] })];
  }
  if (type === "numbered_list_item") {
    return [new Paragraph({ indent: { ...indent, hanging: 260 }, children: [new TextRun({ text: `${block.listIndex || 1}.  ` }), ...richTextToDocx(block.numbered_list_item?.rich_text || [])] })];
  }
  if (type === "to_do") {
    return [new Paragraph({ indent, children: [new TextRun({ text: block.to_do?.checked ? "☑  " : "☐  " }), ...richTextToDocx(block.to_do?.rich_text || [])] })];
  }
  if (type === "quote") {
    return [new Paragraph({
      indent: { left: (indent.left || 0) + 360 },
      border: { left: { style: BorderStyle.SINGLE, color: "CFC8BC", size: 12, space: 16 } },
      children: richTextToDocx(block.quote?.rich_text || [])
    })];
  }
  if (type === "callout") {
    return [new Paragraph({
      indent,
      shading: { type: ShadingType.CLEAR, fill: "F4F3F0" },
      spacing: { before: 100, after: 140 },
      children: [new TextRun({ text: `${block.callout?.icon?.emoji || "ℹ"}  `, bold: true }), ...richTextToDocx(block.callout?.rich_text || [])]
    })];
  }
  if (type === "divider") {
    return [new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, color: "DED8CE", size: 6, space: 8 } }, children: [new TextRun("")] })];
  }
  if (type === "code") return codeBlockToDocx(block, context, indent);
  if (type === "equation") return [equationParagraph(block.equation?.expression || "", true, indent)];
  if (type === "table") return [tableToDocx(block)];
  if (["image", "video", "audio", "pdf", "file", "bookmark", "embed", "link_preview"].includes(type)) {
    return mediaToDocx(block, indent);
  }
  if (type === "child_page" || type === "child_database") {
    return [new Paragraph({ indent, children: [new TextRun({ text: `↳ ${block[type]?.title || type.replaceAll("_", " ")}`, color: "6B665F" })] })];
  }
  return [new Paragraph({ indent, children: [new TextRun({ text: `Bloque de Notion no soportado: ${type}`, italics: true, color: "8A8176" })] })];
}

function codeBlockToDocx(block, context, indent) {
  const text = richTextToPlain(block.code?.rich_text || []);
  if (String(block.code?.language || "").toLowerCase() === "mermaid") {
    const asset = context.mermaidAssets?.[text];
    if (asset?.data) {
      const data = dataUrlToBuffer(asset.data);
      if (data) {
        return [new Paragraph({
          indent,
          alignment: AlignmentType.CENTER,
          children: [new ImageRun({
            data,
            transformation: imageDimensions(asset.width, asset.height, 520, 320)
          })]
        })];
      }
    }
  }
  return [new Paragraph({
    style: "NotionCode",
    indent,
    children: [new TextRun({ text, font: "Consolas", size: 19 })]
  })];
}

function equationParagraph(expression, display, indent) {
  const math = latexToOmmlComponent(expression);
  if (math) {
    return new Paragraph({ indent, alignment: display ? AlignmentType.CENTER : undefined, children: [math] });
  }
  return new Paragraph({
    indent,
    children: [
      new TextRun({ text: "Ecuación no convertible: ", italics: true, color: "8A8176" }),
      new TextRun({ text: String(expression || ""), font: "Consolas" })
    ]
  });
}

function latexToOmmlComponent(expression) {
  const source = String(expression || "").trim();
  if (!source) return null;
  try {
    const mathmlHtml = katex.renderToString(source, {
      output: "mathml",
      displayMode: true,
      throwOnError: false,
      strict: "ignore"
    });
    const mathml = mathmlHtml.match(/<math[\s\S]*<\/math>/)?.[0]
      ?.replace(/<annotation[\s\S]*?<\/annotation>/g, "");
    if (!mathml) return null;
    return ImportedXmlComponent.fromXmlString(mml2omml(mathml));
  } catch {
    return null;
  }
}

function tableToDocx(block) {
  const rows = block.children || [];
  if (!rows.length) {
    return new Paragraph({ children: [new TextRun({ text: "Tabla vacía", italics: true, color: "8A8176" })] });
  }
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.AUTOFIT,
    borders: softTableBorders(),
    rows: rows.map((row, rowIndex) => new TableRow({
      tableHeader: rowIndex === 0 && Boolean(block.table?.has_column_header),
      children: (row.table_row?.cells || []).map((cell) => new TableCell({
        verticalAlign: VerticalAlign.TOP,
        shading: rowIndex === 0 && block.table?.has_column_header ? { type: ShadingType.CLEAR, fill: "F4F3F0" } : undefined,
        children: [new Paragraph({ children: richTextToDocx(cell) })]
      }))
    }))
  });
}

function softTableBorders() {
  const border = { style: BorderStyle.SINGLE, size: 4, color: "DED8CE" };
  return { top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border };
}

function mediaToDocx(block, indent) {
  const payload = block[block.type] || {};
  const file = payload.type === "external" ? payload.external : payload.file;
  const url = file?.url || payload.url || "";
  const label = payload.caption?.length ? richTextToPlain(payload.caption) : block.type.toUpperCase();
  const children = [new TextRun({ text: `${label}: `, bold: true })];
  if (url) {
    children.push(new ExternalHyperlink({ link: url, children: [new TextRun({ text: url, color: "1E6B8F", underline: {} })] }));
  } else {
    children.push(new TextRun({ text: "Archivo de Notion", color: "6B665F" }));
  }
  return [new Paragraph({ indent, children })];
}

function richTextToDocx(richText) {
  const children = [];
  for (const part of richText || []) {
    if (part.type === "equation") {
      const math = latexToOmmlComponent(part.equation?.expression || part.plain_text || "");
      children.push(math || new TextRun({ text: part.equation?.expression || part.plain_text || "", font: "Consolas" }));
      continue;
    }
    const runs = plainTextRuns(part.plain_text || part.text?.content || part.mention?.plain_text || "", part.annotations || {});
    if (part.href) {
      children.push(new ExternalHyperlink({
        link: part.href,
        children: plainTextRuns(part.plain_text || part.text?.content || part.mention?.plain_text || "", {
          ...(part.annotations || {}),
          linkStyle: true
        })
      }));
    } else {
      children.push(...runs);
    }
  }
  return children.length ? children : [new TextRun("")];
}

function plainTextRuns(text, annotations = {}) {
  const pieces = String(text || "").split(/\r?\n/);
  return pieces.flatMap((piece, index) => {
    const run = new TextRun({
      text: piece,
      break: index ? 1 : undefined,
      bold: Boolean(annotations.bold),
      italics: Boolean(annotations.italic),
      strike: Boolean(annotations.strikethrough),
      color: annotations.linkStyle ? "1E6B8F" : undefined,
      underline: annotations.linkStyle || annotations.underline ? {} : undefined,
      font: annotations.code ? "Consolas" : undefined,
      shading: annotations.code ? { type: ShadingType.CLEAR, fill: "EFEDEA" } : undefined
    });
    return [run];
  });
}

function propertyValueText(value) {
  if (value?.kind === "relation") return (value.items || []).map((item) => item.title || item.id).join(", ");
  if (Array.isArray(value)) return value.map(propertyValueText).filter(Boolean).join(", ");
  if (typeof value === "boolean") return value ? "Sí" : "No";
  if (value && typeof value === "object") return value.name || value.title || value.id || Object.values(value).map(propertyValueText).filter(Boolean).join(", ");
  return String(value ?? "");
}

function dataUrlToBuffer(dataUrl) {
  const match = String(dataUrl || "").match(/^data:image\/(?:png|jpeg);base64,(.+)$/);
  return match ? Buffer.from(match[1], "base64") : null;
}

function imageDimensions(width, height, maxWidth, maxHeight) {
  const naturalWidth = Math.max(1, Number(width || maxWidth));
  const naturalHeight = Math.max(1, Number(height || maxHeight));
  const scale = Math.min(maxWidth / naturalWidth, maxHeight / naturalHeight, 1);
  return {
    width: Math.round(naturalWidth * scale),
    height: Math.round(naturalHeight * scale)
  };
}

async function serveStatic(req, res, pathname) {
  const target = pathname === "/" ? "/index.html" : pathname;
  const resolved = path.normalize(path.join(PUBLIC_DIR, target));
  if (!resolved.startsWith(PUBLIC_DIR)) {
    sendJson(res, 403, { error: "Forbidden" });
    return;
  }
  try {
    const data = await fs.readFile(resolved);
    res.writeHead(200, responseHeaders({
      "Content-Type": MIME_TYPES[path.extname(resolved)] || "application/octet-stream",
      "Content-Length": data.length,
      "Cache-Control": target === "/index.html"
        ? "no-cache"
        : "public, max-age=3600, must-revalidate"
    }));
    if (req.method === "HEAD") {
      res.end();
      return;
    }
    res.end(data);
  } catch {
    sendJson(res, 404, { error: "Not found" });
  }
}

async function handleApi(req, res, pathname) {
  const method = req.method || "GET";

  if (method === "GET" && pathname === "/api/state") {
    const state = await ensureState();
    state.savedDatabases = state.savedDatabases.filter((saved) => saved.sourceId !== DEMO_SOURCE_ID && saved.mode !== "demo");
    sendJson(res, 200, { state: redactState(state), defaults: defaultPrintSettings() });
    return;
  }

  if (method === "GET" && pathname === "/api/health") {
    sendJson(res, 200, {
      ok: true,
      name: "notion-pdf-studio",
      version: "0.1.0"
    });
    return;
  }

  const body = await readJson(req);
  const token = getToken(body);

  if (method === "POST" && pathname === "/api/state/token") {
    const state = await ensureState();
    state.notionToken = String(body.token || "").trim();
    await writeState(state);
    sendJson(res, 200, { ok: true, hasToken: Boolean(state.notionToken) });
    return;
  }

  if (method === "POST" && pathname === "/api/notion/schema") {
    const source = await resolveSource(token, body.sourceId);
    sendJson(res, 200, { source: redactSource(source) });
    return;
  }

  if (method === "POST" && pathname === "/api/notion/query") {
    const source = await resolveSource(token, body.sourceId);
    const pages = await queryPages(token, source, body.filters || [], body.sorts || [], Number(body.maxPages || 200));
    sendJson(res, 200, { source: redactSource(source), pages });
    return;
  }

  if (method === "POST" && pathname === "/api/notion/pages") {
    const pageIds = Array.isArray(body.pageIds) ? body.pageIds.slice(0, 50) : [];
    const pages = [];
    for (const pageId of pageIds) {
      pages.push(await fetchPageBundle(token, pageId));
    }
    sendJson(res, 200, { pages });
    return;
  }

  if (method === "POST" && pathname === "/api/export/docx") {
    const bundles = Array.isArray(body.bundles) ? body.bundles.slice(0, 80) : [];
    if (!bundles.length) throw httpError(400, "No pages were provided for DOCX export.");
    const settings = normalizeWordSettings(body.settings || {});
    const hiddenProperties = body.hiddenProperties || [];
    const mermaidAssets = body.mermaidAssets || {};
    const title = body.sourceName || bundles[0]?.title || "Notion export";

    if (settings.exportMode === "separate" && bundles.length > 1) {
      const zip = new JSZip();
      for (const bundle of bundles) {
        const buffer = await createDocxBuffer({
          bundles: [bundle],
          settings: { ...settings, exportMode: "single" },
          hiddenProperties,
          mermaidAssets,
          title: bundle.title || title
        });
        zip.file(`${slugFileName(bundle.title || bundle.id || "notion-page")}.docx`, buffer);
      }
      const zipBuffer = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
      sendBinary(res, 200, zipBuffer, MIME_TYPES[".zip"], `${slugFileName(title)}.zip`);
      return;
    }

    const buffer = await createDocxBuffer({
      bundles,
      settings,
      hiddenProperties,
      mermaidAssets,
      title
    });
    sendBinary(res, 200, buffer, MIME_TYPES[".docx"], `${slugFileName(title)}.docx`);
    return;
  }

  if (method === "POST" && pathname === "/api/state/save-db") {
    if (body.sourceId === DEMO_SOURCE_ID) {
      throw httpError(400, "Demo workspaces cannot be saved in production mode.");
    }
    const state = await ensureState();
    const source = await resolveSource(token, body.sourceId);
    const record = {
      id: source.sourceId,
      sourceId: source.sourceId,
      name: body.name || source.name,
      mode: source.mode,
      createdAt: new Date().toISOString(),
      lastSyncedAt: body.lastSyncedAt || null,
      filters: body.filters || [],
      sorts: body.sorts || [],
      dbSettings: normalizeDbSettings(body.dbSettings),
      settings: { ...defaultPrintSettings(), ...(body.settings || {}) }
    };
    state.savedDatabases = [
      record,
      ...state.savedDatabases.filter((saved) => saved.sourceId !== record.sourceId)
    ].slice(0, 20);
    await writeState(state);
    sendJson(res, 200, { state, saved: record });
    return;
  }

  if (method === "POST" && pathname === "/api/state/remove-db") {
    const state = await ensureState();
    state.savedDatabases = state.savedDatabases.filter((saved) => saved.sourceId !== body.sourceId);
    await writeState(state);
    sendJson(res, 200, { state });
    return;
  }

  if (method === "POST" && pathname === "/api/export/history") {
    const state = await ensureState();
    const entry = {
      id: `${Date.now()}`,
      at: new Date().toISOString(),
      sourceId: body.sourceId || "",
      sourceName: body.sourceName || "",
      pageCount: Number(body.pageCount || 0),
      sheetCount: Number(body.sheetCount || 0),
      settings: body.settings || defaultPrintSettings()
    };
    state.exportHistory = [entry, ...state.exportHistory].slice(0, 30);
    await writeState(state);
    sendJson(res, 200, { state, entry });
    return;
  }

  sendJson(res, 404, { error: "Unknown API route." });
}

function normalizeDbSettings(settings = {}) {
  return {
    hiddenProperties: Array.isArray(settings.hiddenProperties) ? settings.hiddenProperties.slice(0, 200) : []
  };
}

function redactSource(source) {
  return {
    id: source.id,
    sourceId: source.sourceId,
    mode: source.mode,
    name: source.name,
    schema: source.schema,
    version: source.version
  };
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url || "/", `http://${req.headers.host || "localhost"}`);
    if (url.pathname.startsWith("/api/")) {
      await handleApi(req, res, url.pathname);
      return;
    }
    await serveStatic(req, res, decodeURIComponent(url.pathname));
  } catch (error) {
    sendError(res, error);
  }
});

server.requestTimeout = 60_000;
server.headersTimeout = 65_000;

server.listen(PORT, () => {
  console.log(`Notion PDF Studio running at http://localhost:${PORT}`);
});

for (const signal of ["SIGINT", "SIGTERM"]) {
  process.on(signal, () => {
    server.close(() => process.exit(0));
  });
}
