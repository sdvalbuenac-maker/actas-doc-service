const express = require("express");
const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const archiver = require("archiver");

const PORT = process.env.PORT || 3000;
const API_KEY = process.env.API_KEY || null;
const TEMPLATES_DIR = path.join(__dirname, "..", "templates");
const MANIFEST_PATH = path.join(TEMPLATES_DIR, "manifest.json");

const app = express();
app.use(express.json({ limit: "5mb" }));

function loadManifest() {
  const raw = fs.readFileSync(MANIFEST_PATH, "utf8");
  const manifest = JSON.parse(raw);
  return manifest.documentos || {};
}

function checkApiKey(req, res, next) {
  if (!API_KEY) return next();
  const provided = req.header("x-api-key");
  if (provided !== API_KEY) {
    return res.status(401).json({ error: "API key inválida o faltante (header x-api-key)." });
  }
  next();
}

function sanitizeCampos(campos) {
  const out = {};
  for (const [key, value] of Object.entries(campos || {})) {
    out[key] = value === null || value === undefined ? "" : value;
  }
  return out;
}


// ---------------------------------------------------------------------------
// Enriquecimiento de campos (cálculos que NO deben depender del modelo de IA)
// ---------------------------------------------------------------------------
const NUM_PALABRAS = ["cero","un","dos","tres","cuatro","cinco","seis","siete","ocho","nueve","diez",
  "once","doce","trece","catorce","quince","dieciséis","diecisiete","dieciocho","diecinueve","veinte"];
const BLANK = "________________";

function cap(s) { return s.charAt(0).toUpperCase() + s.slice(1); }
function dos(n) { return String(n).padStart(2, "0"); }

function enrich(campos) {
  const c = { ...campos };

  // Campos que, si no se informan, deben quedar como espacio en blanco para completar a mano
  for (const k of ["EXPEDIENTE", "NUMERO_OFICIO", "NUM_FOLIOS", "COORDENADAS"]) {
    if (c[k] === undefined || c[k] === null || String(c[k]).trim() === "") c[k] = BLANK;
  }

  // Detenidos: texto para el oficio de remisión y para la ficha final
  const det = Array.isArray(c.DETENIDOS) ? c.DETENIDOS.filter((d) => d && typeof d === "object") : [];
  if (det.length) {
    const n = det.length;
    const palabra = n <= 20 ? NUM_PALABRAS[n] : String(n);
    const ident = det.map((d, i) => {
      const nom = d.NOMBRE_COMPLETO || BLANK;
      const ced = d.CEDULA ? `, titular de la cédula de identidad ${d.CEDULA}` : "";
      return (n > 1 ? `${i + 1}) ` : "") + nom + ced;
    });
    if (c.DETENIDOS_TEXTO === undefined || c.DETENIDOS_TEXTO === "") {
      c.DETENIDOS_TEXTO =
        n === 1
          ? `Un (01) ciudadano en calidad de detenido (${ident[0]})`
          : `${cap(palabra)} (${dos(n)}) ciudadanos en calidad de detenidos (${ident.join("; ")})`;
    }
    if (c.DETENIDOS_FICHA === undefined || c.DETENIDOS_FICHA === "") {
      c.DETENIDOS_FICHA = det
        .map((d, i) => (n > 1 ? `${i + 1}) ` : "") + String(d.NOMBRE_COMPLETO || BLANK).toUpperCase() +
          (d.CEDULA ? `, TITULAR DE LA CÉDULA DE IDENTIDAD: ${d.CEDULA}` : ""))
        .join("\n");
    }
  } else {
    // Sin lista estructurada: un solo bloque vacío para que la sección de derechos se genere una vez
    c.DETENIDOS = [{}];
    if (!c.DETENIDOS_TEXTO) c.DETENIDOS_TEXTO = "Un (01) ciudadano en calidad de detenido";
    if (!c.DETENIDOS_FICHA) c.DETENIDOS_FICHA = BLANK;
  }

  // Firmas de funcionarios actuantes en dos columnas: FUNCIONARIOS_FIRMANTES = ["OFICIAL JEFE (CPNB) MEDINA CARLOS", ...]
  if (Array.isArray(c.FUNCIONARIOS_FIRMANTES) && !Array.isArray(c.FILAS_FIRMAS)) {
    const f = c.FUNCIONARIOS_FIRMANTES.map((x) => `____________________________\n${x}`);
    const filas = [];
    for (let i = 0; i < f.length; i += 2) filas.push({ IZQ: f[i], DER: f[i + 1] || "" });
    c.FILAS_FIRMAS = filas;
  }
  if (!Array.isArray(c.FILAS_FIRMAS)) c.FILAS_FIRMAS = [{ IZQ: "____________________________", DER: "" }];

  return c;
}

// Cuando una sección con dibujos se repite (un bloque por detenido), Word necesita ids únicos en wp:docPr
function renumberDrawingIds(zip) {
  const names = Object.keys(zip.files).filter((n) => /^word\/(document|header\d*|footer\d*)\.xml$/.test(n));
  let next = 1;
  for (const name of names) {
    const xml = zip.file(name).asText();
    const out = xml.replace(/(<wp:docPr\b[^>]*?\bid=")\d+(")/g, (_, a, b) => `${a}${next++}${b}`);
    if (out !== xml) zip.file(name, out);
  }
}

function renderDocx(templateFilePath, campos) {
  if (!fs.existsSync(templateFilePath)) {
    const err = new Error(
      `Plantilla no encontrada en el servidor: ${path.basename(templateFilePath)}. Súbela a la carpeta templates/ del microservicio.`
    );
    err.code = "TEMPLATE_NOT_FOUND";
    throw err;
  }
  const content = fs.readFileSync(templateFilePath, "binary");
  const zip = new PizZip(content);
  // Las plantillas .docx reales usan llaves dobles {{campo}} (no {campo} como
  // dice el comentario original del manifest). Sin este delimiter explícito,
  // docxtemplater interpreta cada {{ como una apertura duplicada y falla el
  // render de las 19 plantillas con "Multi error" / "Duplicate open tag".
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: "{{", end: "}}" },
    nullGetter: () => "",
  });

  try {
    doc.render(enrich(sanitizeCampos(campos)));
  } catch (error) {
    const detalles =
      error.properties && error.properties.errors
        ? error.properties.errors.map((e) => e.properties && e.properties.explanation).filter(Boolean)
        : [error.message];
    const wrapped = new Error(
      `Error al rellenar la plantilla ${path.basename(templateFilePath)}: ${detalles.join("; ")}`
    );
    wrapped.code = "RENDER_ERROR";
    throw wrapped;
  }

  const outZip = doc.getZip();
  renumberDrawingIds(outZip);
  return outZip.generate({ type: "nodebuffer", compression: "DEFLATE" });
}

function slugify(text) {
  return String(text || "documento")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, "-")
    .replace(/(^-|-$)/g, "")
    .toLowerCase();
}

app.get("/health", (req, res) => {
  res.json({ status: "ok" });
});

app.get("/plantillas", checkApiKey, (req, res) => {
  const manifest = loadManifest();
  const estado = Object.entries(manifest).map(([nombre, cfg]) => ({
    documento: nombre,
    archivo: cfg.archivo,
    subida: fs.existsSync(path.join(TEMPLATES_DIR, cfg.archivo)),
  }));
  res.json({ plantillas: estado });
});

app.post("/generar", checkApiKey, async (req, res) => {
  const { documentos, campos } = req.body || {};

  if (!Array.isArray(documentos) || documentos.length === 0) {
    return res.status(400).json({ error: "Falta 'documentos' (array de nombres) en el body." });
  }
  if (!campos || typeof campos !== "object") {
    return res.status(400).json({ error: "Falta 'campos' (objeto) en el body." });
  }

  let manifest;
  try {
    manifest = loadManifest();
  } catch (e) {
    return res.status(500).json({ error: `No se pudo leer templates/manifest.json: ${e.message}` });
  }

  const generados = [];
  const errores = [];

  for (const nombreDocumento of documentos) {
    const cfg = manifest[nombreDocumento];
    if (!cfg) {
      errores.push(`"${nombreDocumento}" no está registrado en templates/manifest.json (revisa el nombre exacto).`);
      continue;
    }
    try {
      const buffer = renderDocx(path.join(TEMPLATES_DIR, cfg.archivo), campos);
      generados.push({ nombre: nombreDocumento, buffer });
    } catch (e) {
      errores.push(e.message);
    }
  }

  if (generados.length === 0) {
    return res.status(422).json({ error: "No se pudo generar ningún documento.", detalles: errores });
  }

  const nombreCaso = slugify(
    campos.EXPEDIENTE || campos.DATOS_APREHENDIDO || campos.aprehendido_nombre || campos.caso || "caso"
  );

  if (generados.length === 1 && errores.length === 0) {
    const filename = `${slugify(generados[0].nombre)}-${nombreCaso}.docx`;
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    return res.send(generados[0].buffer);
  }

  const zipFilename = `actas-${nombreCaso}.zip`;
  res.setHeader("Content-Type", "application/zip");
  res.setHeader("Content-Disposition", `attachment; filename="${zipFilename}"`);
  if (errores.length) res.setHeader("X-Generation-Warnings", encodeURIComponent(errores.join(" | ")));

  const archive = archiver("zip", { zlib: { level: 9 } });
  archive.on("error", () => { if (!res.headersSent) res.status(500); res.end(); });
  archive.pipe(res);
  for (const doc of generados) {
    archive.append(doc.buffer, { name: `${slugify(doc.nombre)}-${nombreCaso}.docx` });
  }
  archive.finalize();
});

app.listen(PORT, () => {
  console.log(`actas-doc-service escuchando en el puerto ${PORT}`);
});
    
