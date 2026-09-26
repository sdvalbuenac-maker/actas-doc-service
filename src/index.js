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
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

  try {
    doc.render(sanitizeCampos(campos));
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

  return doc.getZip().generate({ type: "nodebuffer" });
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

  const nombreCaso = slugify(campos.aprehendido_nombre || campos.caso || "caso");

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
