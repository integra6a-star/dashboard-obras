/* Automatiza a leitura do Relatório de Monitoramento do SolcadGIS.
 *
 * Primeiro uso: o navegador abre, entre no SolcadGIS manualmente e deixe chegar
 * na tela principal. A sessão fica salva em .solcadgis-profile/ neste computador.
 * Depois disso, atualizar_monitoramento.bat reaproveita a sessão e gera
 * monitoramento_topografico.json sem publicar senha/cookies.
 */

const fs = require("fs");
const path = require("path");

let chromium;
try {
  const runtimeModules = 'C:\\Users\\micro\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\node\\node_modules';
  const runtimeExtraModules = [
    runtimeModules,
    path.join(runtimeModules, '.pnpm', 'node_modules'),
    path.join(runtimeModules, '.pnpm', 'playwright@1.61.1', 'node_modules'),
  ];
  const currentNodePath = process.env.NODE_PATH ? process.env.NODE_PATH.split(path.delimiter) : [];
  process.env.NODE_PATH = [...new Set([...currentNodePath, ...runtimeExtraModules])].join(path.delimiter);
  require("module").Module._initPaths();
  try {
    ({ chromium } = require("playwright"));
  } catch (_) {
    ({ chromium } = require(path.join(runtimeModules, "playwright")));
  }
} catch (error) {
  console.error("Playwright não encontrado no runtime do Codex.");
  console.error(error.message || error);
  process.exit(1);
}
const ROOT = path.resolve(__dirname, "..");
const PROFILE_DIR = path.join(ROOT, ".solcadgis-profile");
const DOWNLOAD_DIR = path.join(ROOT, ".solcadgis-downloads");
const OUT_JSON = path.join(ROOT, "monitoramento_topografico.json");
const DOCS_JSON = path.join(ROOT, "docs", "monitoramento_topografico.json");
const RAW_JSON = path.join(ROOT, "monitoramento_raw.json");
const SOLCAD_URL = "https://www.solcadgis.com/#/";

function nowIso() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

function normalize(text) {
  return String(text || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function parseMm(text) {
  const matches = String(text || "").match(/-?\d+(?:[,.]\d+)?\s*mm/gi) || [];
  const values = matches
    .map((m) => Number(m.replace(/mm/i, "").trim().replace(",", ".")))
    .filter(Number.isFinite);
  if (!values.length) return 0;
  return values.reduce((max, v) => Math.abs(v) > Math.abs(max) ? v : max, 0);
}

function classify(text, maiorVariacao) {
  const n = normalize(text);
  if (n.includes("ALERTA") || Math.abs(maiorVariacao) > 5) return "alerta";
  if (n.includes("ATENCAO") || n.includes("POSSIVEL") || Math.abs(maiorVariacao) >= 3) return "atencao";
  if (n.includes("ESTAVEL") || n.includes("ESTÁVEL")) return "estavel";
  return "sem_classificacao";
}

function hasAccessDenied(rawTexts) {
  const text = normalize((rawTexts || []).join("\n"));
  return text.includes("ACESSO NAO AUTORIZADO") || text.includes("VOCE NAO TEM PERMISSAO");
}

function buildAccessDeniedPayload(rawTexts, source) {
  return {
    atualizado_em: nowIso(),
    fonte: source || "SolcadGIS - Relatorio de Monitoramento",
    status_geral: "sem_acesso",
    resumo: {
      total_pocos: 0,
      estaveis: 0,
      atencao: 0,
      alerta: 0,
      sem_medicao_recente: 0,
      maior_variacao_mm: 0,
      ultima_medicao: null,
    },
    alertas: [{
      nivel: "atencao",
      mensagem: "SolcadGIS retornou acesso nao autorizado para o Relatorio de Monitoramento. Solicite a liberacao dessa permissao no perfil do usuario.",
    }],
    pocos: [],
    bruto_amostra: (rawTexts || [])
      .filter((text) => !normalize(text).includes("CLEULTON"))
      .slice(0, 12),
  };
}

function buildPayload(items, rawTexts, source) {
  const pocos = items.map((item) => {
    const maiorVariacao = parseMm(item.texto);
    const status = classify(item.texto, maiorVariacao);
    return {
      obra: item.obra || null,
      poco: item.poco || "Poço sem identificação",
      status,
      maior_variacao_mm: maiorVariacao,
      ultima_medicao: item.ultima_medicao || null,
      medicoes: item.medicoes || null,
      detalhe: item.texto.slice(0, 600),
    };
  });

  const resumo = {
    total_pocos: pocos.length,
    estaveis: pocos.filter((p) => p.status === "estavel").length,
    atencao: pocos.filter((p) => p.status === "atencao").length,
    alerta: pocos.filter((p) => p.status === "alerta").length,
    sem_medicao_recente: pocos.filter((p) => normalize(p.detalhe).includes("SEM MEDICAO") || normalize(p.detalhe).includes("NENHUMA MEDICAO")).length,
    maior_variacao_mm: pocos.reduce((max, p) => Math.max(max, Math.abs(Number(p.maior_variacao_mm || 0))), 0),
    ultima_medicao: null,
  };

  const alertas = [];
  pocos.filter((p) => p.status === "alerta").forEach((p) => {
    alertas.push({ nivel: "alerta", mensagem: `${p.poco}: variação ${p.maior_variacao_mm.toLocaleString("pt-BR")} mm` });
  });
  pocos.filter((p) => p.status === "atencao").forEach((p) => {
    alertas.push({ nivel: "atencao", mensagem: `${p.poco}: acompanhar variação ${p.maior_variacao_mm.toLocaleString("pt-BR")} mm` });
  });
  if (!pocos.length) {
    alertas.push({
      nivel: "info",
      mensagem: "Automação acessou o SolcadGIS, mas ainda não conseguiu identificar poços no relatório. Verifique se a sessão está logada e se o relatório abriu corretamente.",
    });
  }

  return {
    atualizado_em: nowIso(),
    fonte: source || "SolcadGIS - Relatório de Monitoramento",
    status_geral: resumo.alerta ? "alerta" : resumo.atencao ? "atencao" : pocos.length ? "estavel" : "sem_dados",
    resumo,
    alertas,
    pocos,
    bruto_amostra: rawTexts.slice(0, 20),
  };
}

async function clickByText(page, labels, timeout = 2500) {
  for (const label of labels) {
    const candidates = [
      page.getByText(label, { exact: true }),
      page.getByText(label),
      page.locator(`text=${label}`),
    ];
    for (const locator of candidates) {
      try {
        const first = locator.first();
        await first.waitFor({ state: "visible", timeout });
        await first.click();
        return true;
      } catch (_) {
        // tenta próximo seletor/texto
      }
    }
  }
  return false;
}

async function waitForUserLoginIfNeeded(page) {
  const bodyText = normalize(await page.locator("body").innerText().catch(() => ""));
  if (bodyText.includes("USUARIO") || bodyText.includes("SENHA") || bodyText.includes("ENTRAR")) {
    console.log("Entre no SolcadGIS na janela aberta. Aguardando chegar à tela principal...");
    await page.waitForFunction(() => {
      const text = document.body.innerText.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
      return text.includes("MEDICAO TOPOGRAFICA") || text.includes("RELATORIO DE MONITORAMENTO") || text.includes("SOLCADVIEW");
    }, { timeout: 180000 });
  }
}

async function openMonitoringReport(page) {
  await page.goto(SOLCAD_URL, { waitUntil: "domcontentloaded", timeout: 60000 });
  await page.waitForTimeout(2500);
  await waitForUserLoginIfNeeded(page);

  const mainText = normalize(await page.locator("body").innerText().catch(() => ""));
  if (!mainText.includes("RELATORIO DE MONITORAMENTO")) {
    await clickByText(page, ["MEDIÇÃO TOPOGRÁFICA", "Medição Topográfica", "MEDICAO TOPOGRAFICA"], 4000);
    await page.waitForTimeout(1200);
  }
  await clickByText(page, ["Relatório de Monitoramento", "RELATÓRIO DE MONITORAMENTO", "Relatorio de Monitoramento"], 5000);
  await page.waitForLoadState("domcontentloaded").catch(() => {});
  await page.waitForTimeout(1200);

  const reportText = normalize(await page.locator("body").innerText().catch(() => ""));
  if (reportText.includes("MONITORAMENTO DE POCO")) {
    await clickByText(page, ["Monitoramento de Poço", "MONITORAMENTO DE POÇO", "Monitoramento de Poco"], 5000);
    await page.waitForLoadState("domcontentloaded").catch(() => {});
  }
  await page.waitForTimeout(2500);
}

async function collectFromPage(page) {
  const rawTexts = await page.locator("body *").evaluateAll((nodes) => {
    return nodes
      .map((node) => (node.innerText || "").trim())
      .filter((text) => text && text.length > 2 && text.length < 1200)
      .slice(0, 600);
  }).catch(() => []);

  if (hasAccessDenied(rawTexts)) {
    return { items: [], rawTexts, selects: [], accessDenied: true };
  }

  const selects = await page.locator("select").evaluateAll((nodes) => {
    return nodes.map((select) => ({
      label: select.getAttribute("aria-label") || select.name || "",
      options: Array.from(select.options || []).map((opt) => ({ text: opt.textContent.trim(), value: opt.value })),
    }));
  }).catch(() => []);

  const pocoOptions = [];
  selects.forEach((select) => {
    select.options.forEach((opt) => {
      if (/PV|PVE|PIE|PO[CÇ]O|POSTE/i.test(opt.text) && !/SELECIONE/i.test(opt.text)) {
        pocoOptions.push(opt);
      }
    });
  });

  const items = [];
  const seen = new Set();

  for (const option of pocoOptions.slice(0, 80)) {
    const name = option.text.trim();
    if (!name || seen.has(name)) continue;
    seen.add(name);
    try {
      const select = page.locator("select").filter({ hasText: name }).first();
      await select.selectOption({ label: name }).catch(async () => select.selectOption(option.value));
      await page.waitForTimeout(900);
      const text = await page.locator("body").innerText({ timeout: 5000 });
      items.push({ poco: name, texto: text });
    } catch (_) {
      items.push({ poco: name, texto: rawTexts.join("\n") });
    }
  }

  if (!items.length) {
    const body = await page.locator("body").innerText().catch(() => "");
    const blocks = body.split(/\n{2,}/).map((x) => x.trim()).filter(Boolean);
    blocks.forEach((block) => {
      if (normalize(block).includes("MONITORAMENTO DE POCO")) return;
      const m = block.match(/\b(POSTE\s+PV-\d+|PV-\d+(?:\.\d+)?|PVE-\d+|PIE-\d+|PO[ÇC]O\s+[A-Z0-9.-]+)/i);
      if (m) items.push({ poco: m[1].trim(), texto: block });
    });
  }

  return { items, rawTexts, selects };
}

async function main() {
  fs.mkdirSync(DOWNLOAD_DIR, { recursive: true });
  const edgePath = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe";
  const chromePath = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe";
  const executablePath = fs.existsSync(edgePath) ? edgePath : fs.existsSync(chromePath) ? chromePath : undefined;
  const context = await chromium.launchPersistentContext(PROFILE_DIR, {
    headless: false,
    executablePath,
    acceptDownloads: true,
    downloadsPath: DOWNLOAD_DIR,
    viewport: { width: 1440, height: 900 },
  });
  const page = context.pages()[0] || await context.newPage();

  try {
    await openMonitoringReport(page);
    const collected = await collectFromPage(page);
    const payload = collected.accessDenied
      ? buildAccessDeniedPayload(collected.rawTexts, "SolcadGIS - Relatorio de Monitoramento")
      : buildPayload(collected.items, collected.rawTexts, "SolcadGIS - Relatório de Monitoramento");
    fs.writeFileSync(OUT_JSON, JSON.stringify(payload, null, 2), "utf8");
    fs.writeFileSync(DOCS_JSON, JSON.stringify(payload, null, 2), "utf8");
    fs.writeFileSync(RAW_JSON, JSON.stringify({ coletado_em: nowIso(), pagina: collected }, null, 2), "utf8");
    console.log(`Monitoramento salvo: ${payload.resumo.total_pocos} poço(s), ${payload.alertas.length} alerta(s).`);
  } finally {
    await context.close();
  }
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});

