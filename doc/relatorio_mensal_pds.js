(() => {
  const mesesNome = {
    "01": "Janeiro", "02": "Fevereiro", "03": "Março", "04": "Abril",
    "05": "Maio", "06": "Junho", "07": "Julho", "08": "Agosto",
    "09": "Setembro", "10": "Outubro", "11": "Novembro", "12": "Dezembro"
  };

  let PDS = [];
  let MAPA = { trechos: [] };
  let ENDERECOS = {};
  const HOJE = new Date();

  const obraMap = [
    [/MAR DE CORAL/i, "Coletor Tronco Secundário Mar de Coral", ["cts_mar_de_coral"]],
    [/AGUA VERMELHA|ÁGUA VERMELHA/i, "Coletor Tronco Secundário Água Vermelha ME", ["agua"]],
    [/AGUA VERMELHA.*M\.?D|ÁGUA VERMELHA.*M\.?D|M\.D/i, "Coletor Tronco Secundário Água Vermelha MD", ["agua_md"]],
    [/ARARIBA|ARARIBÁ/i, "Rede Coletora de Esgoto Barão Luís de Araribá", ["rce_barao_luis_de_arariba"]],
    [/LOURDES/i, "Coletor Tronco Secundário Jardim Lourdes", ["cts_jardim_lourdes"]],
    [/CONCEI/i, "Coletor Tronco Secundário Conceição", ["conceicao"]],
    [/LAGEADO/i, "Coletor Tronco Lageado", ["lageado_montante_bl54_55_56"]],
    [/LUIS BOTELHO|LUÍS BOTELHO/i, "Coletor Tronco Secundário Luís Botelho Mourão", ["luis_botelho"]],
    [/AGUA BOA|ÁGUA BOA/i, "Interligações Água Boa", ["agua_boa"]],
    [/PACHECO/i, "Interligação Fernando Pacheco Jordão", ["pacheco"]],
    [/SOUTO/i, "Rede coletora Francisco Souto Maior", ["rce_souto_maior"]]
  ];

  const ordemObras = [
    "Coletor Tronco Secundário Água Vermelha ME",
    "Coletor Tronco Secundário Água Vermelha MD",
    "Coletor Tronco Secundário Mar de Coral",
    "Coletor Tronco Secundário Luís Botelho Mourão",
    "Coletor Tronco Secundário Conceição",
    "Interligações Água Boa",
    "Interligação Fernando Pacheco Jordão",
    "Rede coletora Francisco Souto Maior",
    "Rede Coletora de Esgoto Barão Luís de Araribá",
    "Coletor Tronco Secundário Jardim Lourdes",
    "Coletor Tronco Lageado"
  ];

  function esc(s) {
    return String(s ?? "").replace(/[&<>"']/g, m => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
    }[m]));
  }

  function tituloMes(ym) {
    const [ano, mes] = String(ym).split("-");
    return `${mesesNome[mes] || mes} de ${ano}`;
  }

  function ymFromDate(value) {
    const text = String(value || "").trim();
    const m = text.match(/^(\d{4})-(\d{2})/);
    return m ? `${m[1]}-${m[2]}` : "";
  }

  function parseDateLocal(value) {
    const text = String(value || "").trim();
    const m = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (!m) return null;
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }

  function ymHoje() {
    const y = HOJE.getFullYear();
    const m = String(HOJE.getMonth() + 1).padStart(2, "0");
    return `${y}-${m}`;
  }

  function dentroDoPeriodo(row, ym) {
    if (ymFromDate(row.data) !== ym) return false;
    if (ym !== ymHoje()) return true;
    const data = parseDateLocal(row.data);
    if (!data) return false;
    return data <= new Date(HOJE.getFullYear(), HOJE.getMonth(), HOJE.getDate(), 23, 59, 59);
  }

  function normalizeText(value) {
    return String(value || "")
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
  }

  function pvNorm(value) {
    const text = String(value || "").replace(",", ".").toUpperCase();
    const match = text.match(/\d+(?:\.\d+)?/);
    return match ? match[0].replace(/^0+(?=\d)/, "") : "";
  }

  function pvList(value) {
    const text = String(value || "").replace(/,/g, ".").toUpperCase();
    const matches = text.match(/\d+(?:\.\d+)?/g) || [];
    return [...new Set(matches.map(v => v.replace(/^0+(?=\d)/, "")).filter(Boolean))];
  }

  function canonicalObra(raw) {
    const source = String(raw || "");
    const n = normalizeText(source);
    if (!n || n === "VACAL" || n.includes("GUINDAUTO")) return "";
    const md = /AGUA VERMELHA/.test(n) && (/\bM D\b/.test(n) || /\bMD\b/.test(n));
    if (md) return "Coletor Tronco Secundário Água Vermelha MD";
    for (const [regex, name] of obraMap) {
      if (regex.test(source)) return name;
    }
    return source.trim();
  }

  function obraIds(obra) {
    const found = obraMap.find(([, name]) => name === obra);
    return found ? found[2] : [];
  }

  function pvLabel(obra, pv) {
    if (obra.includes("Água Vermelha ME")) return `PVE ${pv}`;
    if (obra.includes("Água Vermelha MD")) return `PVD ${pv}`;
    return `PV ${pv}`;
  }

  function metodoTexto(value) {
    const n = normalizeText(value);
    if (n.includes("MND") || n.includes("NAO DESTRUTIVO")) return "Não Destrutiva";
    if (n.includes("VCA") || n.includes("VALA")) return "Vala a Céu Aberto";
    if (n.includes("CRAV")) return "Tubo Cravado";
    return "metodologia aplicável";
  }

  function inferMetodo(row, mapaInfo) {
    const source = `${row.tipo || ""} ${row.atividade || ""} ${mapaInfo?.metodo || ""}`;
    return metodoTexto(source);
  }

  function inferTipo(atividade) {
    const n = normalizeText(atividade);
    if (n.includes("SHAFT")) return "Shaft";
    if (n.includes("TRANSFORM")) return "Transformação";
    if (n.includes("APOIO HDD")) return "Apoio HDD";
    if (n.includes("FURO")) return "Furo piloto";
    if (n.includes("PILOTO") && n.includes("PUXE")) return "Piloto e puxe";
    if (n.includes("PUXE")) return "Puxe";
    if (n.includes("VCA")) return "VCA";
    if (n.includes("INTERLIG")) return "Interligação";
    return "";
  }

  function cacheKey(lat, lon) {
    return `${Number(lat).toFixed(6)},${Number(lon).toFixed(6)}`;
  }

  function mapaPorObraPv(obra, pv) {
    const ids = obraIds(obra);
    for (const id of ids) {
      for (const t of MAPA.trechos || []) {
        if (String(t.obra_id || "") !== id) continue;
        const ini = pvNorm(t.pv_inicio);
        const fim = pvNorm(t.pv_fim);
        if (ini === pv && t.lat_inicio && t.lon_inicio) return { ...t, lat: t.lat_inicio, lon: t.lon_inicio };
        if (fim === pv && t.lat_fim && t.lon_fim) return { ...t, lat: t.lat_fim, lon: t.lon_fim };
      }
    }
    return null;
  }

  function enderecoDoMapa(info) {
    if (!info) {
      return {
        address: "endereço a validar, nº a validar, bairro a validar, São Paulo",
        road: "endereço a validar",
        suburb: "bairro a validar",
        city: "São Paulo"
      };
    }

    const cached = ENDERECOS[cacheKey(info.lat, info.lon)] || {};
    let road = cached.road || "endereço a validar";
    let number = cached.number || "a validar";
    let suburb = cached.suburb || cached.neighbourhood || "bairro a validar";
    let city = cached.city || cached.town || cached.municipality || "São Paulo";
    let state = cached.state || "São Paulo";

    if (/logradouro não localizado/i.test(road)) road = "endereço a validar";
    if (/bairro não localizado/i.test(suburb)) suburb = "bairro a validar";

    const numero = String(number).toLowerCase() === "s/n" ? "s/n" : `nº ${number}`;
    return {
      address: `${road}, ${numero}, ${suburb}, ${city}, ${state}`.replace(/, São Paulo, São Paulo$/, ", São Paulo"),
      road,
      suburb,
      city
    };
  }

  function highlightText(text) {
    let out = esc(text);
    const terms = [
      "Fevereiro de 2026", "Março de 2026", "Abril de 2026", "Maio de 2026",
      "Coletores Tronco Secundários", "Água Vermelha", "Mar de Coral",
      "Barão Luís de Araribá", "Jardim Lourdes", "Conceição", "Luís Botelho",
      "Água Boa", "Pacheco Jordão", "Souto Maior", "Itaim Paulista",
      "São Paulo", "Lajeado", "Vila Nova Curuçá", "Rua", "Avenida",
      "Travessa", "Viela", "Subprefeitura", "Não Destrutiva", "Vala a Céu Aberto"
    ];
    for (const term of terms) {
      const safe = esc(term).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      out = out.replace(new RegExp(`(${safe})`, "gi"), '<span class="hl">$1</span>');
    }
    return out;
  }

  function linhasMes(ym) {
    return PDS.flatMap(r => {
      if (!dentroDoPeriodo(r, ym)) return [];
      const obra = canonicalObra(r.obra || "");
      if (!obra) return [];
      const pvs = pvList(r.pv || r.PV_Inicio || r.atividade);
      return pvs.map(pv => ({
        data: r.data,
        ym: ymFromDate(r.data),
        obraRaw: r.obra || "",
        obra,
        equipe: r.equipe || "",
        atividade: r.atividade || "",
        pv,
        tipo: r.tipo || r.Tipo_Atividade || inferTipo(r.atividade)
      }));
    }).filter(r => r.obra && r.pv);
  }

  function agruparMes(ym) {
    const grupos = new Map();
    for (const row of linhasMes(ym)) {
      const key = `${row.obra}||${row.pv}`;
      if (!grupos.has(key)) {
        const mapaInfo = mapaPorObraPv(row.obra, row.pv);
        const end = enderecoDoMapa(mapaInfo);
        grupos.set(key, {
          obra: row.obra,
          pv: row.pv,
          first: row.data,
          last: row.data,
          rows: [],
          mapaInfo,
          endereco: end,
          metodo: inferMetodo(row, mapaInfo)
        });
      }
      const g = grupos.get(key);
      g.rows.push(row);
      if (row.data < g.first) g.first = row.data;
      if (row.data > g.last) g.last = row.data;
      const metodo = inferMetodo(row, g.mapaInfo);
      if (metodo !== "metodologia aplicável") g.metodo = metodo;
    }

    const list = [...grupos.values()].sort((a, b) => {
      const oi = ordemObras.indexOf(a.obra);
      const oj = ordemObras.indexOf(b.obra);
      const ao = oi < 0 ? 999 : oi;
      const bo = oj < 0 ? 999 : oj;
      return ao - bo || a.first.localeCompare(b.first) || Number(a.pv) - Number(b.pv);
    });

    const porObra = new Map();
    for (const item of list) {
      if (!porObra.has(item.obra)) porObra.set(item.obra, []);
      porObra.get(item.obra).push(item);
    }
    return porObra;
  }

  function resumoObras(porObra) {
    return [...porObra.keys()].map(o =>
      o.replace("Coletor Tronco Secundário ", "").replace("Rede Coletora de Esgoto ", "")
    ).join(", ");
  }

  function ruasPrincipais(items) {
    const roads = [];
    for (const item of items) {
      const road = item.endereco.road;
      if (road && !/validar/i.test(road) && !roads.includes(road)) roads.push(road);
      if (roads.length >= 4) break;
    }
    return roads.length ? roads.join(", ") : "endereços a validar pela planilha de mapa";
  }

  function gerarAcompanhamento(ref, porObra) {
    const obrasTxt = resumoObras(porObra) || "frentes registradas na PDS";
    const html = [];
    html.push(`<p>${highlightText(`No mês de ${ref}, as atividades executadas concentraram-se nas frentes dos Coletores Tronco Secundários ${obrasTxt}, localizados nos municípios e subprefeituras indicados nos pontos georreferenciados, no Estado de São Paulo.`)}</p>`);
    html.push(`<p>O planejamento mensal foi elaborado com base nos cronogramas específicos de cada coletor.</p>`);
    html.push(`<p>Os serviços executados compreenderam a mobilização de equipes, a implantação de sinalização temporária, a execução de Poços de Visita (PVs), o assentamento de tubulações, bem como a recomposição provisória e definitiva das vias públicas.</p>`);

    let index = 0;
    for (const [obra, items] of porObra) {
      const primeiro = items[0];
      const ultimo = items[items.length - 1];
      const cidade = items.find(i => i.endereco.city)?.endereco.city || "São Paulo";
      const sub = items.find(i => !/validar/i.test(i.endereco.suburb))?.endereco.suburb || "bairro a validar";
      const metodo = items.find(i => i.metodo !== "metodologia aplicável")?.metodo || "metodologia aplicável";
      const prefixo = index === 0 ? "No " : "De forma semelhante, ";
      html.push(`<p>${highlightText(`${prefixo}${obra} as atividades tiveram início nas ${ruasPrincipais(items)}, no município de ${cidade}, Subprefeitura ${sub}, abrangendo o trecho entre os ${pvLabel(obra, primeiro.pv)} e ${pvLabel(obra, ultimo.pv)}, com utilização da metodologia ${metodo} como técnica principal. A ocupação das vias foi planejada de modo a preservar o acesso local e minimizar as interferências no trânsito, em conformidade com os critérios definidos no PGSV.`)}</p>`);
      index += 1;
    }
    return html.join("");
  }

  function gerarAtividades(ref, porObra) {
    const obrasTxt = resumoObras(porObra) || "frentes registradas na PDS";
    const html = [`<p>${highlightText(`Durante o mês de ${ref}, as ações operacionais concentraram-se na execução das redes coletoras e das interligações dos ${obrasTxt}.`)}</p>`];
    for (const [obra, items] of porObra) {
      html.push(`<p><b>${highlightText(obra)}</b></p>`);
      const seen = new Set();
      for (const item of items) {
        const key = `${item.pv}|${item.endereco.address}`;
        if (seen.has(key)) continue;
        seen.add(key);
        html.push(`<p><b>${esc(pvLabel(obra, item.pv))}:</b> ${highlightText(item.endereco.address)}.</p>`);
      }
      if (items.length) {
        html.push(`<p>${highlightText(`As frentes de trabalho desenvolveram-se no trecho compreendido entre os PVs ${items[0].pv} ao ${items[items.length - 1].pv}, abrangendo intervenções nos endereços relacionados acima.`)}</p>`);
        html.push(`<p>As atividades executadas adotaram a metodologia aplicável a cada trecho, conforme registros da PDS e da planilha_base_mapa.xlsx. Dentre os serviços realizados, destacam-se:</p>`);
        html.push(`<ul>
          <li>Perfuração dos pontos de entrada e saída para lançamento das tubulações por MND;</li>
          <li>Assentamento de tubulações em PVC e PEAD, em conformidade com as especificações de projeto;</li>
          <li>Execução de Poços de Visita (PVs) nos pontos de acesso e conexão;</li>
          <li><span class="hl">Recomposição provisória</span> localizada do pavimento nas áreas diretamente afetadas pelas intervenções;</li>
          <li>Implantação de sinalização provisória e dispositivos de segurança viária, em atendimento aos Projetos de Sinalização Viária e às diretrizes do PGSV.</li>
        </ul>`);
      }
    }
    return html.join("");
  }

  async function carregarPdsRelatorio() {
    try {
      const [pdsResp, mapaResp, endResp] = await Promise.all([
        fetch("./pds_data.json?cb=" + Date.now()),
        fetch("./dados_mapa.json?cb=" + Date.now()),
        fetch("./relatorio_enderecos.json?cb=" + Date.now()).catch(() => null)
      ]);
      PDS = await pdsResp.json();
      MAPA = await mapaResp.json();
      ENDERECOS = endResp && endResp.ok ? await endResp.json() : {};
    } catch (err) {
      console.error("Falha ao carregar bases do relatório", err);
      PDS = [];
      MAPA = { trechos: [] };
      ENDERECOS = {};
    }

    const mesesSet = new Set(PDS.map(r => ymFromDate(r.data)).filter(Boolean));
    mesesSet.add(ymHoje());
    const meses = [...mesesSet].sort();
    const sel = document.getElementById("mes");
    if (sel) {
      const anterior = sel.value;
      sel.innerHTML = "";
      const preferido = meses.includes(anterior) ? anterior : (meses.includes(ymHoje()) ? ymHoje() : meses[meses.length - 1]);
      meses.forEach(m => {
        const option = document.createElement("option");
        option.value = m;
        option.textContent = m === ymHoje() ? `${tituloMes(m)} (até hoje)` : tituloMes(m);
        if (m === preferido) option.selected = true;
        sel.appendChild(option);
      });
      sel.onchange = gerarRelatorioPds;
    }
    gerarRelatorioPds();
  }

  function gerarRelatorioPds() {
    const ym = document.getElementById("mes")?.value || ymHoje();
    const ref = tituloMes(ym);
    const porObra = agruparMes(ym);

    document.getElementById("cReferencia").textContent = ref.replace(" de ", " /");
    document.getElementById("cContrato").textContent = document.getElementById("contrato").value || "";
    document.getElementById("cContratada").textContent = document.getElementById("contratada").value || "Consórcio XXXXX Pacote XX";

    if (!porObra.size) {
      document.getElementById("textoAcompanhamento").innerHTML = `<p>No mês de <span class="hl">${esc(ref)}</span>, não foram localizados registros de PDS com PV informado para gerar o acompanhamento no padrão combinado.</p>`;
      document.getElementById("textoAtividades").innerHTML = `<p>Não há atividades com PV registrado para o período selecionado.</p>`;
    } else {
      document.getElementById("textoAcompanhamento").innerHTML = gerarAcompanhamento(ref, porObra);
      document.getElementById("textoAtividades").innerHTML = gerarAtividades(ref, porObra);
    }

    const totalPvs = [...porObra.values()].reduce((acc, rows) => acc + rows.length, 0);
    document.getElementById("indicadores").innerHTML = `No período, foram consolidados <b>${totalPvs}</b> PVs/frentes com registros na PDS, distribuídos em <b>${porObra.size}</b> obras ou coletores.`;
    document.getElementById("conclusao").innerHTML = `
      <p>Durante o mês de <span class="hl">${esc(ref)}</span>, as frentes de trabalho foram acompanhadas conforme registros de PDS, pontos georreferenciados da planilha de mapa e diretrizes do Plano de Gestão da Qualidade (PGQ).</p>
      <p>As atividades inspecionáveis permaneceram monitoradas e registradas, mantendo a rastreabilidade necessária para acompanhamento mensal.</p>
      <p>Não foram identificadas não conformidades que impactassem a qualidade ou a continuidade dos serviços no período, ressalvados os campos marcados como endereço a validar.</p>
    `;
  }

  window.gerarRelatorio = gerarRelatorioPds;
  window.gerarRelatorioPds = gerarRelatorioPds;

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", carregarPdsRelatorio);
  } else {
    carregarPdsRelatorio();
  }
})();
