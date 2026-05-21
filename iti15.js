const pvs = [
  { id: "PV-01", total: 3.2, excavated: 3.2, team: "Equipe A", start: "13/05/2026" },
  { id: "PV-02", total: 3.8, excavated: 2.7, team: "Equipe A", start: "14/05/2026" },
  { id: "PV-03", total: 4.6, excavated: 1.6, team: "Equipe B", start: "15/05/2026" },
  { id: "PV-04", total: 2.9, excavated: 2.9, team: "Equipe B", start: "15/05/2026" },
  { id: "PV-05", total: 5.4, excavated: 2.1, team: "Equipe C", start: "16/05/2026" },
  { id: "PV-06", total: 4.2, excavated: 0.8, team: "Equipe C", start: "18/05/2026" },
  { id: "PV-07", total: 3.6, excavated: 0, team: "Aguardando", start: "-" },
  { id: "PV-08", total: 4.9, excavated: 1.3, team: "Equipe D", start: "19/05/2026" },
  { id: "PV-09", total: 3.4, excavated: 0, team: "Aguardando", start: "-" },
  { id: "PV-10", total: 5.8, excavated: 3.4, team: "Equipe D", start: "17/05/2026" },
  { id: "PV-11", total: 2.7, excavated: 0, team: "Aguardando", start: "-" },
  { id: "PV-12", total: 4.1, excavated: 0.6, team: "Equipe E", start: "20/05/2026" },
];

const networks = [
  { from: "PV-01", to: "PV-02", total: 38, done: 38 },
  { from: "PV-02", to: "PV-03", total: 44, done: 29 },
  { from: "PV-03", to: "PV-04", total: 27, done: 12 },
  { from: "PV-04", to: "PV-05", total: 52, done: 21 },
  { from: "PV-05", to: "PV-06", total: 35, done: 8 },
  { from: "PV-06", to: "PV-07", total: 41, done: 0 },
];

const maxDepth = 6;
let selectedPvId = "PV-05";

const pvProfile = document.querySelector("#pvProfile");
const selectedDetails = document.querySelector("#selectedDetails");
const selectedStatus = document.querySelector("#selectedStatus");
const networkBars = document.querySelector("#networkBars");

function percent(value, total) {
  if (!total) return 0;
  return Math.min(100, Math.round((value / total) * 100));
}

function statusFor(pv) {
  if (pv.excavated <= 0) return "Pendente";
  if (pv.excavated >= pv.total) return "Concluido";
  return "Em escavacao";
}

function renderPvs() {
  pvProfile.innerHTML = "";

  pvs.forEach((pv) => {
    const progress = percent(pv.excavated, pv.total);
    const height = Math.max(62, (pv.total / maxDepth) * 342);
    const card = document.createElement("article");
    card.className = `pv-card ${progress === 100 ? "complete" : ""} ${progress === 0 ? "pending" : ""} ${
      selectedPvId === pv.id ? "selected" : ""
    }`;
    card.tabIndex = 0;
    card.setAttribute("role", "button");
    card.setAttribute("aria-label", `${pv.id}, ${progress}% escavado`);
    card.innerHTML = `
      <div class="pv-label">
        <span>${pv.id}</span>
        <small>${progress}%</small>
      </div>
      <div class="shaft-wrap">
        <div class="shaft" style="height:${height}px">
          <span class="cap"></span>
          <span class="fill" style="height:${progress}%"></span>
        </div>
      </div>
      <div class="pv-foot">
        <strong>${pv.excavated.toFixed(1)} m / ${pv.total.toFixed(1)} m</strong>
        <input class="depth-input" type="number" min="0" max="${pv.total}" step="0.1" value="${pv.excavated.toFixed(
          1,
        )}" aria-label="Metros escavados ${pv.id}" />
      </div>
    `;

    card.addEventListener("click", () => {
      selectedPvId = pv.id;
      render();
    });

    card.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        selectedPvId = pv.id;
        render();
      }
    });

    card.querySelector("input").addEventListener("click", (event) => event.stopPropagation());
    card.querySelector("input").addEventListener("change", (event) => {
      pv.excavated = Math.min(pv.total, Math.max(0, Number(event.target.value) || 0));
      selectedPvId = pv.id;
      render();
    });

    pvProfile.appendChild(card);
  });
}

function renderSummary() {
  const done = pvs.filter((pv) => pv.excavated >= pv.total).length;
  const totalDepth = pvs.reduce((sum, pv) => sum + pv.total, 0);
  const excavated = pvs.reduce((sum, pv) => sum + pv.excavated, 0);

  document.querySelector("#donePvs").textContent = `${done}/${pvs.length}`;
  document.querySelector("#totalDepth").textContent = `${totalDepth.toFixed(1)} m`;
  document.querySelector("#overallProgress").textContent = `${percent(excavated, totalDepth)}%`;
}

function renderSelected() {
  const pv = pvs.find((item) => item.id === selectedPvId) || pvs[0];
  const progress = percent(pv.excavated, pv.total);
  const remaining = Math.max(0, pv.total - pv.excavated);
  const status = statusFor(pv);

  selectedStatus.textContent = status;
  selectedStatus.classList.toggle("green", status === "Concluido");

  selectedDetails.innerHTML = `
    <div class="metric"><span>PV</span><strong>${pv.id}</strong></div>
    <div class="metric"><span>Status</span><strong>${status}</strong></div>
    <div class="metric"><span>Profundidade prevista</span><strong>${pv.total.toFixed(2)} m</strong></div>
    <div class="metric"><span>Escavado</span><strong>${pv.excavated.toFixed(2)} m</strong></div>
    <div class="metric"><span>Falta escavar</span><strong>${remaining.toFixed(2)} m</strong></div>
    <div class="metric"><span>Avanco</span><strong>${progress}%</strong></div>
    <div class="metric"><span>Equipe</span><strong>${pv.team}</strong></div>
    <div class="metric"><span>Inicio</span><strong>${pv.start}</strong></div>
  `;
}

function renderNetworks() {
  const total = networks.reduce((sum, item) => sum + item.total, 0);
  const done = networks.reduce((sum, item) => sum + item.done, 0);
  document.querySelector("#networkPercent").textContent = `${percent(done, total)}%`;

  networkBars.innerHTML = networks
    .map((item) => {
      const progress = percent(item.done, item.total);
      return `
        <div class="network-row">
          <strong>${item.from} - ${item.to}</strong>
          <div class="track" aria-label="${progress}% executado">
            <span style="width:${progress}%"></span>
          </div>
          <span class="meta">${item.done}/${item.total} m</span>
        </div>
      `;
    })
    .join("");
}

function advanceDay() {
  pvs.forEach((pv) => {
    if (pv.excavated < pv.total) {
      pv.excavated = Math.min(pv.total, Number((pv.excavated + 0.35).toFixed(2)));
    }
  });

  networks.forEach((network) => {
    if (network.done < network.total) {
      network.done = Math.min(network.total, network.done + 6);
    }
  });

  render();
}

function render() {
  renderSummary();
  renderPvs();
  renderSelected();
  renderNetworks();
}

document.querySelector("#advanceDay").addEventListener("click", advanceDay);
render();
