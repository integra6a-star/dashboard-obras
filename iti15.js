const pvs = [
  { id: "PV-1", total: 5.68, excavated: 0, team: "A definir", start: "-", station: "0+0,00" },
  { id: "PV-2", total: 5.6, excavated: 0, team: "A definir", start: "-", station: "0+4,00" },
  { id: "PV-3", total: 5.64, excavated: 0, team: "A definir", start: "-", station: "5+9,50" },
  { id: "PV-4", total: 5.78, excavated: 0, team: "A definir", start: "-", station: "6+2,50" },
  { id: "PV-5", total: 8.4, excavated: 0, team: "A definir", start: "-", station: "13+2,50" },
  { id: "PV-6", total: 6.87, excavated: 0, team: "A definir", start: "-", station: "17+19,00" },
  { id: "PV-7", total: 4.87, excavated: 0, team: "A definir", start: "-", station: "22+15,50" },
  { id: "PV-8", total: 4.65, excavated: 0, team: "A definir", start: "-", station: "23+4,50" },
];

const networks = [
  { from: "PV-1", to: "PV-2", total: 4, done: 0 },
  { from: "PV-2", to: "PV-3", total: 105.5, done: 0 },
  { from: "PV-3", to: "PV-4", total: 13, done: 0 },
  { from: "PV-4", to: "PV-5", total: 140, done: 0 },
  { from: "PV-5", to: "PV-6", total: 96.5, done: 0 },
  { from: "PV-6", to: "PV-7", total: 96.5, done: 0 },
  { from: "PV-7", to: "PV-8", total: 9, done: 0 },
];

const maxDepth = 9;
let selectedPvId = "PV-5";

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
    <div class="metric"><span>Estaca</span><strong>${pv.station}</strong></div>
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
