(function(){
  const state = { pds: [], dados: null, ready: false };
  const css = `
    .obra-ai-fab{position:fixed;right:18px;bottom:18px;z-index:9999;border:0;border-radius:999px;background:#184f9f;color:#fff;height:54px;padding:0 18px;font-weight:900;box-shadow:0 14px 30px rgba(24,79,159,.28);cursor:pointer;display:flex;align-items:center;gap:10px}
    .obra-ai-fab span{display:inline-flex;align-items:center;justify-content:center;width:26px;height:26px;border-radius:999px;background:rgba(255,255,255,.18)}
    .obra-ai-panel{position:fixed;right:18px;bottom:84px;width:min(430px,calc(100vw - 28px));height:min(650px,calc(100vh - 110px));z-index:9999;background:#fff;border:1px solid #dbe5f2;border-radius:18px;box-shadow:0 24px 70px rgba(15,35,65,.22);display:none;overflow:hidden;color:#1f2a44;font-family:Arial,Helvetica,sans-serif}
    .obra-ai-panel.open{display:grid;grid-template-rows:auto 1fr auto}
    .obra-ai-head{background:#123f7d;color:#fff;padding:14px 16px;display:flex;align-items:center;justify-content:space-between;gap:10px}
    .obra-ai-title{font-size:16px;font-weight:900}
    .obra-ai-sub{font-size:12px;opacity:.85;margin-top:3px}
    .obra-ai-close{border:0;background:rgba(255,255,255,.14);color:#fff;border-radius:10px;width:34px;height:34px;font-size:20px;cursor:pointer}
    .obra-ai-body{padding:14px;overflow:auto;background:#f6f9fd;display:grid;align-content:start;gap:10px}
    .obra-ai-msg{max-width:92%;border-radius:14px;padding:10px 12px;font-size:13px;line-height:1.45;white-space:pre-wrap}
    .obra-ai-msg.bot{background:#fff;border:1px solid #dfe8f3;color:#273142}
    .obra-ai-msg.user{background:#184f9f;color:#fff;justify-self:end}
    .obra-ai-suggestions{display:flex;gap:8px;flex-wrap:wrap;margin-top:4px}
    .obra-ai-chip{border:1px solid #dbe5f2;background:#fff;color:#184f9f;border-radius:999px;padding:8px 10px;font-weight:800;font-size:12px;cursor:pointer}
    .obra-ai-form{display:grid;grid-template-columns:1fr 46px;gap:8px;padding:12px;background:#fff;border-top:1px solid #dbe5f2}
    .obra-ai-input{height:44px;border:1px solid #dbe5f2;border-radius:14px;padding:0 12px;font-size:14px;outline:none}
    .obra-ai-send{height:44px;border:0;border-radius:14px;background:#184f9f;color:#fff;font-weight:900;cursor:pointer}
    @media(max-width:760px){.obra-ai-fab{right:12px;bottom:12px}.obra-ai-panel{right:8px;bottom:74px;width:calc(100vw - 16px);height:calc(100vh - 92px)}}
  `;

  function norm(v){
    return String(v || '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim();
  }
  function esc(v){
    return String(v || '').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;').replaceAll("'","&#39;");
  }
  function fmtDate(iso){
    if(!iso) return '-';
    const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    return m ? `${m[3]}/${m[2]}/${m[1]}` : String(iso);
  }
  function uniq(arr){ return [...new Set(arr.filter(Boolean))]; }
  function latestDate(){
    return uniq(state.pds.map(x=>x.data)).sort().at(-1) || '';
  }
  function rowsForQuestion(q){
    let rows = state.pds.slice();
    const qn = norm(q);
    const data = qn.includes('hoje') || qn.includes('atual') ? latestDate() : '';
    if(data) rows = rows.filter(r=>r.data === data);
    const obras = uniq(state.pds.map(r=>r.obra)).sort((a,b)=>b.length-a.length);
    const obra = obras.find(o => qn.includes(norm(o)) || norm(o).split(/\s+/).filter(x=>x.length>2).every(p=>qn.includes(p)));
    if(obra) rows = rows.filter(r=>r.obra === obra);
    return rows;
  }
  function groupBy(rows, key){
    const map = new Map();
    rows.forEach(r=>{
      const k = r[key] || '-';
      if(!map.has(k)) map.set(k, []);
      map.get(k).push(r);
    });
    return [...map.entries()];
  }
  function renderResumo(rows, data){
    if(!rows.length) return 'Não encontrei lançamentos para esse filtro.';
    return groupBy(rows, 'obra').map(([obra, itens])=>{
      const equipes = uniq(itens.map(x=>x.equipe)).join(', ');
      const atividades = itens.slice(0,5).map(x=>`- ${x.equipe}: ${x.atividade}${x.pv ? ` (PV ${x.pv})` : ''}`).join('\n');
      const extra = itens.length > 5 ? `\n- mais ${itens.length - 5} lançamento(s)` : '';
      return `${obra}: ${itens.length} atividade(s), ${uniq(itens.map(x=>x.equipe)).length} equipe(s).\nEquipes: ${equipes}\n${atividades}${extra}`;
    }).join('\n\n');
  }
  function answerTransformacao(){
    const rows = state.pds.filter(r=>{
      const a = norm(`${r.atividade} ${r.obra}`);
      return a.includes('transform') || a.includes('acabamento');
    });
    if(!rows.length) return 'Ainda não encontrei transformação/acabamento de PV no PDS carregado.';
    const meses = new Map();
    rows.forEach(r=>{
      const mes = String(r.data || '').slice(0,7);
      meses.set(mes, (meses.get(mes) || 0) + 1);
    });
    const media = rows.length / Math.max(meses.size, 1);
    const atual = latestDate().slice(0,7);
    const mesAtual = meses.get(atual) || 0;
    return `Transformação/acabamento de PV no histórico do PDS:\nTotal: ${rows.length} lançamento(s)\nMeses com registro: ${meses.size}\nMédia: ${media.toFixed(1).replace('.', ',')} por mês\nMês atual (${atual}): ${mesAtual}\n\nÚltimos registros:\n` +
      rows.slice(-6).map(r=>`- ${fmtDate(r.data)} | ${r.obra} | ${r.equipe}: ${r.atividade}`).join('\n');
  }
  function answer(q){
    const qn = norm(q);
    const rows = rowsForQuestion(q);
    const data = qn.includes('hoje') || qn.includes('atual') ? latestDate() : '';
    if(qn.includes('transform') || qn.includes('acabamento')) return answerTransformacao();
    if(qn.includes('hdd') || qn.includes('furo') || qn.includes('puxe')){
      const hdd = rows.filter(r=>/hdd|furo|puxe/i.test(r.atividade || ''));
      return hdd.length ? `Encontrei ${hdd.length} atividade(s) de HDD/furo/puxe${data ? ` em ${fmtDate(data)}` : ''}:\n` + hdd.map(r=>`- ${r.obra} | ${r.equipe}: ${r.atividade}${r.pv ? ` (PV ${r.pv})` : ''}`).join('\n') : 'Não encontrei atividades de HDD/furo/puxe nesse filtro.';
    }
    if(qn.includes('quant') && qn.includes('equipe')){
      const equipes = uniq(rows.map(r=>r.equipe));
      return `${data ? `Em ${fmtDate(data)}, ` : ''}encontrei ${equipes.length} equipe(s): ${equipes.join(', ') || '-'}.`;
    }
    if(qn.includes('obra') && (qn.includes('quant') || qn.includes('quais'))){
      const obras = uniq(rows.map(r=>r.obra));
      return `${data ? `Em ${fmtDate(data)}, ` : ''}encontrei ${obras.length} obra(s):\n` + obras.map(o=>`- ${o}`).join('\n');
    }
    if(qn.includes('pv')){
      const refs = rows.filter(r=>String(r.pv || r.trecho || '').trim());
      return refs.length ? `Encontrei ${refs.length} lançamento(s) com PV/trecho:\n` + refs.slice(0,12).map(r=>`- ${r.obra} | ${r.equipe}: ${r.atividade} | PV ${r.pv || r.trecho}`).join('\n') : 'Não encontrei PV/trecho preenchido nesse filtro.';
    }
    return renderResumo(rows, data);
  }
  function addMessage(text, who='bot'){
    const body = document.querySelector('.obra-ai-body');
    const div = document.createElement('div');
    div.className = `obra-ai-msg ${who}`;
    div.innerHTML = esc(text);
    body.appendChild(div);
    body.scrollTop = body.scrollHeight;
  }
  async function load(){
    try{
      const pdsRes = await fetch('./pds_data.json?ai=' + Date.now());
      state.pds = pdsRes.ok ? await pdsRes.json() : [];
      try{
        const dadosRes = await fetch('./dados.json?ai=' + Date.now());
        state.dados = dadosRes.ok ? await dadosRes.json() : null;
      }catch(e){}
      state.ready = true;
    }catch(e){
      state.ready = false;
    }
  }
  function mount(){
    if(document.getElementById('obraAiPanel')) return;
    const style = document.createElement('style');
    style.textContent = css;
    document.head.appendChild(style);
    const btn = document.createElement('button');
    btn.className = 'obra-ai-fab';
    btn.type = 'button';
    btn.innerHTML = '<span>AI</span> Assistente';
    document.body.appendChild(btn);
    const panel = document.createElement('section');
    panel.id = 'obraAiPanel';
    panel.className = 'obra-ai-panel';
    panel.innerHTML = `
      <div class="obra-ai-head">
        <div><div class="obra-ai-title">Assistente Integra 6A</div><div class="obra-ai-sub">Consulta online os dados publicados do app</div></div>
        <button class="obra-ai-close" type="button" aria-label="Fechar">×</button>
      </div>
      <div class="obra-ai-body"></div>
      <form class="obra-ai-form">
        <input class="obra-ai-input" placeholder="Pergunte sobre obra, PDS, equipe ou PV" autocomplete="off" />
        <button class="obra-ai-send" type="submit">➜</button>
      </form>`;
    document.body.appendChild(panel);
    const open = ()=>{
      panel.classList.add('open');
      if(!panel.dataset.started){
        panel.dataset.started = '1';
        const data = latestDate();
        addMessage(`Pronto. Já carreguei ${state.pds.length} lançamentos do PDS${data ? `; última data ${fmtDate(data)}` : ''}.\n\nPode perguntar, por exemplo:`);
        const body = panel.querySelector('.obra-ai-body');
        const chips = document.createElement('div');
        chips.className = 'obra-ai-suggestions';
        ['Resumo do PDS de hoje','Quantas equipes temos hoje?','Quais obras têm HDD hoje?','Transformação/acabamento de PV por mês'].forEach(t=>{
          const c = document.createElement('button');
          c.type = 'button';
          c.className = 'obra-ai-chip';
          c.textContent = t;
          c.onclick = ()=>ask(t);
          chips.appendChild(c);
        });
        body.appendChild(chips);
      }
      panel.querySelector('.obra-ai-input').focus();
    };
    const close = ()=>panel.classList.remove('open');
    btn.onclick = open;
    panel.querySelector('.obra-ai-close').onclick = close;
    panel.querySelector('.obra-ai-form').onsubmit = e=>{
      e.preventDefault();
      const input = panel.querySelector('.obra-ai-input');
      const q = input.value.trim();
      if(!q) return;
      input.value = '';
      ask(q);
    };
  }
  function ask(q){
    addMessage(q, 'user');
    if(!state.ready){
      addMessage('Ainda não consegui carregar os dados online do app. Recarregue a página e tente novamente.');
      return;
    }
    setTimeout(()=>addMessage(answer(q)), 120);
  }
  load().then(mount);
})();
