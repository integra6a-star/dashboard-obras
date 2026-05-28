export default async function handler(request, response) {
  const allowedSources = new Set([
    "dados.json",
    "funcionarios.json",
    "pds_data.json",
    "medicao.json",
    "eap_producao.json",
  ]);
  const source = request.query.source || "dados.json";

  if (!allowedSources.has(source)) {
    response.status(400).json({ error: "Fonte nao permitida." });
    return;
  }

  const supabaseUrl = process.env.SUPABASE_URL;
  const supabaseKey = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SECRET_KEY;

  if (!supabaseUrl || !supabaseKey) {
    response.status(503).json({ error: "Supabase nao configurado na Vercel." });
    return;
  }

  const endpoint = new URL("/rest/v1/dashboard_snapshots", supabaseUrl);
  endpoint.searchParams.set("source", `eq.${source}`);
  endpoint.searchParams.set("select", "payload,source_updated_at,created_at");
  endpoint.searchParams.set("order", "created_at.desc");
  endpoint.searchParams.set("limit", "1");

  const supabaseResponse = await fetch(endpoint, {
    headers: {
      apikey: supabaseKey,
      Authorization: `Bearer ${supabaseKey}`,
    },
  });

  if (!supabaseResponse.ok) {
    const detail = await supabaseResponse.text();
    response.status(502).json({ error: "Erro ao consultar Supabase.", detail });
    return;
  }

  const rows = await supabaseResponse.json();
  if (!rows.length) {
    response.status(404).json({ error: "Snapshot nao encontrado." });
    return;
  }

  response.setHeader("Cache-Control", "s-maxage=60, stale-while-revalidate=300");
  response.status(200).json(rows[0].payload);
}
