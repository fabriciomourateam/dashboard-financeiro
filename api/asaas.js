const ASAAS_BASE = 'https://api.asaas.com/v3';
const KEY = process.env.ASAAS_API_KEY;

async function asaas(path) {
  const res = await fetch(`${ASAAS_BASE}${path}`, {
    headers: { 'access_token': KEY, 'Content-Type': 'application/json' }
  });
  if (!res.ok) throw new Error(`Asaas ${path}: ${res.status}`);
  return res.json();
}

// Soma o campo value de todos os registros retornados
function somarValores(resp) {
  return (resp.data || []).reduce((acc, p) => acc + (p.value || 0), 0);
}

function hoje() {
  return new Date().toISOString().slice(0, 10);
}
function primeiroDia(offset = 0) {
  const d = new Date();
  d.setMonth(d.getMonth() + offset, 1);
  return d.toISOString().slice(0, 10);
}
function ultimoDia(offset = 0) {
  const d = new Date();
  d.setMonth(d.getMonth() + offset + 1, 0);
  return d.toISOString().slice(0, 10);
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Cache-Control', 's-maxage=300'); // cache 5 min

  if (!KEY) return res.status(500).json({ error: 'ASAAS_API_KEY não configurada' });

  try {
    const [
      balanceData,
      recebidoMes,
      vencerM0,
      vencerM1,
      vencerM2,
      vencerM3,
      inadimplentes,
      ativas,
    ] = await Promise.all([
      // 1. Saldo atual
      asaas('/finance/balance'),

      // 2. Recebido no mês atual — limit=500 para somar todos os valores
      asaas(`/payments?status=RECEIVED&paymentDate[ge]=${primeiroDia(0)}&paymentDate[le]=${ultimoDia(0)}&limit=500`),

      // 3. Parcelas a receber — mês atual
      asaas(`/payments?status=PENDING&dueDate[ge]=${primeiroDia(0)}&dueDate[le]=${ultimoDia(0)}&limit=500`),

      // 4. Parcelas a receber — próximo mês
      asaas(`/payments?status=PENDING&dueDate[ge]=${primeiroDia(1)}&dueDate[le]=${ultimoDia(1)}&limit=500`),

      // 5. Parcelas a receber — mês +2
      asaas(`/payments?status=PENDING&dueDate[ge]=${primeiroDia(2)}&dueDate[le]=${ultimoDia(2)}&limit=500`),

      // 6. Parcelas a receber — mês +3
      asaas(`/payments?status=PENDING&dueDate[ge]=${primeiroDia(3)}&dueDate[le]=${ultimoDia(3)}&limit=500`),

      // 7. Cobranças vencidas / inadimplentes
      asaas(`/payments?status=OVERDUE&dueDate[le]=${hoje()}&limit=500`),

      // 8. Cobranças ativas (PENDING futuras)
      asaas(`/payments?status=PENDING&dueDate[ge]=${hoje()}&limit=500`),
    ]);

    res.status(200).json({
      saldo: balanceData.balance ?? 0,
      recebidoMes: {
        valor: somarValores(recebidoMes),
        count: recebidoMes.totalCount ?? 0,
      },
      receber: [
        { mes: primeiroDia(0).slice(0, 7), valor: somarValores(vencerM0), count: vencerM0.totalCount ?? 0 },
        { mes: primeiroDia(1).slice(0, 7), valor: somarValores(vencerM1), count: vencerM1.totalCount ?? 0 },
        { mes: primeiroDia(2).slice(0, 7), valor: somarValores(vencerM2), count: vencerM2.totalCount ?? 0 },
        { mes: primeiroDia(3).slice(0, 7), valor: somarValores(vencerM3), count: vencerM3.totalCount ?? 0 },
      ],
      inadimplencia: {
        valor: somarValores(inadimplentes),
        count: inadimplentes.totalCount ?? 0,
      },
      ativas: {
        count: ativas.totalCount ?? 0,
        valor: somarValores(ativas),
      },
      atualizadoEm: new Date().toISOString(),
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
}
