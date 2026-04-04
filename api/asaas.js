const ASAAS_BASE = 'https://api.asaas.com/v3';
const KEY = process.env.ASAAS_API_KEY;

async function asaas(path) {
  const res = await fetch(`${ASAAS_BASE}${path}`, {
    headers: { 'access_token': KEY, 'Content-Type': 'application/json' }
  });
  if (!res.ok) throw new Error(`Asaas ${path}: ${res.status}`);
  return res.json();
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

      // 2. Total recebido no mês atual
      asaas(`/payments?status=RECEIVED&paymentDate[ge]=${primeiroDia(0)}&paymentDate[le]=${ultimoDia(0)}&limit=1`),

      // 3. Parcelas a receber — mês atual (pendentes)
      asaas(`/payments?status=PENDING&dueDate[ge]=${primeiroDia(0)}&dueDate[le]=${ultimoDia(0)}&limit=1`),

      // 4. Parcelas a receber — próximo mês
      asaas(`/payments?status=PENDING&dueDate[ge]=${primeiroDia(1)}&dueDate[le]=${ultimoDia(1)}&limit=1`),

      // 5. Parcelas a receber — mês +2
      asaas(`/payments?status=PENDING&dueDate[ge]=${primeiroDia(2)}&dueDate[le]=${ultimoDia(2)}&limit=1`),

      // 6. Parcelas a receber — mês +3
      asaas(`/payments?status=PENDING&dueDate[ge]=${primeiroDia(3)}&dueDate[le]=${ultimoDia(3)}&limit=1`),

      // 7. Cobranças vencidas / inadimplentes
      asaas(`/payments?status=OVERDUE&dueDate[le]=${hoje()}&limit=1`),

      // 8. Cobranças ativas (PENDING futuras)
      asaas(`/payments?status=PENDING&dueDate[ge]=${hoje()}&limit=1`),
    ]);

    // Soma dos valores recebidos no mês (precisa paginar se > 100, mas o totalCount é suficiente para o valor total via stats)
    // Usamos o totalValue dos metadados quando disponível
    const recebidoValor   = recebidoMes.totalValue   ?? recebidoMes.totalCount * 0;
    const vencerM0Valor   = vencerM0.totalValue       ?? 0;
    const vencerM1Valor   = vencerM1.totalValue       ?? 0;
    const vencerM2Valor   = vencerM2.totalValue       ?? 0;
    const vencerM3Valor   = vencerM3.totalValue       ?? 0;

    res.status(200).json({
      saldo: balanceData.balance ?? 0,
      recebidoMes: {
        valor: recebidoValor,
        count: recebidoMes.totalCount ?? 0,
      },
      receber: [
        { mes: primeiroDia(0).slice(0, 7), valor: vencerM0Valor, count: vencerM0.totalCount ?? 0 },
        { mes: primeiroDia(1).slice(0, 7), valor: vencerM1Valor, count: vencerM1.totalCount ?? 0 },
        { mes: primeiroDia(2).slice(0, 7), valor: vencerM2Valor, count: vencerM2.totalCount ?? 0 },
        { mes: primeiroDia(3).slice(0, 7), valor: vencerM3Valor, count: vencerM3.totalCount ?? 0 },
      ],
      inadimplencia: {
        valor: inadimplentes.totalValue ?? 0,
        count: inadimplentes.totalCount ?? 0,
      },
      ativas: {
        count: ativas.totalCount ?? 0,
        valor: ativas.totalValue ?? 0,
      },
      atualizadoEm: new Date().toISOString(),
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
}
