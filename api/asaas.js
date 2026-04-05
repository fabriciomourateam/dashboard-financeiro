const ASAAS_BASE = 'https://api.asaas.com/v3';
const KEY = process.env.ASAAS_API_KEY;

async function asaas(path) {
  const res = await fetch(`${ASAAS_BASE}${path}`, {
    headers: { 'access_token': KEY, 'Content-Type': 'application/json' }
  });
  if (!res.ok) throw new Error(`Asaas ${path}: ${res.status}`);
  return res.json();
}

// Busca todas as páginas e retorna { totalCount, totalValue, data[] }
async function asaasTodos(path) {
  const limit = 100;
  let offset = 0;
  let totalCount = 0;
  let totalValue = 0;

  while (true) {
    const sep = path.includes('?') ? '&' : '?';
    const resp = await asaas(`${path}${sep}limit=${limit}&offset=${offset}`);

    totalCount = resp.totalCount ?? 0;
    const items = resp.data || [];
    totalValue += items.reduce((acc, p) => acc + (p.value || 0), 0);

    if (!resp.hasMore || items.length === 0) break;
    offset += limit;
  }

  return { totalCount, totalValue };
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
  res.setHeader('Cache-Control', 's-maxage=3600); // cache 1 hora

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

      // 2. Recebido no mês atual — todas as páginas
      asaasTodos(`/payments?status=RECEIVED&paymentDate[ge]=${primeiroDia(0)}&paymentDate[le]=${ultimoDia(0)}`),

      // 3. Parcelas a receber — mês atual
      asaasTodos(`/payments?status=PENDING&dueDate[ge]=${primeiroDia(0)}&dueDate[le]=${ultimoDia(0)}`),

      // 4. Parcelas a receber — próximo mês
      asaasTodos(`/payments?status=PENDING&dueDate[ge]=${primeiroDia(1)}&dueDate[le]=${ultimoDia(1)}`),

      // 5. Parcelas a receber — mês +2
      asaasTodos(`/payments?status=PENDING&dueDate[ge]=${primeiroDia(2)}&dueDate[le]=${ultimoDia(2)}`),

      // 6. Parcelas a receber — mês +3
      asaasTodos(`/payments?status=PENDING&dueDate[ge]=${primeiroDia(3)}&dueDate[le]=${ultimoDia(3)}`),

      // 7. Cobranças vencidas / inadimplentes
      asaasTodos(`/payments?status=OVERDUE&dueDate[le]=${hoje()}`),

      // 8. Cobranças ativas (PENDING futuras)
      asaasTodos(`/payments?status=PENDING&dueDate[ge]=${hoje()}`),
    ]);

    res.status(200).json({
      saldo: balanceData.balance ?? 0,
      recebidoMes: {
        valor: recebidoMes.totalValue,
        count: recebidoMes.totalCount,
      },
      receber: [
        { mes: primeiroDia(0).slice(0, 7), valor: vencerM0.totalValue, count: vencerM0.totalCount },
        { mes: primeiroDia(1).slice(0, 7), valor: vencerM1.totalValue, count: vencerM1.totalCount },
        { mes: primeiroDia(2).slice(0, 7), valor: vencerM2.totalValue, count: vencerM2.totalCount },
        { mes: primeiroDia(3).slice(0, 7), valor: vencerM3.totalValue, count: vencerM3.totalCount },
      ],
      inadimplencia: {
        valor: inadimplentes.totalValue,
        count: inadimplentes.totalCount,
      },
      ativas: {
        count: ativas.totalCount,
        valor: ativas.totalValue,
      },
      atualizadoEm: new Date().toISOString(),
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
}
