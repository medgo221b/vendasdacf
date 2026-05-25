import { useState, useEffect, useCallback } from "react";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";
import {
  AreaChart, Area, BarChart, Bar, XAxis, YAxis,
  CartesianGrid, Tooltip, ResponsiveContainer, Cell
} from "recharts";

// ─── CONFIG ─────────────────────────────────────────────────────
const SUPABASE_URL = "https://hngneexnuyjwisvhtavk.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImhuZ25lZXhudXlqd2lzdmh0YXZrIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzcxNjQ4NzQsImV4cCI6MjA5Mjc0MDg3NH0.FFOZLP0t8H9lH6LvRN-G3GSMnjR-mIQW-RQ8GXRs5vA";

const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

// ─── CORES ──────────────────────────────────────────────────────
const C = {
  navy:   "#0F2240",
  blue:   "#1A3C5E",
  teal:   "#0D7377",
  green:  "#143A26", 
  gold:   "#C89B3C",
  red:    "#C0392B",
  purple: "#6B3FA0",
  bg:     "#081C15", 
  card:   "rgba(17, 30, 53, 0.65)", // Card semi-transparente
  border: "rgba(30, 48, 80, 0.5)",
  text:   "#E8EDF5",
  muted:  "#7A90B0",
};

// ─── UTILITÁRIOS ────────────────────────────────────────────────
const fmtR = (v) =>
  `R$ ${Number(v || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`;
const fmtP = (v) => (v != null ? `${Number(v).toFixed(1)}%` : "—");
const hoje = () => new Date().toISOString().split("T")[0];

// ─── ESTILOS GLOBAIS E RESPONSIVOS ──────────────────────────────
const globalCss = `
  @import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500&display=swap');
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { 
    background: ${C.bg}; 
    color: ${C.text}; 
    font-family: 'DM Sans', sans-serif; 
    min-height: 100vh; 
    overflow-x: hidden;
    position: relative;
  }
  
  /* Marca d'água de fundo */
  body::before {
    content: "";
    position: fixed;
    top: 0; left: 0; right: 0; bottom: 0;
    background: url('logo.png') no-repeat center center;
    background-size: 50%;
    opacity: 0.08;
    pointer-events: none;
    z-index: -1;
  }

  /* Efeito de vidro para os cards */
  .glass-card {
    backdrop-filter: blur(10px);
    -webkit-backdrop-filter: blur(10px);
  }

  ::-webkit-scrollbar { width: 6px; } ::-webkit-scrollbar-track { background: ${C.bg}; }
  ::-webkit-scrollbar-thumb { background: ${C.border}; border-radius: 3px; }
  input, select, textarea { font-family: inherit; }
  button { cursor: pointer; font-family: inherit; }

  /* Ajustes para telas pequenas */
  @media (max-width: 768px) {
    main { margin-left: 0 !important; padding: 16px !important; margin-top: 60px !important; }
    .sidebar { width: 100% !important; height: auto !important; position: fixed !important; top: 0 !important; left: 0 !important; border-right: none !important; border-bottom: 1px solid ${C.border} !important; z-index: 1000 !important; }
    .sidebar-header { display: none !important; }
    .sidebar-nav { display: flex !important; flex-direction: row !important; overflow-x: auto !important; padding: 8px !important; gap: 8px !important; }
    .sidebar-nav button { white-space: nowrap !important; margin-bottom: 0 !important; padding: 8px 12px !important; }
    .sidebar-footer { display: none !important; }
    .stat-grid { grid-template-columns: 1fr 1fr !important; }
    .chart-grid { grid-template-columns: 1fr !important; }
  }

  @media (max-width: 480px) {
    .stat-grid { grid-template-columns: 1fr !important; }
  }
`;

// ─── COMPONENTES BASE ───────────────────────────────────────────
const Card = ({ children, style, className }) => (
  <div className={`glass-card ${className || ""}`} style={{
    background: C.card, border: `1px solid ${C.border}`,
    borderRadius: 16, padding: 24, ...style
  }}>{children}</div>
);

const Btn = ({ children, onClick, variant = "primary", disabled, style, type = "button" }) => {
  const styles = {
    primary:  { background: C.teal,  color: "#fff" },
    success:  { background: C.green, color: "#fff" },
    danger:   { background: C.red,   color: "#fff" },
    ghost:    { background: "transparent", color: C.muted, border: `1px solid ${C.border}` },
  };
  return (
    <button type={type} onClick={onClick} disabled={disabled} style={{
      padding: "10px 20px", borderRadius: 10, border: "none",
      fontWeight: 600, fontSize: 14, transition: "all .2s",
      opacity: disabled ? 0.5 : 1, ...styles[variant], ...style
    }}>{children}</button>
  );
};

const Input = ({ label, ...props }) => (
  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
    {label && <label style={{ fontSize: 12, color: C.muted, fontWeight: 500 }}>{label}</label>}
    <input {...props} style={{
      background: "#0A1628", border: `1px solid ${C.border}`, borderRadius: 10,
      padding: "10px 14px", color: C.text, fontSize: 14, outline: "none",
      transition: "border .2s", ...props.style
    }} />
  </div>
);

const Select = ({ label, children, ...props }) => (
  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
    {label && <label style={{ fontSize: 12, color: C.muted, fontWeight: 500 }}>{label}</label>}
    <select {...props} style={{
      background: "#0A1628", border: `1px solid ${C.border}`, borderRadius: 10,
      padding: "10px 14px", color: C.text, fontSize: 14, outline: "none", ...props.style
    }}>{children}</select>
  </div>
);

const Badge = ({ children, color = C.teal }) => (
  <span style={{
    background: color + "22", color, border: `1px solid ${color}44`,
    borderRadius: 6, padding: "3px 10px", fontSize: 12, fontWeight: 600
  }}>{children}</span>
);

const StatCard = ({ label, value, sub, color = C.teal, icon }) => (
  <Card>
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
      <div>
        <p style={{ fontSize: 12, color: C.muted, marginBottom: 8, fontWeight: 500 }}>{label}</p>
        <p style={{ fontSize: 26, fontWeight: 800, fontFamily: "Syne", color }}>{value}</p>
        {sub && <p style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>{sub}</p>}
      </div>
      <span style={{ fontSize: 28, opacity: 0.7 }}>{icon}</span>
    </div>
  </Card>
);

const Alert = ({ msg, type = "error", onClose }) => type && msg ? (
  <div style={{
    padding: "12px 16px", borderRadius: 10, marginBottom: 16, fontSize: 14,
    background: type === "error" ? "#C0392B22" : "#21734622",
    border: `1px solid ${type === "error" ? "#C0392B66" : "#21734666"}`,
    color: type === "error" ? "#FF6B6B" : "#4CAF50",
    display: "flex", justifyContent: "space-between", alignItems: "center"
  }}>
    {msg}
    {onClose && <button onClick={onClose} style={{ background: "none", border: "none", color: "inherit", fontSize: 18 }}>×</button>}
  </div>
) : null;

// ─── LOGIN ──────────────────────────────────────────────────────
function Login({ onLogin }) {
  const [email, setEmail] = useState("");
  const [senha, setSenha] = useState("");
  const [erro, setErro]   = useState("");
  const [load, setLoad]   = useState(false);

  const entrar = async (e) => {
    e.preventDefault();
    setErro(""); setLoad(true);
    const { error } = await supabase.auth.signInWithPassword({ email, password: senha });
    if (error) setErro("Email ou senha incorretos.");
    else onLogin();
    setLoad(false);
  };

  return (
    <div style={{
      minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center",
      background: `radial-gradient(ellipse at 30% 50%, ${C.navy} 0%, ${C.bg} 70%)`
    }}>
      <div style={{ width: "100%", maxWidth: 400, padding: 24 }}>
        <div style={{ textAlign: "center", marginBottom: 40 }}>
          <img 
            src="logo.png" 
            alt="Logo D.A." 
            style={{ width: 80, height: 80, marginBottom: 16, borderRadius: 16, objectFit: 'cover' }}
            onError={(e) => { e.target.style.display = 'none'; e.target.nextSibling.style.display = 'block'; }}
          />
          <div style={{ fontSize: 48, marginBottom: 12, display: 'none' }}>🛒</div>
          <h1 style={{ fontFamily: "Syne", fontSize: 22, fontWeight: 800, color: C.text, lineHeight: 1.2 }}>
            Diretório Acadêmico Cleusa Ferri
          </h1>
          <p style={{ color: C.muted, fontSize: 13, marginTop: 8 }}>UAM SJC - GESTÃO 2026</p>
        </div>
        <Card>
          <form onSubmit={entrar} style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            <Alert msg={erro} type="error" onClose={() => setErro("")} />
            <Input label="Email" type="email" value={email}
              onChange={e => setEmail(e.target.value)} placeholder="seu@email.com" required />
            <Input label="Senha" type="password" value={senha}
              onChange={e => setSenha(e.target.value)} placeholder="••••••••" required />
            <Btn type="submit" disabled={load} style={{ width: "100%", padding: 14, marginTop: 8 }}>
              {load ? "Entrando..." : "Entrar"}
            </Btn>
          </form>
        </Card>
      </div>
    </div>
  );
}

// ─── SIDEBAR ────────────────────────────────────────────────────
const MENU_ESTRUTURA = [
  { id: "dashboard", label: "Dashboard", icon: "📊" },
  { 
    id: "grupo_vendas", 
    label: "Vendas", 
    icon: "🛒", 
    subs: [
      { id: "vendas", label: "Nova Venda", icon: "🆕" },
      { id: "historico", label: "Histórico", icon: "🗂️" }
    ]
  },
  { 
    id: "grupo_produtos", 
    label: "Produtos", 
    icon: "📦", 
    subs: [
      { id: "produtos", label: "Gestão / Estoque", icon: "📉" },
      { id: "kits", label: "Kits / Combos", icon: "🍱" }
    ]
  },
  { id: "financeiro", label: "Financeiro", icon: "💸" },
];

function Sidebar({ tab, setTab, onLogout, user }) {
  const [alertas, setAlertas] = useState(0);
  const [abertos, setAbus] = useState(["grupo_vendas", "grupo_produtos"]); // Mantém grupos abertos

  useEffect(() => {
    supabase.from("produtos").select("id").eq("ativo", true).lte("estoque_atual", 5)
      .then(({ data }) => setAlertas(data?.length || 0));
  }, [tab]);

  const toggleGrupo = (id) => setAbus(prev => prev.includes(id) ? prev.filter(x => x !== id) : [...prev, id]);

  const RenderBtn = ({ item, isSub = false }) => {
    const active = tab === item.id;
    return (
      <button onClick={() => setTab(item.id)} style={{
        width: "100%", display: "flex", alignItems: "center", gap: 10,
        padding: isSub ? "9px 14px 9px 40px" : "11px 14px", 
        borderRadius: 10, border: "none", marginBottom: 2,
        background: active ? C.teal + "22" : "transparent",
        color: active ? C.teal : C.muted,
        fontWeight: active ? 600 : 400, fontSize: isSub ? 13 : 14, transition: "all .2s",
        borderLeft: active ? `3px solid ${C.teal}` : "3px solid transparent",
        position: 'relative'
      }}>
        <span>{item.icon}</span> {item.label}
        {item.id === 'produtos' && alertas > 0 && (
          <span style={{ position: 'absolute', right: 10, background: C.red, color: '#fff', fontSize: 10, padding: '2px 6px', borderRadius: 10 }}>{alertas}</span>
        )}
      </button>
    );
  };

  return (
    <aside className="sidebar" style={{
      width: 220, height: "100vh", background: C.card,
      borderRight: `1px solid ${C.border}`, display: "flex",
      flexDirection: "column", position: "fixed", top: 0, left: 0, zIndex: 100
    }}>
      <div className="sidebar-header" style={{ padding: "24px 20px", borderBottom: `1px solid ${C.border}`, textAlign: 'center' }}>
        <img src="logo.png" alt="Logo" style={{ width: 50, height: 50, marginBottom: 10, borderRadius: 10, objectFit: 'cover' }}
          onError={(e) => { e.target.style.display = 'none'; e.target.nextSibling.style.display = 'block'; }} />
        <div style={{ fontSize: 24, marginBottom: 6, display: 'none' }}>🛒</div>
        <h2 style={{ fontFamily: "Syne", fontWeight: 800, fontSize: 13, color: C.text }}>D.A. Cleusa Ferri</h2>
        <p style={{ fontSize: 9, color: C.muted }}>UAM SJC - 2026</p>
      </div>

      <nav className="sidebar-nav" style={{ flex: 1, padding: "12px", overflowY: 'auto' }}>
        {MENU_ESTRUTURA.map(m => (
          <div key={m.id}>
            {m.subs ? (
              <>
                <button onClick={() => toggleGrupo(m.id)} style={{
                  width: "100%", display: "flex", alignItems: "center", justifyContent: 'space-between',
                  padding: "11px 14px", background: 'transparent', border: 'none', color: C.text,
                  fontSize: 14, fontWeight: 700, cursor: 'pointer', marginTop: 8
                }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}><span>{m.icon}</span> {m.label}</div>
                  <span style={{ fontSize: 10, opacity: 0.5 }}>{abertos.includes(m.id) ? "▼" : "▶"}</span>
                </button>
                {abertos.includes(m.id) && m.subs.map(s => <RenderBtn key={s.id} item={s} isSub />)}
              </>
            ) : <RenderBtn item={m} />}
          </div>
        ))}
      </nav>

      <div className="sidebar-footer" style={{ padding: "16px", borderTop: `1px solid ${C.border}` }}>
        <Btn variant="ghost" onClick={onLogout} style={{ width: "100%", fontSize: 12 }}>Sair</Btn>
      </div>
    </aside>
  );
}

// ─── DASHBOARD ──────────────────────────────────────────────────
function Dashboard() {
  const [stats, setStats]    = useState(null);
  const [vendas, setVendas]  = useState([]);
  const [prods, setProds]    = useState([]);
  const [turmas, setTurmas]  = useState([]);
  const [loading, setLoading] = useState(true);
  
  const [filtroData, setFiltroData] = useState("ultimo_mes");
  const [dataDe, setDataDe]         = useState("");
  const [dataAte, setDataAte]       = useState("");

  const carregar = useCallback(async () => {
    setLoading(true);
    const dHoje = hoje();
    
    let queryVendas = supabase.from("vendas").select("*").eq("status", "Pago e Entregue");
    let queryMovs = supabase.from("movimentacoes_financeiras").select("*");
    
    if (filtroData === "hoje") {
      queryVendas = queryVendas.eq("data_venda", dHoje);
      queryMovs = queryMovs.gte("created_at", dHoje);
    } else if (filtroData === "ultimo_mes") {
      const umMesAtras = new Date();
      umMesAtras.setDate(umMesAtras.getDate() - 30);
      const dataIso = umMesAtras.toISOString().split("T")[0];
      queryVendas = queryVendas.gte("data_venda", dataIso);
      queryMovs = queryMovs.gte("created_at", dataIso);
    } else if (filtroData === "custom") {
      if (dataDe) { queryVendas = queryVendas.gte("data_venda", dataDe); queryMovs = queryMovs.gte("created_at", dataDe); }
      if (dataAte) { queryVendas = queryVendas.lte("data_venda", dataAte); queryMovs = queryMovs.lte("created_at", dataAte + "T23:59:59"); }
    }

    const [{ data: v }, { data: p }, { data: m }] = await Promise.all([
      queryVendas,
      supabase.from("dashboard_produtos").select("*"),
      queryMovs
    ]);

    const vs = v || []; const ps = p || []; const ms = m || [];
    
    // Cálculo de Financeiro
    const totalEntradas = vs.reduce((s, r) => s + r.preco_venda * r.quantidade, 0);
    const totalSaidas   = ms.filter(mov => mov.tipo === "saida").reduce((s, mov) => s + mov.valor, 0);
    const lucroBruto    = totalEntradas; // Receita das vendas
    
    // Custo de mercadoria (opcional, se quiser lucro líquido absoluto)
    const custoMap   = Object.fromEntries(ps.map(p => [p.id, p.preco_custo]));
    const custoProdutos = vs.reduce((s, r) => s + (custoMap[r.produto_id] || 0) * r.quantidade, 0);
    
    const lucroLiquidoReal = totalEntradas - totalSaidas - custoProdutos;

    // Agrupa por dia para o gráfico
    const diasMap = {};
    vs.forEach(r => {
      const d = r.data_venda;
      if (!diasMap[d]) diasMap[d] = { data: d, receita: 0, saida: 0 };
      diasMap[d].receita += r.preco_venda * r.quantidade;
    });
    ms.forEach(mov => {
      const d = mov.created_at.split("T")[0];
      if (!diasMap[d]) diasMap[d] = { data: d, receita: 0, saida: 0 };
      if (mov.tipo === "saida") diasMap[d].saida += mov.valor;
    });

    const porDia = Object.values(diasMap)
      .sort((a, b) => a.data.localeCompare(b.data))
      .map(d => ({ ...d, data: d.data.slice(5).replace("-", "/") }));
    
    // Ranking de turmas
    const tMap = {};
    vs.forEach(r => {
      const t = r.turma || "Não informada";
      if (!tMap[t]) tMap[t] = { nome: t, total: 0 };
      tMap[t].total += r.preco_venda * r.quantidade;
    });
    const rankingTurmas = Object.values(tMap).sort((a, b) => b.total - a.total).slice(0, 5);

    // Stats de Hoje
    let totalEntradasHoje = 0;
    let totalSaidasHoje = 0;
    if (filtroData === "hoje") {
      totalEntradasHoje = totalEntradas;
      totalSaidasHoje = totalSaidas;
    } else {
      const [{ data: vh }, { data: mh }] = await Promise.all([
        supabase.from("vendas").select("preco_venda, quantidade").eq("status", "Pago e Entregue").eq("data_venda", dHoje),
        supabase.from("movimentacoes_financeiras").select("valor").eq("tipo", "saida").gte("created_at", dHoje)
      ]);
      totalEntradasHoje = (vh || []).reduce((s, r) => s + r.preco_venda * r.quantidade, 0);
      totalSaidasHoje = (mh || []).reduce((s, r) => s + r.valor, 0);
    }

    setStats({ 
      totalEntradas, totalSaidas, lucroLiquidoReal, 
      faturamentoHoje: totalEntradasHoje,
      saldoHoje: totalEntradasHoje - totalSaidasHoje,
      totalPedidos: vs.length,
      margemGeral: totalEntradas > 0 ? (lucroLiquidoReal / totalEntradas * 100).toFixed(1) : 0 
    });
    setVendas(porDia);
    setTurmas(rankingTurmas);
    setProds(ps); 
    setLoading(false);
  }, [filtroData, dataDe, dataAte]);

  useEffect(() => {
    carregar();
    const ch = supabase.channel("dashboard")
      .on("postgres_changes", { event: "*", schema: "public", table: "vendas" }, carregar)
      .on("postgres_changes", { event: "*", schema: "public", table: "produtos" }, carregar)
      .subscribe();
    return () => supabase.removeChannel(ch);
  }, [carregar]);

  if (loading && !stats) return <div style={{ color: C.muted, padding: 40 }}>Carregando...</div>;

  const prodsCriticos = prods.filter(p => p.estoque_atual <= 5);

  return (
    <div>
      <div style={{ display: "flex", flexWrap: "wrap", justifyContent: "space-between", alignItems: "center", marginBottom: 24, gap: 16 }}>
        <h1 style={{ fontFamily: "Syne", fontSize: 24, fontWeight: 800 }}>Dashboard</h1>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 10, alignItems: "end" }}>
          <Select label="Período" value={filtroData} onChange={e => setFiltroData(e.target.value)} style={{ width: 140 }}>
            <option value="hoje">Hoje</option>
            <option value="ultimo_mes">Últimos 30 dias</option>
            <option value="custom">Personalizado</option>
          </Select>
          {filtroData === "custom" && (
            <>
              <Input type="date" label="De" value={dataDe} onChange={e => setDataDe(e.target.value)} />
              <Input type="date" label="Até" value={dataAte} onChange={e => setDataAte(e.target.value)} />
            </>
          )}
          <Btn variant="ghost" onClick={carregar} style={{ height: 40 }}>🔄</Btn>
        </div>
      </div>

      {prodsCriticos.length > 0 && (
        <Card style={{ background: C.red + '11', borderColor: C.red + '33', marginBottom: 24 }}>
          <h4 style={{ color: C.red, fontSize: 14, display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12 }}>
            ⚠️ ATENÇÃO: Produtos com estoque crítico
          </h4>
          <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
            {prodsCriticos.map(p => (
              <Badge key={p.id} color={C.red}>{p.nome}: {p.estoque_atual} un.</Badge>
            ))}
          </div>
        </Card>
      )}

      <div className="stat-grid" style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 16, marginBottom: 24 }}>
        <StatCard label="Faturamento Hoje" value={fmtR(stats?.faturamentoHoje)} icon="🚀" color={C.teal} sub="Total bruto vendido hoje" />
        <StatCard label="Saldo de Caixa Hoje" value={fmtR(stats?.saldoHoje)} icon="👛" color={stats?.saldoHoje >= 0 ? C.green : C.red} sub="Entrou - Saiu hoje" />
        <StatCard label="Faturamento Período" value={fmtR(stats?.totalEntradas)} icon="💰" color={C.green} />
        <StatCard label="Lucro Líquido Real" value={fmtR(stats?.lucroLiquidoReal)} icon="📈" color={stats?.lucroLiquidoReal >= 0 ? C.green : C.red} sub="Período selecionado" />
      </div>
      
      <div className="chart-grid" style={{ display: "grid", gridTemplateColumns: "1.6fr 1fr", gap: 16, marginBottom: 24 }}>
        <Card>
          <h3 style={{ fontFamily: "Syne", fontWeight: 700, marginBottom: 20, fontSize: 15 }}>Fluxo de Caixa Diário</h3>
          {vendas.length > 0 ? (
            <ResponsiveContainer width="100%" height={250}>
              <AreaChart data={vendas}>
                <defs>
                  <linearGradient id="gradEntrada" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%"  stopColor={C.teal} stopOpacity={0.3} />
                    <stop offset="95%" stopColor={C.teal} stopOpacity={0} />
                  </linearGradient>
                  <linearGradient id="gradSaida" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%"  stopColor={C.red} stopOpacity={0.2} />
                    <stop offset="95%" stopColor={C.red} stopOpacity={0} />
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                <XAxis dataKey="data" tick={{ fill: C.muted, fontSize: 11 }} />
                <YAxis tick={{ fill: C.muted, fontSize: 11 }} />
                <Tooltip contentStyle={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 8 }} />
                <Area type="monotone" dataKey="receita" stroke={C.teal} name="Entradas" strokeWidth={2} fill="url(#gradEntrada)" />
                <Area type="monotone" dataKey="saida" stroke={C.red} name="Saídas" strokeWidth={2} fill="url(#gradSaida)" />
              </AreaChart>
            </ResponsiveContainer>
          ) : <p style={{ color: C.muted, textAlign: "center", paddingTop: 60 }}>Sem dados no período</p>}
        </Card>

        <Card>
          <h3 style={{ fontFamily: "Syne", fontWeight: 700, marginBottom: 20, fontSize: 15 }}>Ranking de Turmas (R$)</h3>
          {turmas.length > 0 ? (
            <ResponsiveContainer width="100%" height={250}>
              <BarChart data={turmas} layout="vertical" margin={{ left: -10 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                <XAxis type="number" tick={{ fill: C.muted, fontSize: 11 }} />
                <YAxis type="category" dataKey="nome" width={90} tick={{ fill: C.muted, fontSize: 10 }} />
                <Tooltip contentStyle={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 8 }} formatter={v => [fmtR(v), "Total"]} />
                <Bar dataKey="total" radius={[0, 6, 6, 0]} fill={C.gold} />
              </BarChart>
            </ResponsiveContainer>
          ) : <p style={{ color: C.muted, textAlign: "center", paddingTop: 60 }}>Sem dados de turmas</p>}
        </Card>
      </div>

      <div className="chart-grid" style={{ display: "grid", gridTemplateColumns: "1fr", gap: 16, marginBottom: 24 }}>
        <Card>
          <h3 style={{ fontFamily: "Syne", fontWeight: 700, marginBottom: 20, fontSize: 15 }}>Top Produtos (Vendas)</h3>
          {prods.length > 0 ? (
            <ResponsiveContainer width="100%" height={250}>
              <BarChart data={prods.filter(p => p.total_vendido > 0).sort((a,b) => b.total_vendido - a.total_vendido).slice(0,10)} layout="vertical" margin={{ left: -10 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                <XAxis type="number" tick={{ fill: C.muted, fontSize: 11 }} />
                <YAxis type="category" dataKey="nome" width={110} tick={{ fill: C.muted, fontSize: 10 }} tickFormatter={v => v.length > 20 ? v.slice(0, 20) + "…" : v} />
                <Tooltip contentStyle={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 8 }} formatter={v => [v + " un.", "Vendido"]} />
                <Bar dataKey="total_vendido" radius={[0, 6, 6, 0]}>
                  {prods.map((_, i) => <Cell key={i} fill={[C.teal, C.green, C.purple, C.gold, C.blue][i % 5]} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          ) : <p style={{ color: C.muted, textAlign: "center", paddingTop: 60 }}>Sem dados ainda</p>}
        </Card>
      </div>

      <Card>
        <h3 style={{ fontFamily: "Syne", fontWeight: 700, marginBottom: 16, fontSize: 15 }}>Estoque e Margem Geral</h3>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, minWidth: 600 }}>
            <thead>
              <tr style={{ borderBottom: `1px solid ${C.border}` }}>
                {["Produto","Est.","Vendido","Receita","Margem %","Margem DA %","Status"].map(h => (
                  <th key={h} style={{ padding: "8px 12px", color: C.muted, fontWeight: 500, textAlign: "left", whiteSpace: "nowrap" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {prods.map((p, i) => {
                const estColor = p.status_estoque === "Esgotado" ? C.red : p.status_estoque === "Baixo" ? C.gold : C.green;
                return (
                  <tr key={p.id} style={{ borderBottom: `1px solid ${C.border}22`, background: i % 2 === 0 ? "transparent" : "#ffffff04" }}>
                    <td style={{ padding: "10px 12px", fontWeight: 500 }}>{p.nome}</td>
                    <td style={{ padding: "10px 12px", color: estColor, fontWeight: 700 }}>{p.estoque_atual}</td>
                    <td style={{ padding: "10px 12px", color: C.muted }}>{p.total_vendido || 0}</td>
                    <td style={{ padding: "10px 12px" }}>{fmtR(p.receita_total)}</td>
                    <td style={{ padding: "10px 12px" }}>
                      {p.margem_normal_pct != null ? <span style={{ color: p.margem_normal_pct < 20 ? C.red : p.margem_normal_pct >= 40 ? C.green : C.gold }}>{fmtP(p.margem_normal_pct)}</span> : "—"}
                    </td>
                    <td style={{ padding: "10px 12px" }}>
                      {p.margem_da_pct != null ? <span style={{ color: C.purple }}>{fmtP(p.margem_da_pct)}</span> : "—"}
                    </td>
                    <td style={{ padding: "10px 12px" }}>
                      <Badge color={estColor}>{p.status_estoque}</Badge>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

// ─── NOVA VENDA ─────────────────────────────────────────────────
function NovaVenda() {
  const [prods, setProds]       = useState([]);
  const [kits, setKits]         = useState([]);
  const [itens, setItens]       = useState([{ prodId: "", qtd: 1, isKit: false }]);
  const [comprador, setComprador] = useState("");
  const [turma, setTurma]       = useState("");
  const [data, setData]         = useState(hoje());
  const [pgto, setPgto]         = useState("Pix");
  const [msg, setMsg]           = useState(null);
  const [load, setLoad]         = useState(false);

  useEffect(() => {
    Promise.all([
      supabase.from("produtos").select("*").eq("ativo", true),
      supabase.from("kits").select("*, kit_itens(*)")
    ]).then(([{ data: p }, { data: k }]) => {
      setProds(p || []);
      setKits(k || []);
    });
  }, []);

  const addItem = () => setItens([...itens, { prodId: "", qtd: 1, isKit: false }]);
  const remItem = (i) => setItens(itens.filter((_, j) => j !== i));
  const setItem = (i, field, val) => setItens(itens.map((it, j) => j === i ? { ...it, [field]: val } : it));
  
  const getProd = (id, isKit) => isKit ? kits.find(k => k.id === id) : prods.find(p => p.id === id);
  const total = itens.reduce((s, it) => {
    const item = getProd(it.prodId, it.isKit);
    return s + (item ? (item.preco_normal || item.preco_venda) * it.qtd : 0);
  }, 0);

  const registrar = async () => {
    const validos = itens.filter(it => it.prodId);
    if (!comprador.trim()) return setMsg({ t: "error", v: "Informe o comprador." });
    if (!validos.length) return setMsg({ t: "error", v: "Adicione itens." });

    setLoad(true); setMsg(null);
    let erros = [];

    for (const it of validos) {
      if (it.isKit) {
        const kit = kits.find(k => k.id === it.prodId);
        const { error } = await supabase.from("vendas").insert({
          produto_nome: `[KIT] ${kit.nome}`,
          comprador: comprador.trim(),
          turma: turma.trim(),
          data_venda: data,
          preco_venda: kit.preco_venda,
          quantidade: it.qtd,
          forma_pagamento: pgto,
          status: "Pago e Entregue"
        });
        
        if (error) erros.push(`Erro no kit: ${error.message}`);
        for (const ki of kit.kit_itens) {
          await supabase.rpc("ajustar_estoque", { p_id: ki.produto_id, p_qtd: -(ki.quantidade * it.qtd) });
        }
      } else {
        const p = prods.find(p => p.id === it.prodId);
        const { data: res } = await supabase.rpc("registrar_venda", {
          p_produto_id: it.prodId, p_comprador: comprador.trim(), p_turma: turma.trim(), 
          p_data_venda: data, p_preco_venda: p.preco_normal, p_quantidade: it.qtd, p_forma_pagamento: pgto,
        });
        if (!res?.ok) erros.push(res?.erro);
      }
    }

    if (erros.length) setMsg({ t: "error", v: erros.join(" | ") });
    else {
      setMsg({ t: "success", v: "Venda registrada!" });
      setItens([{ prodId: "", qtd: 1, isKit: false }]);
      setComprador(""); setTurma("");
    }
    setLoad(false);
  };

  return (
    <div style={{ maxWidth: 680, margin: "0 auto" }}>
      <h1 style={{ fontFamily: "Syne", fontSize: 24, fontWeight: 800, marginBottom: 24 }}>🆕 Nova Venda</h1>
      <Card style={{ marginBottom: 16 }}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <div style={{ gridColumn: "1/-1" }}><Input label="Comprador *" value={comprador} onChange={e => setComprador(e.target.value)} /></div>
          <Input label="Turma" value={turma} onChange={e => setTurma(e.target.value)} />
          <Input label="Data *" type="date" value={data} onChange={e => setData(e.target.value)} />
          <Select label="Pagamento" value={pgto} onChange={e => setPgto(e.target.value)} style={{ gridColumn: "1/-1" }}>
            {["Pix","Dinheiro","Cartão","Transferência"].map(o => <option key={o}>{o}</option>)}
          </Select>
        </div>
      </Card>
      <Card style={{ marginBottom: 16 }}>
        <h3 style={{ fontSize: 14, color: C.muted, textTransform: 'uppercase', marginBottom: 16 }}>Itens da Venda</h3>
        {itens.map((it, i) => (
          <div key={i} style={{ display: "grid", gridTemplateColumns: "100px 1fr 80px 36px", gap: 8, marginBottom: 12, alignItems: "end" }}>
            <Select label={i === 0 ? "Tipo" : ""} value={it.isKit ? "kit" : "prod"} onChange={e => setItem(i, "isKit", e.target.value === "kit")}>
              <option value="prod">Produto</option>
              <option value="kit">Combo/Kit</option>
            </Select>
            <Select label={i === 0 ? "Item" : ""} value={it.prodId} onChange={e => setItem(i, "prodId", e.target.value)}>
              <option value="">Selecionar...</option>
              {it.isKit ? kits.map(k => <option key={k.id} value={k.id}>{k.nome}</option>) : prods.map(p => <option key={p.id} value={p.id}>{p.nome} ({p.estoque_atual} un.)</option>)}
            </Select>
            <Input label={i === 0 ? "Qtd" : ""} type="number" value={it.qtd} onChange={e => setItem(i, "qtd", parseInt(e.target.value) || 1)} />
            <button onClick={() => remItem(i)} style={{ background: "#C0392B22", border: `1px solid ${C.red}44`, borderRadius: 8, color: C.red, height: 38 }}>×</button>
          </div>
        ))}
        <Btn variant="ghost" onClick={addItem} style={{ width: "100%", marginBottom: 16 }}>+ Adicionar item</Btn>
        <div style={{ background: C.blue + "33", borderRadius: 12, padding: 14, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <span>Total</span>
          <span style={{ fontSize: 22, fontWeight: 800, color: C.teal }}>{fmtR(total)}</span>
        </div>
      </Card>
      <Alert msg={msg?.v} type={msg?.t} onClose={() => setMsg(null)} />
      <Btn variant="success" onClick={registrar} disabled={load} style={{ width: "100%", padding: 14 }}>{load ? "Registrando..." : "Finalizar Venda"}</Btn>
    </div>
  );
}

// ─── PRODUTOS ───────────────────────────────────────────────────
function Produtos() {
  const [prods, setProds]   = useState([]);
  const [form, setForm]     = useState({ nome: "", estoque_inicial: 0, preco_normal: "", preco_da: "", preco_custo: "" });
  const [msg, setMsg]       = useState(null);
  const [load, setLoad]     = useState(false);
  const [editId, setEditId] = useState(null);

  const carregar = async () => {
    const { data } = await supabase.from("dashboard_produtos").select("*").order("nome");
    setProds(data || []);
  };

  useEffect(() => {
    carregar();
    const ch = supabase.channel("produtos_ch").on("postgres_changes", { event: "*", schema: "public", table: "produtos" }, carregar).subscribe();
    return () => supabase.removeChannel(ch);
  }, []);

  const salvar = async () => {
    if (!form.nome.trim()) return setMsg({ t: "error", v: "Informe o nome." });
    if (!Number(form.preco_normal)) return setMsg({ t: "error", v: "Informe o preço de venda." });
    setLoad(true); setMsg(null);
    const payload = {
      nome: form.nome.trim(), estoque_inicial: Number(form.estoque_inicial) || 0,
      estoque_atual: editId ? undefined : Number(form.estoque_inicial) || 0,
      preco_normal: Number(form.preco_normal) || 0, preco_da: Number(form.preco_da) || 0,
      preco_custo: Number(form.preco_custo) || 0,
    };
    if (editId) delete payload.estoque_atual;
    const { error } = editId ? await supabase.from("produtos").update(payload).eq("id", editId) : await supabase.from("produtos").insert(payload);
    if (error) setMsg({ t: "error", v: error.message.includes("unique") ? `Produto "${form.nome}" já existe.` : error.message });
    else {
      setMsg({ t: "success", v: editId ? "Produto atualizado!" : `"${form.nome}" cadastrado!` });
      setForm({ nome: "", estoque_inicial: 0, preco_normal: "", preco_da: "", preco_custo: "" });
      setEditId(null); carregar();
    }
    setLoad(false);
  };

  const editar = (p) => {
    setEditId(p.id);
    setForm({ nome: p.nome, estoque_inicial: p.estoque_inicial, preco_normal: p.preco_normal, preco_da: p.preco_da, preco_custo: p.preco_custo });
  };

  const F = (field, label, opts = {}) => <Input label={label} value={form[field]} onChange={e => setForm({ ...form, [field]: e.target.value })} {...opts} />;

  return (
    <div style={{ maxWidth: "1200px", margin: "0 auto" }}>
      <h1 style={{ fontFamily: "Syne", fontSize: 24, fontWeight: 800, marginBottom: 24 }}>📦 Produtos</h1>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 24, alignItems: "start" }}>
        <div style={{ flex: "1 1 340px", minWidth: 300 }}>
          <Card>
            <h3 style={{ fontFamily: "Syne", fontWeight: 700, marginBottom: 16, fontSize: 14, color: C.muted, textTransform: "uppercase", letterSpacing: 1 }}>{editId ? "✏️ Editar Produto" : "➕ Novo Produto"}</h3>
            <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
              <Alert msg={msg?.v} type={msg?.t} onClose={() => setMsg(null)} />
              {F("nome", "Nome do produto *", { placeholder: "Ex: Camiseta Preta XG" })}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                {F("estoque_inicial", "Estoque inicial", { type: "number", min: 0 })}
                {F("preco_normal", "Preço venda (R$) *", { type: "number", step: "0.01", placeholder: "0,00" })}
                {F("preco_da", "Preço DA (R$)", { type: "number", step: "0.01", placeholder: "0,00" })}
                {F("preco_custo", "Preço custo (R$)", { type: "number", step: "0.01", placeholder: "0,00" })}
              </div>
              <p style={{ fontSize: 11, color: C.muted }}>O custo calcula a margem de lucro no Dashboard.</p>
              <div style={{ display: "flex", gap: 8 }}>
                <Btn variant="success" onClick={salvar} disabled={load} style={{ flex: 1 }}>{load ? "Salvando..." : editId ? "Salvar alterações" : "Cadastrar produto"}</Btn>
                {editId && <Btn variant="ghost" onClick={() => { setEditId(null); setForm({ nome:"",estoque_inicial:0,preco_normal:"",preco_da:"",preco_custo:""}); }}>✕</Btn>}
              </div>
            </div>
          </Card>
        </div>
        <div style={{ flex: "2 1 600px", minWidth: 300 }}>
          <Card>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, minWidth: 500 }}>
                <thead>
                  <tr style={{ borderBottom: `1px solid ${C.border}` }}>
                    {["Produto","Est.","Preço","Custo","Margem%","Margem DA%","Status",""].map(h => (
                      <th key={h} style={{ padding: "8px 10px", color: C.muted, fontWeight: 500, textAlign: "left", whiteSpace: "nowrap" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {prods.map((p, i) => {
                    const estColor = p.status_estoque === "Esgotado" ? C.red : p.status_estoque === "Baixo" ? C.gold : C.green;
                    return (
                      <tr key={p.id} style={{ borderBottom: `1px solid ${C.border}22`, background: i % 2 === 0 ? "transparent" : "#ffffff04" }}>
                        <td style={{ padding: "9px 10px", fontWeight: 500 }}>{p.nome}</td>
                        <td style={{ padding: "9px 10px", color: estColor, fontWeight: 700 }}>{p.estoque_atual}</td>
                        <td style={{ padding: "9px 10px" }}>{fmtR(p.preco_normal)}</td>
                        <td style={{ padding: "9px 10px", color: C.muted }}>{fmtR(p.preco_custo)}</td>
                        <td style={{ padding: "9px 10px" }}>
                          {p.margem_normal_pct != null ? <span style={{ color: p.margem_normal_pct < 20 ? C.red : p.margem_normal_pct >= 40 ? C.green : C.gold }}>{fmtP(p.margem_normal_pct)}</span> : "—"}
                        </td>
                        <td style={{ padding: "9px 10px", color: C.purple }}>{p.margem_da_pct != null ? fmtP(p.margem_da_pct) : "—"}</td>
                        <td style={{ padding: "9px 10px" }}><Badge color={estColor}>{p.status_estoque}</Badge></td>
                        <td style={{ padding: "9px 10px" }}><button onClick={() => editar(p)} style={{ background: "none", border: "none", color: C.muted, fontSize: 16 }}>✏️</button></td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              {prods.length === 0 && <p style={{ color: C.muted, textAlign: "center", padding: 40 }}>Nenhum produto cadastrado ainda.</p>}
            </div>
          </Card>
        </div>
      </div>
    </div>
  );
}

// ─── KITS / COMBOS ──────────────────────────────────────────────
function Kits() {
  const [kits, setKits]     = useState([]);
  const [prods, setProds]   = useState([]);
  const [loading, setLoading] = useState(true);
  const [showAdd, setShow]  = useState(false);
  const [form, setForm]     = useState({ nome: "", preco: "", itens: [{ prodId: "", qtd: 1 }] });

  const carregar = useCallback(async () => {
    setLoading(true);
    const [{ data: k }, { data: p }] = await Promise.all([
      supabase.from("kits").select("*, kit_itens(id, produto_id, quantidade, produtos(nome))"),
      supabase.from("produtos").select("*").eq("ativo", true)
    ]);
    setKits(k || []);
    setProds(p || []);
    setLoading(false);
  }, []);

  useEffect(() => { carregar(); }, [carregar]);

  const addItem = () => setForm({ ...form, itens: [...form.itens, { prodId: "", qtd: 1 }] });
  const remItem = (i) => setForm({ ...form, itens: form.itens.filter((_, j) => j !== i) });
  const setItem = (i, field, val) => {
    const newItens = [...form.itens];
    newItens[i][field] = val;
    setForm({ ...form, itens: newItens });
  };

  const salvarKit = async (e) => {
    e.preventDefault();
    if (!form.nome || !form.preco || form.itens.some(it => !it.prodId)) return alert("Preencha tudo!");

    const { data: kit, error: ek } = await supabase.from("kits").insert({
      nome: form.nome,
      preco_venda: Number(form.preco)
    }).select().single();

    if (ek) return alert(ek.message);

    const itensPayload = form.itens.map(it => ({
      kit_id: kit.id,
      produto_id: it.prodId,
      quantidade: it.qtd
    }));

    const { error: ei } = await supabase.from("kit_itens").insert(itensPayload);
    if (ei) alert(ei.message);
    else {
      setShow(false); setForm({ nome: "", preco: "", itens: [{ prodId: "", qtd: 1 }] });
      carregar();
    }
  };

  return (
    <div style={{ maxWidth: 900, margin: '0 auto' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 24 }}>
        <h1 style={{ fontFamily: "Syne", fontSize: 24, fontWeight: 800 }}>🍱 Kits e Combos</h1>
        <Btn variant="primary" onClick={() => setShow(true)}>+ Novo Kit</Btn>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: 20 }}>
        {kits.map(k => (
          <Card key={k.id}>
            <h3 style={{ fontFamily: 'Syne', fontSize: 18, marginBottom: 8 }}>{k.nome}</h3>
            <p style={{ color: C.teal, fontWeight: 800, fontSize: 20, marginBottom: 16 }}>{fmtR(k.preco_venda)}</p>
            <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 12 }}>
              <p style={{ fontSize: 12, color: C.muted, marginBottom: 8, textTransform: 'uppercase' }}>Composição:</p>
              {k.kit_itens?.map(it => (
                <div key={it.id} style={{ fontSize: 13, marginBottom: 4 }}>
                  • {it.quantidade}x {it.produtos?.nome}
                </div>
              ))}
            </div>
            <Btn variant="ghost" style={{ width: '100%', marginTop: 16, fontSize: 12, color: C.red }} onClick={async () => {
              if (confirm('Excluir este kit?')) {
                await supabase.from("kits").delete().eq("id", k.id);
                carregar();
              }
            }}>Remover Kit</Btn>
          </Card>
        ))}
      </div>

      {showAdd && (
        <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.85)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000, padding: 16 }}>
          <Card style={{ width: '100%', maxWidth: 500, maxHeight: '90vh', overflowY: 'auto' }}>
            <h2 style={{ fontFamily: 'Syne', marginBottom: 20 }}>Criar Novo Combo</h2>
            <form onSubmit={salvarKit} style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
              <Input label="Nome do Kit" value={form.nome} onChange={e => setForm({...form, nome: e.target.value})} required placeholder="Ex: Kit Básico" />
              <Input label="Preço de Venda (R$)" type="number" step="0.01" value={form.preco} onChange={e => setForm({...form, preco: e.target.value})} required placeholder="0,00" />
              
              <div style={{ background: C.bg, padding: 16, borderRadius: 12 }}>
                <p style={{ fontSize: 12, color: C.muted, marginBottom: 12, textTransform: 'uppercase' }}>Produtos do Combo</p>
                {form.itens.map((it, i) => (
                  <div key={i} style={{ display: 'flex', gap: 8, marginBottom: 8, alignItems: 'end' }}>
                    <div style={{ flex: 1 }}>
                      <Select label={i === 0 ? "Produto" : ""} value={it.prodId} onChange={e => setItem(i, 'prodId', e.target.value)}>
                        <option value="">Escolha...</option>
                        {prods.map(p => <option key={p.id} value={p.id}>{p.nome}</option>)}
                      </Select>
                    </div>
                    <div style={{ width: 70 }}>
                      <Input label={i === 0 ? "Qtd" : ""} type="number" value={it.qtd} onChange={e => setItem(i, 'qtd', parseInt(e.target.value) || 1)} />
                    </div>
                    <button type="button" onClick={() => remItem(i)} style={{ padding: 10, background: 'none', border: 'none', color: C.red }}>×</button>
                  </div>
                ))}
                <Btn variant="ghost" onClick={addItem} style={{ width: '100%', fontSize: 12, marginTop: 8 }}>+ Adicionar Item ao Combo</Btn>
              </div>

              <div style={{ display: 'flex', gap: 10 }}>
                <Btn variant="ghost" onClick={() => setShow(false)} style={{ flex: 1 }}>Cancelar</Btn>
                <Btn type="submit" variant="success" style={{ flex: 1 }}>Salvar Kit</Btn>
              </div>
            </form>
          </Card>
        </div>
      )}
    </div>
  );
}

// ─── HISTÓRICO DE VENDAS ─────────────────────────────────────────
function Historico() {
  const [vendas, setVendas] = useState([]);
  const [filtro, setFiltro] = useState("");
  const [loading, setLoading] = useState(true);
  const [editando, setEditando] = useState(null);
  const [prods, setProds] = useState([]);
  const [loadSave, setLoadSave] = useState(false);

  // Filtros de Data
  const [filtroData, setFiltroData] = useState("todos"); 
  const [dataDe, setDataDe]         = useState("");
  const [dataAte, setDataAte]       = useState("");

  const carregar = useCallback(async () => {
    setLoading(true);
    let query = supabase.from("vendas").select("*");
    
    if (filtroData === "hoje") {
      query = query.eq("data_venda", hoje());
    } else if (filtroData === "mes") {
      const inicioMes = new Date();
      inicioMes.setDate(1);
      query = query.gte("data_venda", inicioMes.toISOString().split("T")[0]);
    } else if (filtroData === "custom") {
      if (dataDe)  query = query.gte("data_venda", dataDe);
      if (dataAte) query = query.lte("data_venda", dataAte);
    }

    const { data: v } = await query.order("data_venda", { ascending: false }).order("created_at", { ascending: false }).limit(500);
    const { data: p } = await supabase.from("produtos").select("*").eq("ativo", true);
    setVendas(v || []); setProds(p || []); setLoading(false);
  }, [filtroData, dataDe, dataAte]);

  useEffect(() => {
    carregar();
    const ch = supabase.channel("vendas_hist").on("postgres_changes", { event: "*", schema: "public", table: "vendas" }, carregar).subscribe();
    return () => supabase.removeChannel(ch);
  }, [carregar]);

  const filtradas = vendas.filter(v => !filtro || v.comprador.toLowerCase().includes(filtro.toLowerCase()) || v.produto_nome.toLowerCase().includes(filtro.toLowerCase()));
  const STATUS_COLOR = { "Pago e Entregue": C.green, "Pendente": C.gold, "Pago Aguardando": C.teal, "Reembolsado": C.red };

  const exportarExcel = () => {
    // 1. Dados da Planilha de Vendas
    const dadosVendas = filtradas.map(v => ({
      "DATA": v.data_venda?.split("-").reverse().join("/"),
      "PRODUTO": v.produto_nome,
      "COMPRADOR": v.comprador,
      "TURMA": v.turma || "—",
      "QTD": v.quantidade,
      "VALOR UNIT.": v.preco_venda,
      "TOTAL": v.preco_venda * v.quantidade,
      "PAGAMENTO": v.forma_pagamento,
      "STATUS": v.status
    }));

    // 2. Dados da Capa (Introdução Branded)
    const periodoStr = filtroData === "todos" ? "Todo o Histórico" : 
                      filtroData === "hoje" ? `Hoje (${hoje().split("-").reverse().join("/")})` :
                      filtroData === "mes" ? "Mês Atual" : `Personalizado (${dataDe} até ${dataAte})`;

    const dadosCapa = [
      ["DIRETÓRIO ACADÊMICO CLEUSA FERRI - UAM SJC"],
      ["GESTÃO 2026"],
      [""],
      ["RELATÓRIO OFICIAL DE VENDAS"],
      [""],
      ["RESUMO DO DOCUMENTO:"],
      ["PERÍODO SELECIONADO:", periodoStr],
      ["GERADO EM:", new Date().toLocaleString("pt-BR")],
      [""],
      ["MÉTRICAS DO PERÍODO:"],
      ["Total de Itens Vendidos:", filtradas.reduce((s,v) => s + v.quantidade, 0)],
      ["Nº Total de Pedidos:", filtradas.length],
      ["Receita Bruta Total:", fmtR(filtradas.reduce((s, v) => s + (v.preco_venda * v.quantidade), 0))],
      [""],
      ["------------------------------------------------------------"],
      ["ESTE É UM DOCUMENTO AUTOMATIZADO - D.A. CLEUSA FERRI 2026"]
    ];

    // 3. Criação do Livro e Abas
    const wb = XLSX.utils.book_new();
    const wsVendas = XLSX.utils.json_to_sheet(dadosVendas);
    const wsCapa = XLSX.utils.aoa_to_sheet(dadosCapa);

    // Ajuste de largura das colunas para UX
    wsVendas['!cols'] = [
      { wch: 12 }, { wch: 30 }, { wch: 30 }, { wch: 10 }, 
      { wch: 8 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 20 }
    ];
    wsCapa['!cols'] = [{ wch: 35 }, { wch: 50 }];

    XLSX.utils.book_append_sheet(wb, wsCapa, "CAPA");
    XLSX.utils.book_append_sheet(wb, wsVendas, "LISTAGEM DE VENDAS");

    // 4. Download com Nome Profissional
    const timestamp = new Date().toISOString().slice(0,10);
    const fileName = `Relatorio_Vendas_DA_CleusaFerri_${timestamp}.xlsx`;
    XLSX.writeFile(wb, fileName);
  };

  const salvarEdicao = async (e) => {
    e.preventDefault(); setLoadSave(true);
    const vAntiga = vendas.find(v => v.id === editando.id);
    await supabase.rpc("ajustar_estoque", { p_id: vAntiga.produto_id, p_qtd: vAntiga.quantidade });
    const prodNovo = prods.find(p => p.id === editando.produto_id);
    const { data: resEstoque } = await supabase.rpc("ajustar_estoque", { p_id: editando.produto_id, p_qtd: -editando.quantidade });
    if (resEstoque?.error) { alert("Erro de estoque: " + resEstoque.error); setLoadSave(false); return; }
    const { error } = await supabase.from("vendas").update({
      produto_id: editando.produto_id, produto_nome: prodNovo.nome, comprador: editando.comprador,
      turma: editando.turma, quantidade: editando.quantidade, preco_venda: prodNovo.preco_normal,
      forma_pagamento: editando.forma_pagamento, status: editando.status, data_venda: editando.data_venda
    }).eq("id", editando.id);
    if (error) alert("Erro ao salvar: " + error.message);
    else { setEditando(null); carregar(); }
    setLoadSave(false);
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24, gap: 16, flexWrap: 'wrap' }}>
        <h1 style={{ fontFamily: "Syne", fontSize: 24, fontWeight: 800 }}>🗂️ Histórico</h1>
        
        <div style={{ display: "flex", gap: 10, alignItems: "end", flexWrap: 'wrap' }}>
          <Select label="Período" value={filtroData} onChange={e => setFiltroData(e.target.value)} style={{ width: 140 }}>
            <option value="todos">Todos os dados</option>
            <option value="hoje">Hoje</option>
            <option value="mes">Mês Atual</option>
            <option value="custom">Personalizado</option>
          </Select>
          {filtroData === "custom" && (
            <>
              <Input type="date" label="De" value={dataDe} onChange={e => setDataDe(e.target.value)} />
              <Input type="date" label="Até" value={dataAte} onChange={e => setDataAte(e.target.value)} />
            </>
          )}
          <Btn variant="ghost" onClick={exportarExcel} style={{ height: 40, display: 'flex', alignItems: 'center', gap: 8, background: C.green + '33', color: C.green, border: `1px solid ${C.green}44` }}>
            <span>📥</span> Exportar Excel
          </Btn>
        </div>
      </div>
      
      <Card>
        <Input placeholder="Buscar por comprador ou produto..." value={filtro} onChange={e => setFiltro(e.target.value)} style={{ marginBottom: 16 }} />
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, minWidth: 700 }}>
            <thead>
              <tr style={{ borderBottom: `1px solid ${C.border}` }}>
                {["Data","Produto","Comprador","Turma","Qtd","Preço","Pagamento","Status","Ações"].map(h => (
                  <th key={h} style={{ padding: "8px 12px", color: C.muted, fontWeight: 500, textAlign: "left", whiteSpace: "nowrap" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtradas.map((v, i) => (
                <tr key={v.id} style={{ borderBottom: `1px solid ${C.border}22`, background: i % 2 === 0 ? "transparent" : "#ffffff04" }}>
                  <td style={{ padding: "9px 12px", color: C.muted, whiteSpace: "nowrap" }}>{v.data_venda?.split("-").reverse().join("/")}</td>
                  <td style={{ padding: "9px 12px", fontWeight: 500 }}>{v.produto_nome}</td>
                  <td style={{ padding: "9px 12px" }}>{v.comprador}</td>
                  <td style={{ padding: "9px 12px", color: C.muted }}>{v.turma || "—"}</td>
                  <td style={{ padding: "9px 12px", textAlign: "center" }}>{v.quantidade}</td>
                  <td style={{ padding: "9px 12px" }}>{fmtR(v.preco_venda * v.quantidade)}</td>
                  <td style={{ padding: "9px 12px", color: C.muted }}>{v.forma_pagamento}</td>
                  <td style={{ padding: "9px 12px" }}><Badge color={STATUS_COLOR[v.status] || C.muted}>{v.status}</Badge></td>
                  <td style={{ padding: "9px 12px" }}>
                    <div style={{ display: "flex", gap: 8 }}>
                      <button onClick={() => {
                        const msg = `Olá ${v.comprador}, o D.A. Cleusa Ferri confirma sua compra de ${v.quantidade}x ${v.produto_nome}. Total: ${fmtR(v.preco_venda * v.quantidade)}. Obrigado!`;
                        window.open(`https://wa.me/?text=${encodeURIComponent(msg)}`, '_blank');
                      }} style={{ background: "none", border: "none", fontSize: 16 }} title="Enviar Recibo WhatsApp">📱</button>
                      <button onClick={() => setEditando({ ...v })} style={{ background: "none", border: "none", fontSize: 16 }}>✏️</button>
                      <button onClick={async () => { if (confirm(`Excluir?`)) { await supabase.rpc("ajustar_estoque", { p_id: v.produto_id, p_qtd: v.quantidade }); await supabase.from("vendas").delete().eq("id", v.id); carregar(); } }} style={{ background: "none", border: "none", fontSize: 16 }}>🗑️</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
      {editando && (
        <div style={{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,0.8)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 16 }}>
          <Card style={{ width: "100%", maxWidth: 500 }}>
            <h2 style={{ fontFamily: "Syne", marginBottom: 20 }}>Editar Venda</h2>
            <form onSubmit={salvarEdicao} style={{ display: "flex", flexDirection: "column", gap: 12 }}>
              <Input label="Comprador" value={editando.comprador} onChange={e => setEditando({...editando, comprador: e.target.value})} required />
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                <Input label="Turma" value={editando.turma} onChange={e => setEditando({...editando, turma: e.target.value})} />
                <Input label="Data" type="date" value={editando.data_venda} onChange={e => setEditando({...editando, data_venda: e.target.value})} required />
              </div>
              <Select label="Produto" value={editando.produto_id} onChange={e => setEditando({...editando, produto_id: e.target.value})}>{prods.map(p => <option key={p.id} value={p.id}>{p.nome}</option>)}</Select>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                <Input label="Qtd" type="number" value={editando.quantidade} onChange={e => setEditando({...editando, quantidade: parseInt(e.target.value) || 1})} required />
                <Select label="Pagamento" value={editando.forma_pagamento} onChange={e => setEditando({...editando, forma_pagamento: e.target.value})}>{["Pix","Dinheiro","Cartão","Transferência"].map(o => <option key={o} value={o}>{o}</option>)}</Select>
              </div>
              <Select label="Status" value={editando.status} onChange={e => setEditando({...editando, status: e.target.value})}>{Object.keys(STATUS_COLOR).map(s => <option key={s} value={s}>{s}</option>)}</Select>
              <div style={{ display: "flex", gap: 10, marginTop: 10 }}>
                <Btn variant="ghost" onClick={() => setEditando(null)} style={{ flex: 1 }}>Cancelar</Btn>
                <Btn type="submit" variant="success" disabled={loadSave} style={{ flex: 1 }}>{loadSave ? "Salvando..." : "Salvar"}</Btn>
              </div>
            </form>
          </Card>
        </div>
      )}
    </div>
  );
}

// ─── FINANCEIRO ─────────────────────────────────────────────────
function Financeiro() {
  const [caixas, setCaixas]     = useState([]);
  const [movimentos, setMovs]   = useState([]);
  const [vendasHoje, setVendas] = useState([]);
  const [loading, setLoading]   = useState(true);
  
  // Modais
  const [showFechar, setShow] = useState(false);
  const [showGasto, setShowGasto] = useState(false);
  
  // Forms
  const [conferido, setConf]  = useState({ dinheiro: "", pix: "", observacao: "" });
  const [gasto, setGasto]     = useState({ valor: "", descricao: "", tipo: "saida", status: "Pago", id: null });

  const carregar = useCallback(async () => {
    setLoading(true);
    const dHoje = hoje();
    
    const [{ data: c }, { data: m }, { data: v }] = await Promise.all([
      supabase.from("caixas").select("*").order("data_fechamento", { ascending: false }).limit(30),
      supabase.from("movimentacoes_financeiras").select("*").order("created_at", { ascending: false }).limit(50),
      supabase.from("vendas").select("*").eq("data_venda", dHoje).eq("status", "Pago e Entregue")
    ]);

    setCaixas(c || []);
    setMovs(m || []);
    setVendas(v || []);
    setLoading(false);
  }, []);

  useEffect(() => { carregar(); }, [carregar]);

  const totalSistema = {
    dinheiro: vendasHoje.filter(v => v.forma_pagamento === "Dinheiro").reduce((s, v) => s + (v.preco_venda * v.quantidade), 0),
    pix: vendasHoje.filter(v => v.forma_pagamento === "Pix").reduce((s, v) => s + (v.preco_venda * v.quantidade), 0)
  };

  const fecharCaixa = async (e) => {
    e.preventDefault();
    const payload = {
      data_fechamento: hoje(),
      valor_esperado_dinheiro: totalSistema.dinheiro,
      valor_informado_dinheiro: Number(conferido.dinheiro) || 0,
      valor_esperado_pix: totalSistema.pix,
      valor_informado_pix: Number(conferido.pix) || 0,
      observacao: conferido.observacao
    };

    const { error } = await supabase.from("caixas").insert(payload);
    if (error) alert("Erro ao fechar caixa: " + error.message);
    else {
      setShow(false); setConf({ dinheiro: "", pix: "", observacao: "" });
      carregar();
    }
  };

  const salvarGasto = async (e) => {
    e.preventDefault();
    if (!gasto.valor || !gasto.descricao) return alert("Preencha valor e descrição!");
    
    const payload = {
      tipo: gasto.tipo,
      valor: Number(gasto.valor),
      descricao: gasto.descricao,
      status: gasto.status || "Pago"
    };

    const { error } = gasto.id 
      ? await supabase.from("movimentacoes_financeiras").update(payload).eq("id", gasto.id)
      : await supabase.from("movimentacoes_financeiras").insert(payload);

    if (error) alert("Erro ao salvar: " + error.message);
    else {
      setShowGasto(false); setGasto({ valor: "", descricao: "", tipo: "saida", status: "Pago", id: null });
      carregar();
    }
  };

  const excluirMov = async (id) => {
    if (confirm("Excluir este registro permanentemente?")) {
      const { error } = await supabase.from("movimentacoes_financeiras").delete().eq("id", id);
      if (error) alert(error.message);
      else carregar();
    }
  };

  return (
    <div style={{ maxWidth: 1000, margin: '0 auto' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 24, flexWrap: 'wrap', gap: 16 }}>
        <h1 style={{ fontFamily: "Syne", fontSize: 24, fontWeight: 800 }}>💸 Financeiro e Caixa</h1>
        <div style={{ display: 'flex', gap: 10 }}>
          <Btn variant="ghost" onClick={() => { setGasto({ valor: "", descricao: "", tipo: "saida", status: "Pago", id: null }); setShowGasto(true); }}>🧾 Registrar Gasto</Btn>
          <Btn variant="success" onClick={() => setShow(true)}>🔒 Fechar Caixa do Dia</Btn>
        </div>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1.5fr', gap: 20, marginBottom: 24 }}>
        <Card>
          <h3 style={{ fontSize: 14, color: C.muted, textTransform: 'uppercase', marginBottom: 16 }}>Esperado em Caixa (Hoje)</h3>
          <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 8 }}>
            <span>💵 Dinheiro:</span>
            <span style={{ fontWeight: 700, color: C.green }}>{fmtR(totalSistema.dinheiro)}</span>
          </div>
          <div style={{ display: 'flex', justifyContent: 'space-between' }}>
            <span>📱 Pix/Transferência:</span>
            <span style={{ fontWeight: 700, color: C.teal }}>{fmtR(totalSistema.pix)}</span>
          </div>
        </Card>

        <Card style={{ background: C.navy + '33' }}>
          <h3 style={{ fontSize: 14, color: C.muted, textTransform: 'uppercase', marginBottom: 16 }}>Últimas Movimentações</h3>
          <div style={{ maxHeight: 200, overflowY: 'auto' }}>
            {movimentos.map(m => (
              <div key={m.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 0', borderBottom: `1px solid ${C.border}22`, fontSize: 13 }}>
                <div style={{ flex: 1 }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <span style={{ color: C.muted, fontSize: 10 }}>{new Date(m.created_at).toLocaleDateString()}</span>
                    <Badge color={m.status === 'Sinal/Parcial' ? C.gold : m.status === 'Agendado' ? C.purple : C.red}>{m.status || 'Pago'}</Badge>
                  </div>
                  <div style={{ fontWeight: 500, marginTop: 2 }}>{m.descricao}</div>
                </div>
                <div style={{ textAlign: 'right', display: 'flex', alignItems: 'center', gap: 12 }}>
                  <span style={{ color: C.red, fontWeight: 700 }}>- {fmtR(m.valor)}</span>
                  <div style={{ display: 'flex', gap: 4 }}>
                    <button onClick={() => { setGasto({ ...m }); setShowGasto(true); }} style={{ background: 'none', border: 'none', cursor: 'pointer' }}>✏️</button>
                    <button onClick={() => excluirMov(m.id)} style={{ background: 'none', border: 'none', cursor: 'pointer' }}>🗑️</button>
                  </div>
                </div>
              </div>
            ))}
            {movimentos.length === 0 && <p style={{ color: C.muted, fontSize: 12 }}>Nenhum gasto registrado.</p>}
          </div>
        </Card>
      </div>

      <Card>
        <h3 style={{ fontSize: 15, fontWeight: 700, marginBottom: 16 }}>Histórico de Fechamentos Diários</h3>
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13, minWidth: 600 }}>
            <thead>
              <tr style={{ borderBottom: `1px solid ${C.border}` }}>
                {["Data", "Dinheiro (Inf/Esp)", "Pix (Inf/Esp)", "Diferença", "Status"].map(h => (
                  <th key={h} style={{ padding: '10px', textAlign: 'left', color: C.muted }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {caixas.map(c => {
                const dif = (c.valor_informado_dinheiro + c.valor_informado_pix) - (c.valor_esperado_dinheiro + c.valor_esperado_pix);
                return (
                  <tr key={c.id} style={{ borderBottom: `1px solid ${C.border}22` }}>
                    <td style={{ padding: '10px' }}>{c.data_fechamento?.split('-').reverse().join('/')}</td>
                    <td style={{ padding: '10px' }}>{fmtR(c.valor_informado_dinheiro)} / {fmtR(c.valor_esperado_dinheiro)}</td>
                    <td style={{ padding: '10px' }}>{fmtR(c.valor_informado_pix)} / {fmtR(c.valor_esperado_pix)}</td>
                    <td style={{ padding: '10px', color: dif === 0 ? C.green : C.red, fontWeight: 700 }}>{fmtR(dif)}</td>
                    <td style={{ padding: '10px' }}><Badge color={dif === 0 ? C.green : C.red}>{dif === 0 ? "OK" : "Divergência"}</Badge></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>

      {/* Modal Fechamento */}
      {showFechar && (
        <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.85)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000, padding: 16 }}>
          <Card style={{ width: '100%', maxWidth: 450 }}>
            <h2 style={{ fontFamily: 'Syne', marginBottom: 20 }}>🔒 Fechar Caixa do Dia</h2>
            <form onSubmit={fecharCaixa} style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
              <Input label="Total em Dinheiro (Físico)" type="number" step="0.01" value={conferido.dinheiro} onChange={e => setConf({...conferido, dinheiro: e.target.value})} required placeholder="0,00" />
              <Input label="Total em Pix (Conferido no extrato)" type="number" step="0.01" value={conferido.pix} onChange={e => setConf({...conferido, pix: e.target.value})} required placeholder="0,00" />
              <Input label="Observações" value={conferido.observacao} onChange={e => setConf({...conferido, observacao: e.target.value})} placeholder="Ex: Faltou troco..." />
              <div style={{ background: C.bg, padding: 12, borderRadius: 8, fontSize: 13 }}>
                <p><strong>Esperado em Dinheiro:</strong> {fmtR(totalSistema.dinheiro)}</p>
                <p><strong>Esperado em Pix:</strong> {fmtR(totalSistema.pix)}</p>
              </div>
              <div style={{ display: 'flex', gap: 10 }}>
                <Btn variant="ghost" onClick={() => setShow(false)} style={{ flex: 1 }}>Cancelar</Btn>
                <Btn type="submit" variant="success" style={{ flex: 1 }}>Confirmar</Btn>
              </div>
            </form>
          </Card>
        </div>
      )}

      {/* Modal Novo Gasto */}
      {showGasto && (
        <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.85)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000, padding: 16 }}>
          <Card style={{ width: '100%', maxWidth: 450 }}>
            <h2 style={{ fontFamily: 'Syne', marginBottom: 20 }}>{gasto.id ? '✏️ Editar Gasto' : '🧾 Registrar Gasto'}</h2>
            <form onSubmit={salvarGasto} style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
              <Input label="Valor (R$)" type="number" step="0.01" value={gasto.valor} onChange={e => setGasto({...gasto, valor: e.target.value})} required placeholder="0,00" />
              <Input label="Descrição / Motivo" value={gasto.descricao} onChange={e => setGasto({...gasto, descricao: e.target.value})} required placeholder="Ex: Pagamento Fornecedor" />
              
              <Select label="Status do Pagamento" value={gasto.status || "Pago"} onChange={e => setGasto({...gasto, status: e.target.value})}>
                <option value="Pago">Pago (Total)</option>
                <option value="Sinal/Parcial">Sinal / Parcial</option>
                <option value="Agendado">Agendado / Pendente</option>
              </Select>

              <div style={{ display: 'flex', gap: 10, marginTop: 10 }}>
                <Btn variant="ghost" onClick={() => setShowGasto(false)} style={{ flex: 1 }}>Cancelar</Btn>
                <Btn type="submit" variant="success" style={{ flex: 1 }}>{gasto.id ? 'Salvar Alterações' : 'Registrar'}</Btn>
              </div>
            </form>
          </Card>
        </div>
      )}
    </div>
  );
}

// ─── APP PRINCIPAL ───────────────────────────────────────────────
export default function App() {
  const [session, setSession] = useState(null);
  const [tab, setTab]         = useState("dashboard");
  const [checking, setChecking] = useState(true);

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => { setSession(session); setChecking(false); });
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_, s) => setSession(s));
    return () => subscription.unsubscribe();
  }, []);

  const logout = async () => { await supabase.auth.signOut(); setSession(null); setTab("dashboard"); };
  if (checking) return <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: C.bg }}><div style={{ color: C.muted, fontSize: 32 }}>⏳</div></div>;
  if (!session) return <Login onLogin={() => {}} />;
  const PAGES = { 
    dashboard: Dashboard, 
    vendas: NovaVenda, 
    produtos: Produtos, 
    kits: Kits,
    historico: Historico, 
    financeiro: Financeiro 
  };
  const Page  = PAGES[tab] || Dashboard;

  return (
    <>
      <style>{globalCss}</style>
      <Sidebar tab={tab} setTab={setTab} onLogout={logout} user={session.user} />
      <main style={{ marginLeft: 220, padding: 32, minHeight: "100vh" }}>
        <Page />
      </main>
    </>
  );
}
