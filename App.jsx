import { useState, useEffect, useRef, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, RadarChart, PolarGrid, PolarAngleAxis, PolarRadiusAxis, Radar, Cell, PieChart, Pie } from "recharts";
import { supabase } from "./supabaseClient";

// ═══════════════════════════════════════
// IMPROVEMENTS KNOWLEDGE BASE
// ═══════════════════════════════════════
const IMPROVEMENTS = {
  "Direct Materials": [
    "Implement strategic sourcing with 3-year contracts and volume consolidation across sites to achieve 8-15% savings on raw materials.",
    "Develop dual/multi-source strategies for critical materials to reduce dependency and improve negotiation leverage.",
    "Adopt should-cost modeling — decompose supplier pricing into raw material, conversion, and margin components.",
    "Establish vendor-managed inventory (VMI) programs to reduce carrying costs while maintaining supply security."
  ],
  "IT & Technology": [
    "Rationalize SaaS portfolio — most enterprises have 30-40% redundant licenses. Conduct full audit and eliminate overlap.",
    "Negotiate enterprise-wide agreements with cloud providers to unlock tier pricing — typically 20-35% below on-demand.",
    "Consolidate IT service desk contracts under fewer vendors with outcome-based pricing rather than T&M.",
    "Implement FinOps practices — tag resources, set auto-scaling, establish showback/chargeback to business units."
  ],
  "Facilities & Real Estate": [
    "Transition to flexible workspace models — subleasing underutilized floors can offset 15-25% of total rent.",
    "Implement energy management and LED retrofits — typical ROI within 18 months with 20-30% utilities reduction.",
    "Consolidate FM contracts to integrated facility management providers for 10-18% savings.",
    "Renegotiate lease terms during renewal windows — current conditions often favor 5-10% rent reductions."
  ],
  "HR & Workforce": [
    "Implement MSP program for contingent labor — standardize rates and reduce markups by 15-25%.",
    "Consolidate benefits brokers and conduct annual market checks — switching can save 8-12% annually.",
    "Centralize recruitment through RPO to reduce cost-per-hire by 20-30% vs agency-by-agency.",
    "Move training to blended learning (60% digital) — reducing per-employee costs by 35-50%."
  ],
  "Marketing & Advertising": [
    "Conduct agency fee benchmarking — most companies overpay by 15-20% on retainers.",
    "Shift 20-30% of traditional media to performance-based digital with measurable ROAS.",
    "Consolidate creative production under fewer agencies or bring in-house for 25-40% reduction.",
    "Implement programmatic buying — AI-driven optimization improves cost efficiency by 15-25%."
  ],
  "Professional Services": [
    "Establish rate cards by tier (Partner, Manager, Analyst) benchmarked against market data.",
    "Shift from T&M to fixed-fee or outcome-based contracts for recurring engagements.",
    "Bundle audit, tax, and advisory with one firm where allowed — typically saves 10-15%.",
    "Define clear deliverables and milestones for consulting to prevent scope creep."
  ],
  "Travel & Expenses": [
    "Negotiate corporate hotel programs with 15-25% below BAR — concentrate into 3-5 chains.",
    "Implement pre-trip approval and dynamic travel policy — route expensive bookings for approval.",
    "Consolidate ground transport under managed programs — reduces fragmented spending.",
    "Adopt virtual meeting-first for internal travel — each avoided trip saves $1,500-$3,000."
  ],
  "Logistics & Distribution": [
    "Conduct annual freight RFPs with lane-level bidding — typical savings 10-18% on contracted rates.",
    "Optimize warehouse network — model total landed cost for consolidation opportunities saving 12-20%.",
    "Implement TMS for load consolidation and route optimization — reduces freight costs by 8-15%.",
    "Shift parcel to regional carriers for last-mile — 15-25% cheaper than national for ground."
  ],
  "Procurement Operations": [
    "Implement P2P automation (Coupa, Ariba) — 60-70% reduction in PO cycle time.",
    "Establish supplier management office with scorecards for strategic relationships.",
    "Consolidate contract management on CLM platform — reduces value leakage by 5-10%.",
    "Build spend analytics capability — visibility across ERPs enables systematic savings."
  ],
  "MRO & Maintenance": [
    "Implement tail-spend management — consolidate low-value purchases through marketplace solutions.",
    "Transition to preventive/predictive maintenance — reduces unplanned downtime by 30-50%.",
    "Standardize spare parts across sites — reduce SKU proliferation by 25-40%.",
    "Establish consignment inventory for critical spares — pay on consumption, reduce working capital."
  ],
  "R&D Services": [
    "Consolidate CRO into strategic partnerships with 2-3 providers — volume unlocks 10-20% reductions.",
    "Implement preferred supplier agreements for lab supplies — GPO membership saves 15-25%.",
    "Centralize regulatory consulting under framework agreements — avoid 30-50% premium ad-hoc rates.",
    "Negotiate IP/patent fees based on portfolio volume — 50+ filings should get 15-20% below standard."
  ],
  "Capital Expenditure": [
    "Implement TCO analysis for equipment over $100K — include installation, maintenance, energy, disposal.",
    "Negotiate extended warranties at purchase — bundled deals are 20-30% cheaper than after-market.",
    "Standardize equipment specs across facilities — improves maintenance efficiency by 15-25%.",
    "Establish lease vs buy frameworks — for rapidly depreciating tech, leasing reduces cost by 10-15%."
  ]
};

const COLORS = {
  navy: "#0a1628", blue: "#2563eb", lightBlue: "#eff6ff", accent: "#10b981", accentDark: "#059669",
  gold: "#f59e0b", red: "#ef4444", green: "#22c55e", bg: "#f8fafc", card: "#ffffff",
  text: "#0f172a", textSecondary: "#475569", muted: "#94a3b8", border: "#e2e8f0",
  leader: "#22c55e", median: "#2563eb", laggard: "#ef4444",
};

const AGENT_SYSTEM = `You are ProcureBench AI — a senior procurement strategy advisor with 20+ years of experience across Pharmaceutical, Retail, Consumer Goods, and Technology industries.

YOUR ROLE:
- Expert-level analysis of spend benchmarking results with actionable insights
- Contextualize within industry dynamics (pharma GxP, retail omnichannel, FMCG private label, tech R&D)
- Concrete improvement levers with savings ranges and timelines
- Guide users step-by-step through the tool

KEY CONCEPTS:
- Spend as % of Revenue = (Category Spend / Total Revenue) x 100
- Leader (25th %ile): Most efficient. Laggard (75th %ile): Least efficient
- 1 percentage point on $10B revenue = $100M savings opportunity

RESPONSE STYLE:
- Structure: Position assessment → Root cause → Actions with savings → Quick wins vs strategic
- Reference actual benchmark numbers
- Use industry-specific language
- 200-300 words on analysis, shorter for guidance
- Markdown: **bold**, bullet points
- Prioritize by savings opportunity`;

// ═══════════════════════════════════════
// UTILITY FUNCTIONS
// ═══════════════════════════════════════
function normalizeStr(s) {
  return (s || "").toLowerCase().replace(/[^a-z0-9 ]/g, " ").replace(/\s+/g, " ").trim();
}

function matchCategory(rawCategory, industry, taxonomy) {
  const norm = normalizeStr(rawCategory);
  if (!norm) return { l1: null, l2: null, confidence: 0, matches: [] };
  
  const l1Aliases = {
    "Direct Materials": ["direct material", "raw material purchases", "materials"],
    "IT & Technology": ["information technology", "it technology", "technology", "it services"],
    "Facilities & Real Estate": ["facility costs", "real estate", "facilities", "refm", "re fm"],
    "HR & Workforce": ["people hr", "human resources", "hr workforce", "people workforce", "hr"],
    "Marketing & Advertising": ["marketing spend", "marketing advertising", "marketing", "advertising"],
    "Professional Services": ["consulting professional", "professional services", "consulting"],
    "Travel & Expenses": ["t e", "travel expenses", "travel entertainment", "travel"],
    "Logistics & Distribution": ["supply chain logistics", "supply chain", "logistics distribution", "logistics"],
    "Procurement Operations": ["procurement dept", "procurement operations", "procurement"],
    "MRO & Maintenance": ["plant maintenance", "mro maintenance", "mro", "maintenance"],
    "R&D Services": ["r d outsourcing", "r d services", "research development", "r d"],
    "Capital Expenditure": ["capex", "capital expenditure", "capital expense", "cap ex"],
  };
  
  for (const [l1Name, aliases] of Object.entries(l1Aliases)) {
    for (const alias of aliases) {
      if (norm === alias || norm.includes(alias) || alias.includes(norm))
        return { l1: l1Name, l2: null, confidence: 99, reason: "L1 match", matches: [{ l1: l1Name, l2: null, confidence: 99, reason: "L1 match" }] };
    }
  }
  
  const l1Names = [...new Set(taxonomy.map(t => t.standard_l1))];
  for (const l1Name of l1Names) {
    const normL1 = normalizeStr(l1Name);
    if (norm === normL1 || norm.includes(normL1) || normL1.includes(norm))
      return { l1: l1Name, l2: null, confidence: 99, reason: "L1 exact", matches: [{ l1: l1Name, l2: null, confidence: 99, reason: "L1 exact" }] };
    const nw = norm.split(" ").filter(w => w.length > 2);
    const lw = normL1.split(" ").filter(w => w.length > 2);
    const ov = nw.filter(w => lw.some(l => l.includes(w) || w.includes(l))).length;
    if (ov >= 2 && ov >= Math.min(nw.length, lw.length) * 0.5)
      return { l1: l1Name, l2: null, confidence: 95, reason: "L1 word match", matches: [{ l1: l1Name, l2: null, confidence: 95, reason: "L1 word match" }] };
  }
  
  const candidates = [];
  for (const t of taxonomy) {
    if (t.industry_applicability !== "All" && t.industry_applicability !== industry) continue;
    const aliases = (t.common_aliases || "").split(",").map(a => normalizeStr(a));
    const keywords = (t.keywords || "").split(",").map(k => k.trim().toLowerCase());
    for (const alias of aliases) {
      if (alias && norm === alias) candidates.push({ l1: t.standard_l1, l2: t.standard_l2, confidence: 98, reason: "Exact alias" });
    }
    for (const alias of aliases) {
      if (alias && alias.length > 3 && (norm.includes(alias) || alias.includes(norm)))
        if (!candidates.some(c => c.l1 === t.standard_l1 && c.l2 === t.standard_l2 && c.confidence >= 85))
          candidates.push({ l1: t.standard_l1, l2: t.standard_l2, confidence: 85, reason: "Partial alias" });
    }
    let kwHits = 0;
    const normWords = norm.split(" ");
    for (const kw of keywords) { if (kw && normWords.some(w => w.includes(kw) || kw.includes(w))) kwHits++; }
    if (kwHits > 0) candidates.push({ l1: t.standard_l1, l2: t.standard_l2, confidence: Math.min(80, 40 + kwHits * 15), reason: `${kwHits} keywords` });
  }
  const seen = new Set();
  const unique = [];
  candidates.sort((a, b) => b.confidence - a.confidence);
  for (const c of candidates) { const k = `${c.l1}|${c.l2}`; if (!seen.has(k)) { seen.add(k); unique.push(c); } }
  const top = unique[0] || { l1: null, l2: null, confidence: 0, reason: "No match" };
  return { l1: top.l1, l2: top.l2, confidence: top.confidence, reason: top.reason, matches: unique.slice(0, 5) };
}

function getPosition(pct, le, m, la) {
  if (pct <= le) return { label: "Leader", color: COLORS.leader, desc: "Top quartile — best-in-class" };
  if (pct <= m) return { label: "Above Median", color: COLORS.accent, desc: "Better than average" };
  if (pct <= la) return { label: "Below Median", color: COLORS.gold, desc: "Significant savings opportunity" };
  return { label: "Laggard", color: COLORS.red, desc: "Urgent optimization needed" };
}

function parseSpend(val) {
  if (typeof val === "number") return val;
  if (!val) return 0;
  const s = String(val).replace(/[$,\s]/g, "").toUpperCase();
  const mult = s.includes("B") ? 1e9 : s.includes("M") ? 1e6 : s.includes("K") ? 1e3 : 1;
  return parseFloat(s.replace(/[BMKT]/g, "")) * mult || 0;
}

// Extract a 4-digit calendar year from any format
function extractYear(val) {
  if (!val && val !== 0) return null;
  
  // If it's a number that looks like a year directly (2020-2030)
  if (typeof val === "number" && val >= 2000 && val <= 2099) return val;
  
  // If it's an Excel date serial number (e.g., 44958 = some date in 2023)
  if (typeof val === "number" && val > 40000 && val < 60000) {
    const d = new Date((val - 25569) * 86400000);
    return d.getFullYear();
  }
  
  // If it's a Date object
  if (val instanceof Date) return val.getFullYear();
  
  const s = String(val).trim();
  
  // Fiscal year: "2022-23", "22-23", "FY2022-23", "FY22/23"
  const fyMatch = s.match(/(\d{2,4})[\-\/](\d{2,4})/);
  if (fyMatch) {
    let endYear = parseInt(fyMatch[2]);
    // If 2-digit, convert: "23" → 2023
    if (endYear < 100) endYear += 2000;
    if (endYear >= 2000 && endYear <= 2099) return endYear;
  }
  
  // Direct 4-digit year: "2023", "FY2023", "CY2023"
  const y4 = s.match(/\b(20\d{2})\b/);
  if (y4) return parseInt(y4[1]);
  
  // Date formats: "2/21/23", "02/21/2023", "21-Feb-2023", "2023-02-21", etc.
  // Try parsing as date
  const dateAttempt = new Date(s);
  if (!isNaN(dateAttempt.getTime()) && dateAttempt.getFullYear() >= 2000 && dateAttempt.getFullYear() <= 2099) {
    return dateAttempt.getFullYear();
  }
  
  // Last resort: find any 2-digit number that could be a year (e.g., "Q3 23")
  const y2 = s.match(/\b(\d{2})\b/);
  if (y2) {
    const yr = parseInt(y2[1]);
    if (yr >= 18 && yr <= 30) return 2000 + yr;
  }
  
  return null;
}

// ═══════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════
export default function App() {
  const [user, setUser] = useState(null);
  const [view, setView] = useState("quick");
  const [contextMsgs, setContextMsgs] = useState([]);
  const [benchmarks, setBenchmarks] = useState([]);
  const [l1Summary, setL1Summary] = useState([]);
  const [taxonomy, setTaxonomy] = useState([]);
  const [loading, setLoading] = useState(true);
  const [dbError, setDbError] = useState(null);

  const addAgentMsg = useCallback((msg) => setContextMsgs(prev => [...prev, msg]), []);

  useEffect(() => {
    async function fetchData() {
      try {
        if (!supabase) throw new Error("Database not configured");
        const [bRes, lRes, tRes] = await Promise.all([
          supabase.from("benchmarks").select("*"),
          supabase.from("l1_summary").select("*"),
          supabase.from("taxonomy").select("*"),
        ]);
        if (bRes.error) throw bRes.error;
        setBenchmarks(bRes.data || []);
        setL1Summary(lRes.data || []);
        setTaxonomy(tRes.data || []);
      } catch (err) {
        setDbError(err.message);
      }
      setLoading(false);
    }
    fetchData();
  }, []);

  if (!user) return <Landing onLogin={setUser} />;
  if (loading) return <LoadingScreen />;
  if (dbError) return <ErrorScreen error={dbError} />;

  return (
    <div style={{ display: "flex", height: "100vh", fontFamily: "'Inter', -apple-system, sans-serif", background: COLORS.bg, color: COLORS.text, overflow: "hidden" }}>
      <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>
        <Header user={user} view={view} setView={setView} onLogout={() => setUser(null)} />
        <div style={{ flex: 1, padding: "32px 40px", overflow: "auto" }}>
          {view === "quick" && <QuickLookup benchmarks={benchmarks} l1Summary={l1Summary} onAgentMsg={addAgentMsg} />}
          {view === "upload" && <ExcelUpload benchmarks={benchmarks} l1Summary={l1Summary} taxonomy={taxonomy} onAgentMsg={addAgentMsg} />}
        </div>
      </div>
      <AgentPanel contextMsgs={contextMsgs} />
    </div>
  );
}

function LoadingScreen() {
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100vh", background: COLORS.bg, fontFamily: "'Inter', sans-serif" }}>
      <div style={{ textAlign: "center" }}>
        <div style={{ fontSize: 32, fontWeight: 800, marginBottom: 16, color: COLORS.navy }}>Procure<span style={{ color: COLORS.accent }}>Bench</span></div>
        <div style={{ color: COLORS.muted }}>Loading benchmark database...</div>
      </div>
    </div>
  );
}

function ErrorScreen({ error }) {
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100vh", background: COLORS.bg, fontFamily: "'Inter', sans-serif" }}>
      <div style={{ textAlign: "center", maxWidth: 500, padding: 40 }}>
        <div style={{ fontSize: 32, fontWeight: 800, marginBottom: 16, color: COLORS.navy }}>Procure<span style={{ color: COLORS.accent }}>Bench</span></div>
        <div style={{ background: "#fef2f2", border: "1px solid #fecaca", borderRadius: 12, padding: 20, color: "#991b1b", fontSize: 14 }}>
          <strong>Database Connection Error</strong><br/>{error}<br/><br/>
          Please check your Supabase configuration in the environment variables.
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════
// LANDING
// ═══════════════════════════════════════
function Landing({ onLogin }) {
  const [email, setEmail] = useState("");
  const ok = email.includes("@") && email.includes(".");
  return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", minHeight: "100vh", background: `linear-gradient(160deg, ${COLORS.navy} 0%, #162240 40%, #1e3a5f 100%)`, color: "white", padding: 40 }}>
      <div style={{ textAlign: "center", marginBottom: 48 }}>
        <div style={{ fontSize: 56, fontWeight: 900, letterSpacing: "-2px", marginBottom: 12 }}>
          Procure<span style={{ color: COLORS.accent }}>Bench</span>
        </div>
        <div style={{ fontSize: 20, fontWeight: 300, opacity: 0.7, maxWidth: 480, lineHeight: 1.7 }}>
          AI-powered spend benchmarking. Compare your procurement performance against industry leaders.
        </div>
      </div>
      <div style={{ background: "rgba(255,255,255,0.06)", backdropFilter: "blur(24px)", borderRadius: 20, padding: "36px 44px", maxWidth: 440, width: "100%", border: "1px solid rgba(255,255,255,0.1)" }}>
        <div style={{ fontSize: 14, fontWeight: 500, marginBottom: 20, opacity: 0.8 }}>Enter your work email to get started</div>
        <input type="email" placeholder="you@company.com" value={email} onChange={e => setEmail(e.target.value)}
          onKeyDown={e => e.key === "Enter" && ok && onLogin(email)}
          style={{ width: "100%", padding: "14px 18px", borderRadius: 10, border: "1px solid rgba(255,255,255,0.15)", background: "rgba(255,255,255,0.08)", color: "white", fontSize: 16, outline: "none", marginBottom: 16, boxSizing: "border-box" }} />
        <button onClick={() => ok && onLogin(email)} disabled={!ok}
          style={{ width: "100%", padding: 14, borderRadius: 10, border: "none", background: ok ? `linear-gradient(135deg, ${COLORS.accent}, ${COLORS.accentDark})` : "rgba(255,255,255,0.1)", color: "white", fontSize: 16, fontWeight: 600, cursor: ok ? "pointer" : "default", opacity: ok ? 1 : 0.4 }}>
          Get Started
        </button>
      </div>
      <div style={{ display: "flex", gap: 48, marginTop: 56, opacity: 0.4, fontSize: 13, fontWeight: 500, letterSpacing: "0.5px" }}>
        <span>612 DATA POINTS</span><span>1,469 ALIASES</span><span>4 INDUSTRIES</span><span>AI ADVISOR</span>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════
// HEADER
// ═══════════════════════════════════════
function Header({ user, view, setView, onLogout }) {
  const navBtn = (v, label) => (
    <button onClick={() => setView(v)} style={{ padding: "8px 20px", borderRadius: 8, border: "none", cursor: "pointer", fontSize: 13, fontWeight: 600,
      background: view === v ? "rgba(255,255,255,0.15)" : "transparent", color: "white", transition: "all 0.2s" }}>{label}</button>
  );
  return (
    <div style={{ background: `linear-gradient(135deg, ${COLORS.navy}, #162240)`, padding: "16px 36px", color: "white", display: "flex", alignItems: "center", justifyContent: "space-between", flexShrink: 0 }}>
      <div>
        <div style={{ fontSize: 20, fontWeight: 800, letterSpacing: "-0.5px" }}>Procure<span style={{ color: COLORS.accent }}>Bench</span></div>
        <div style={{ fontSize: 11, opacity: 0.5, marginTop: 2, fontWeight: 400 }}>{user}</div>
      </div>
      <div style={{ display: "flex", gap: 6, background: "rgba(255,255,255,0.06)", borderRadius: 10, padding: 4 }}>
        {navBtn("quick", "Quick Lookup")}
        {navBtn("upload", "Excel Upload")}
        <button onClick={onLogout} style={{ padding: "8px 16px", borderRadius: 8, border: "none", cursor: "pointer", fontSize: 13, fontWeight: 500, background: "transparent", color: "rgba(255,255,255,0.5)" }}>Logout</button>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════
// SHARED COMPONENTS
// ═══════════════════════════════════════
const cardStyle = { background: COLORS.card, borderRadius: 16, padding: 28, marginBottom: 24, boxShadow: "0 1px 3px rgba(0,0,0,0.04), 0 4px 20px rgba(0,0,0,0.02)", border: `1px solid ${COLORS.border}` };
const titleStyle = { fontSize: 15, fontWeight: 700, marginBottom: 20, color: COLORS.text, display: "flex", alignItems: "center", gap: 10, letterSpacing: "-0.2px" };
const labelStyle = { fontSize: 11, fontWeight: 600, color: COLORS.muted, marginBottom: 8, display: "block", textTransform: "uppercase", letterSpacing: "1px" };
const selectStyle = { width: "100%", padding: "12px 14px", borderRadius: 10, border: `1.5px solid ${COLORS.border}`, fontSize: 14, background: "white", outline: "none", cursor: "pointer", appearance: "none" };
const inputStyle = { width: "100%", padding: "12px 14px", borderRadius: 10, border: `1.5px solid ${COLORS.border}`, fontSize: 14, outline: "none", boxSizing: "border-box" };
const grid2 = { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 };
const grid3 = { display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 20 };

function Btn({ variant = "default", style: sx, children, ...props }) {
  const styles = {
    primary: { background: `linear-gradient(135deg, ${COLORS.blue}, #1d4ed8)`, color: "white" },
    accent: { background: `linear-gradient(135deg, ${COLORS.accent}, ${COLORS.accentDark})`, color: "white" },
    default: { background: "#f1f5f9", color: COLORS.text },
  };
  const s = styles[variant] || styles.default;
  return <button {...props} style={{ padding: "12px 24px", borderRadius: 10, border: "none", cursor: "pointer", fontSize: 14, fontWeight: 600, transition: "all 0.15s", ...s, ...sx }}>{children}</button>;
}

function Badge({ color, children }) {
  return <span style={{ display: "inline-flex", alignItems: "center", padding: "4px 12px", borderRadius: 20, fontSize: 12, fontWeight: 700, color: "white", background: color }}>{children}</span>;
}

function Metric({ value, label, color }) {
  return (
    <div style={{ textAlign: "center", padding: "20px 16px", background: `linear-gradient(135deg, ${color}08, ${color}04)`, borderRadius: 14, border: `1px solid ${color}20` }}>
      <div style={{ fontSize: 28, fontWeight: 800, color: color || COLORS.navy, letterSpacing: "-1px" }}>{value}</div>
      <div style={{ fontSize: 10, color: COLORS.muted, marginTop: 6, textTransform: "uppercase", fontWeight: 700, letterSpacing: "1px" }}>{label}</div>
    </div>
  );
}

// ═══════════════════════════════════════
// QUICK LOOKUP
// ═══════════════════════════════════════
function QuickLookup({ benchmarks, l1Summary, onAgentMsg }) {
  const [industry, setIndustry] = useState("");
  const [l1, setL1] = useState("");
  const [l2, setL2] = useState("");
  const [year, setYear] = useState("");
  const [spend, setSpend] = useState("");
  const [revenue, setRevenue] = useState("");
  const [pct, setPct] = useState("");
  const [mode, setMode] = useState("amount");
  const [result, setResult] = useState(null);

  const industries = useMemo(() => [...new Set(benchmarks.map(b => b.industry))].sort(), [benchmarks]);
  const l1s = useMemo(() => industry ? [...new Set(benchmarks.filter(b => b.industry === industry).map(b => b.l1_category))].sort() : [], [industry, benchmarks]);
  const l2s = useMemo(() => (industry && l1) ? [...new Set(benchmarks.filter(b => b.industry === industry && b.l1_category === l1).map(b => b.l2_category))].sort() : [], [industry, l1, benchmarks]);
  const years = useMemo(() => (industry && l1 && l2) ? [...new Set(benchmarks.filter(b => b.industry === industry && b.l1_category === l1 && b.l2_category === l2).map(b => b.year))].sort() : [], [industry, l1, l2, benchmarks]);

  useEffect(() => { setL1(""); setL2(""); setYear(""); setResult(null); }, [industry]);
  useEffect(() => { setL2(""); setYear(""); setResult(null); }, [l1]);
  useEffect(() => { setYear(""); setResult(null); }, [l2]);

  const compute = () => {
    let userPct = mode === "pct" ? parseFloat(pct) : (() => { const s = parseSpend(spend), r = parseSpend(revenue); return r ? (s / r) * 100 : NaN; })();
    if (isNaN(userPct)) return;
    const bm = benchmarks.find(b => b.industry === industry && b.l1_category === l1 && b.l2_category === l2 && b.year === Number(year));
    if (!bm) return;
    const pos = getPosition(userPct, bm.leader_spend_pct, bm.median_spend_pct, bm.laggard_spend_pct);
    userPct = Math.round(userPct * 100) / 100;
    setResult({ userPct, bm, pos, industry, l1, l2, year });
    onAgentMsg(`Quick Lookup: ${industry} > ${l1} > ${l2} (${year}). User spend: ${userPct}% of revenue. Median: ${bm.median_spend_pct}%, Leader: ${bm.leader_spend_pct}%, Laggard: ${bm.laggard_spend_pct}%. Position: ${pos.label}. Analyze and give 2-3 improvement tips.`);
  };
  const canGo = industry && l1 && l2 && year && (mode === "pct" ? pct : (spend && revenue));

  return (
    <div>
      <div style={cardStyle}>
        <div style={titleStyle}>Quick Benchmark Lookup</div>
        <div style={grid2}>
          <div><label style={labelStyle}>Industry</label><select style={selectStyle} value={industry} onChange={e => setIndustry(e.target.value)}><option value="">Select...</option>{industries.map(i => <option key={i}>{i}</option>)}</select></div>
          <div><label style={labelStyle}>L1 Category</label><select style={selectStyle} value={l1} onChange={e => setL1(e.target.value)} disabled={!industry}><option value="">Select...</option>{l1s.map(c => <option key={c}>{c}</option>)}</select></div>
          <div><label style={labelStyle}>L2 Category</label><select style={selectStyle} value={l2} onChange={e => setL2(e.target.value)} disabled={!l1}><option value="">Select...</option>{l2s.map(c => <option key={c}>{c}</option>)}</select></div>
          <div><label style={labelStyle}>Year</label><select style={selectStyle} value={year} onChange={e => setYear(e.target.value)} disabled={!l2}><option value="">Select...</option>{years.map(y => <option key={y}>{y}</option>)}</select></div>
        </div>
      </div>
      <div style={cardStyle}>
        <div style={titleStyle}>Your Spend Data</div>
        <div style={{ display: "flex", gap: 10, marginBottom: 20 }}>
          <Btn variant={mode === "amount" ? "primary" : "default"} style={{ flex: 1, fontSize: 13, padding: "10px 16px" }} onClick={() => setMode("amount")}>Spend & Revenue</Btn>
          <Btn variant={mode === "pct" ? "primary" : "default"} style={{ flex: 1, fontSize: 13, padding: "10px 16px" }} onClick={() => setMode("pct")}>Direct %</Btn>
        </div>
        {mode === "amount" ? (
          <div style={grid2}>
            <div><label style={labelStyle}>Category Spend (USD)</label><input style={inputStyle} placeholder="e.g. $50M" value={spend} onChange={e => setSpend(e.target.value)} /></div>
            <div><label style={labelStyle}>Total Revenue (USD)</label><input style={inputStyle} placeholder="e.g. $2.4B" value={revenue} onChange={e => setRevenue(e.target.value)} /></div>
          </div>
        ) : (
          <div><label style={labelStyle}>Spend as % of Revenue</label><input style={inputStyle} placeholder="e.g. 2.5" value={pct} onChange={e => setPct(e.target.value)} /></div>
        )}
        <Btn variant="accent" style={{ marginTop: 20, width: "100%", padding: 14 }} disabled={!canGo} onClick={compute}>Compare Against Benchmarks</Btn>
      </div>
      {result && <QuickResults r={result} benchmarks={benchmarks} />}
    </div>
  );
}

function QuickResults({ r, benchmarks }) {
  const { userPct, bm, pos, industry, l1, l2, year } = r;
  const barData = [
    { name: "Leader", value: bm.leader_spend_pct, fill: COLORS.leader },
    { name: "Median", value: bm.median_spend_pct, fill: COLORS.median },
    { name: "You", value: userPct, fill: pos.color },
    { name: "Laggard", value: bm.laggard_spend_pct, fill: COLORS.laggard },
  ];
  const gapL = userPct - bm.leader_spend_pct, gapM = userPct - bm.median_spend_pct;
  const yoy = benchmarks.filter(b => b.industry === industry && b.l1_category === l1 && b.l2_category === l2).map(b => ({ year: b.year, leader: b.leader_spend_pct, median: b.median_spend_pct, laggard: b.laggard_spend_pct }));

  return (
    <div>
      <div style={{ ...cardStyle, background: `linear-gradient(135deg, ${pos.color}08, transparent)`, borderLeft: `4px solid ${pos.color}` }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <div style={{ fontSize: 12, color: COLORS.muted, marginBottom: 6, fontWeight: 500 }}>{industry} / {l1} / {l2} / {year}</div>
            <div style={{ fontSize: 28, fontWeight: 800, color: pos.color, letterSpacing: "-0.5px" }}>{pos.label}</div>
            <div style={{ fontSize: 13, color: COLORS.textSecondary, marginTop: 4 }}>{pos.desc}</div>
          </div>
          <div style={{ textAlign: "right" }}>
            <div style={{ fontSize: 40, fontWeight: 900, color: pos.color, letterSpacing: "-1px" }}>{userPct.toFixed(2)}%</div>
            <div style={{ fontSize: 11, color: COLORS.muted, fontWeight: 500 }}>Your spend % of revenue</div>
          </div>
        </div>
      </div>
      <div style={{ ...grid3, marginBottom: 24 }}>
        <Metric value={`${bm.leader_spend_pct}%`} label="Leader (25th)" color={COLORS.leader} />
        <Metric value={`${bm.median_spend_pct}%`} label="Median (50th)" color={COLORS.median} />
        <Metric value={`${bm.laggard_spend_pct}%`} label="Laggard (75th)" color={COLORS.laggard} />
      </div>
      <div style={cardStyle}>
        <div style={titleStyle}>Benchmark Comparison</div>
        <ResponsiveContainer width="100%" height={280}>
          <BarChart data={barData} margin={{ top: 20, right: 30, left: 20, bottom: 5 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" /><XAxis dataKey="name" tick={{ fontSize: 12, fill: COLORS.textSecondary }} />
            <YAxis tick={{ fontSize: 12, fill: COLORS.textSecondary }} /><Tooltip formatter={v => `${v.toFixed(2)}%`} />
            <Bar dataKey="value" radius={[8, 8, 0, 0]}>{barData.map((d, i) => <Cell key={i} fill={d.fill} />)}</Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
      <div style={cardStyle}>
        <div style={titleStyle}>Year-over-Year Trend</div>
        <ResponsiveContainer width="100%" height={240}>
          <BarChart data={yoy} margin={{ top: 20, right: 30, left: 20, bottom: 5 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" /><XAxis dataKey="year" tick={{ fontSize: 12 }} /><YAxis tick={{ fontSize: 12 }} /><Tooltip formatter={v => `${v.toFixed(2)}%`} /><Legend />
            <Bar dataKey="leader" fill={COLORS.leader} name="Leader" radius={[4, 4, 0, 0]} />
            <Bar dataKey="median" fill={COLORS.median} name="Median" radius={[4, 4, 0, 0]} />
            <Bar dataKey="laggard" fill={COLORS.laggard} name="Laggard" radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>
      <div style={cardStyle}>
        <div style={titleStyle}>Gap Analysis</div>
        <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0, fontSize: 13 }}>
          <thead><tr><th style={{ padding: "12px 14px", textAlign: "left", fontWeight: 700, fontSize: 10, textTransform: "uppercase", letterSpacing: "1px", color: COLORS.muted, borderBottom: `2px solid ${COLORS.border}` }}>Metric</th><th style={{ padding: "12px 14px", textAlign: "left", fontWeight: 700, fontSize: 10, textTransform: "uppercase", letterSpacing: "1px", color: COLORS.muted, borderBottom: `2px solid ${COLORS.border}` }}>Value</th><th style={{ padding: "12px 14px", textAlign: "left", fontWeight: 700, fontSize: 10, textTransform: "uppercase", letterSpacing: "1px", color: COLORS.muted, borderBottom: `2px solid ${COLORS.border}` }}>Insight</th></tr></thead>
          <tbody>
            <tr><td style={{ padding: "12px 14px", borderBottom: "1px solid #f1f5f9" }}>Gap to Leader</td><td style={{ padding: "12px 14px", borderBottom: "1px solid #f1f5f9", fontWeight: 700, color: gapL > 0 ? COLORS.red : COLORS.green }}>{gapL > 0 ? "+" : ""}{gapL.toFixed(2)} pp</td><td style={{ padding: "12px 14px", borderBottom: "1px solid #f1f5f9" }}>{gapL <= 0 ? "At or below leader level" : "Closing this gap places you among leaders"}</td></tr>
            <tr><td style={{ padding: "12px 14px", borderBottom: "1px solid #f1f5f9" }}>Gap to Median</td><td style={{ padding: "12px 14px", borderBottom: "1px solid #f1f5f9", fontWeight: 700, color: gapM > 0 ? COLORS.gold : COLORS.green }}>{gapM > 0 ? "+" : ""}{gapM.toFixed(2)} pp</td><td style={{ padding: "12px 14px", borderBottom: "1px solid #f1f5f9" }}>{gapM <= 0 ? "At or below median" : `${bm.median_spend_pct > 0 ? ((userPct - bm.median_spend_pct) / bm.median_spend_pct * 100).toFixed(0) : 0}% above median`}</td></tr>
            <tr><td style={{ padding: "12px 14px" }}>Sample Size</td><td style={{ padding: "12px 14px", fontWeight: 700 }}>{bm.sample_size}</td><td style={{ padding: "12px 14px" }}>Based on {bm.sample_size} companies</td></tr>
          </tbody>
        </table>
      </div>
      {IMPROVEMENTS[l1] && (
        <div style={cardStyle}>
          <div style={titleStyle}>Improvement Recommendations — {l1}</div>
          {IMPROVEMENTS[l1].map((tip, i) => (
            <div key={i} style={{ display: "flex", gap: 14, padding: "14px 0", borderBottom: i < 3 ? `1px solid ${COLORS.border}` : "none" }}>
              <Badge color={COLORS.accent}>{i + 1}</Badge>
              <div style={{ fontSize: 13, lineHeight: 1.7, color: COLORS.textSecondary }}>{tip}</div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════
// EXCEL UPLOAD
// ═══════════════════════════════════════
function ExcelUpload({ benchmarks, l1Summary, taxonomy, onAgentMsg }) {
  const [step, setStep] = useState("upload");
  const [rawData, setRawData] = useState([]);
  const [columns, setColumns] = useState({ category: null, spend: null, year: null });
  const [headers, setHeaders] = useState([]);
  const [revenue, setRevenue] = useState("");
  const [industry, setIndustry] = useState("");
  const [mappingYear, setMappingYear] = useState("2024");
  const [mappings, setMappings] = useState([]);
  const [results, setResults] = useState(null);
  const [fileName, setFileName] = useState("");
  const [headerRowNum, setHeaderRowNum] = useState(0);
  const [processing, setProcessing] = useState(false);
  const [fileStats, setFileStats] = useState(null);
  const [detectedYears, setDetectedYears] = useState([]);

  const industries = useMemo(() => [...new Set(benchmarks.map(b => b.industry))].sort(), [benchmarks]);
  const steps = ["Upload", "Columns", "Revenue", "Mapping", "Results"];
  const stepIdx = { upload: 0, configure: 1, revenue: 2, mapping: 3, results: 4 };

  const handleFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setFileName(file.name);
    setProcessing(true);
    const reader = new FileReader();
    reader.onload = (evt) => {
      // Use setTimeout to let the UI show "Processing..." before heavy work
      setTimeout(() => {
        try {
          const wb = XLSX.read(evt.target.result, { type: "array" });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
          if (rawRows.length < 2) { setProcessing(false); return; }

          // Smart header detection
          const hdrKw = /category|spend|amount|cost|total|year|vendor|supplier|description|type|group|revenue|industry|region|name|date|fiscal|budget|department|usd|currency|invoice|l1|l2|l3|l4|procurement|buyer/i;
          let bestRow = 0, bestScore = -1;
          for (let i = 0; i < Math.min(rawRows.length, 15); i++) {
            const row = rawRows[i];
            if (!row || row.length === 0) continue;
            const ne = row.filter(c => c !== "" && c !== null && c !== undefined);
            if (ne.length < 2) continue;
            let score = ne.length * 2 + ne.filter(c => typeof c === "string" && isNaN(Number(c))).length * 3 + ne.filter(c => typeof c === "string" && hdrKw.test(c)).length * 10 - ne.filter(c => typeof c === "number").length * 4;
            if (score > bestScore) { bestScore = score; bestRow = i; }
          }

          const headerRow = rawRows[bestRow].map((c, i) => String(c || "").trim() || `Column_${i + 1}`);

          // Find which columns have actual data (check first 200 data rows)
          // Only store columns that have data — saves huge memory on wide sparse files
          const usefulCols = new Set();
          for (let i = bestRow + 1; i < Math.min(rawRows.length, bestRow + 201); i++) {
            const row = rawRows[i];
            if (!row) continue;
            for (let ci = 0; ci < headerRow.length; ci++) {
              if (row[ci] !== undefined && row[ci] !== "" && row[ci] !== null) usefulCols.add(ci);
            }
          }

          // Build data rows — ALL rows, but only useful columns
          const dataRows = [];
          const totalOrigRows = rawRows.length - bestRow - 1;
          for (let i = bestRow + 1; i < rawRows.length; i++) {
            const row = rawRows[i];
            if (!row || row.every(c => !c && c !== 0)) continue;
            const obj = {};
            for (const ci of usefulCols) {
              obj[headerRow[ci]] = row[ci] !== undefined ? row[ci] : "";
            }
            dataRows.push(obj);
          }
          if (!dataRows.length) { setProcessing(false); return; }

          const hdrs = [...new Set(Array.from(usefulCols).map(ci => headerRow[ci]))].filter(h => h);

          // Smart column detection: check DATA DENSITY, not just name
          // Find ALL candidate columns for each role, then pick the one with most non-empty values
          const sampleSize = Math.min(dataRows.length, 500);
          const sample = dataRows.slice(0, sampleSize);

          const countNonEmpty = (col, numeric = false) => {
            let count = 0;
            for (const row of sample) {
              const v = row[col];
              if (v === "" || v === null || v === undefined) continue;
              if (numeric && typeof v !== "number" && isNaN(parseFloat(String(v).replace(/[$,]/g, "")))) continue;
              count++;
            }
            return count;
          };

          // Category: prefer columns with "category", "l1", "l2", "group", "type" in name, with most non-empty string values
          const catCandidates = hdrs.filter(h => /category|l1|l2|l3|group|type|description|classification|segment|class|procurement/i.test(h) && !/date|id|unique|price|amount|spend|cost|invoice/i.test(h));
          let catCol = null;
          let catBest = 0;
          for (const c of catCandidates) {
            const n = countNonEmpty(c);
            if (n > catBest) { catBest = n; catCol = c; }
          }
          if (!catCol) catCol = hdrs[0];

          // Spend: prefer columns with "amount", "spend", "cost", "invoice amount", "price" that have NUMERIC data
          const spendCandidates = hdrs.filter(h => /amount|spend|cost|total|price|value|invoice.*amount/i.test(h) && !/percent|%|ratio|vs|unit.*price|price.*unit|category|date|#|number|count/i.test(h));
          let spendCol = null;
          let spendBest = 0;
          for (const c of spendCandidates) {
            const n = countNonEmpty(c, true);
            if (n > spendBest) { spendBest = n; spendCol = c; }
          }
          // Fallback: any column with $ or USD in name
          if (!spendCol) {
            const fallback = hdrs.find(h => /usd|\$/i.test(h));
            if (fallback) spendCol = fallback;
          }

          // Year: prefer explicit year column, fallback to date columns
          let yearCol = hdrs.find(h => /^year$/i.test(h)) || hdrs.find(h => /year|fy|fiscal/i.test(h) && !/date|day/i.test(h));
          
          // If no year column, try date columns
          if (!yearCol) {
            const dateCandidates = hdrs.filter(h => /date|period|month/i.test(h));
            for (const dc of dateCandidates) {
              const hasYears = sample.filter(r => extractYear(r[dc]) !== null).length;
              if (hasYears > sampleSize * 0.3) { yearCol = dc; break; }
            }
          }
          
          // Detect unique years in the data
          const detectedYearSet = new Set();
          for (const row of sample) {
            if (yearCol) {
              const yr = extractYear(row[yearCol]);
              if (yr) detectedYearSet.add(yr);
            }
          }
          const detectedYears = [...detectedYearSet].sort();

          setFileStats({ totalRows: totalOrigRows, processedRows: dataRows.length, droppedCols: headerRow.length - usefulCols.size });
          setDetectedYears(detectedYears);
          if (detectedYears.length > 0) setMappingYear(String(Math.max(...detectedYears)));
          setRawData(dataRows);
          setHeaders(hdrs);
          setHeaderRowNum(bestRow);
          setColumns({ category: catCol, spend: spendCol, year: yearCol });
          setStep("configure");
          setProcessing(false);
          onAgentMsg(`Uploaded "${file.name}": ${dataRows.length.toLocaleString()} rows processed, header at row ${bestRow + 1}, ${hdrs.length} columns (${headerRow.length - usefulCols.size} empty columns dropped). Auto-detected Category="${catCol}" (${catBest} values), Spend="${spendCol || "?"}" (${spendBest} values), Year="${yearCol || "none"}". Help verify.`);
        } catch (err) {
          setProcessing(false);
          onAgentMsg(`Error processing file: ${err.message}. The file may be too large or in an unsupported format.`);
        }
      }, 50);
    };
    reader.readAsArrayBuffer(file);
  };

  const proceedToRevenue = () => { if (columns.category && columns.spend) { setStep("revenue"); } };

  const proceedToMapping = () => {
    const rev = parseSpend(revenue);
    if (!rev || !industry) return;
    const catSpend = {};
    for (const row of rawData) {
      if (columns.year) {
        const rowYear = extractYear(row[columns.year]);
        if (rowYear && rowYear !== Number(mappingYear)) continue;
      }
      const cat = String(row[columns.category] || "").trim();
      const s = parseSpend(row[columns.spend]);
      if (cat && s) catSpend[cat] = (catSpend[cat] || 0) + s;
    }
    const maps = Object.entries(catSpend).map(([raw, totalSpend]) => {
      const match = matchCategory(raw, industry, taxonomy);
      return { raw, totalSpend, spendPct: (totalSpend / rev) * 100, ...match, selectedL1: match.l1, selectedL2: match.l2, confirmed: match.confidence >= 80 };
    }).sort((a, b) => b.totalSpend - a.totalSpend);
    setMappings(maps);
    setStep("mapping");
    const mapped = maps.filter(m => m.confidence >= 80).length;
    onAgentMsg(`Rationalized for ${mappingYear}. ${maps.length} categories, ${mapped} mapped (>=80%), ${maps.length - mapped} need review.`);
  };

  const confirmAll = () => setMappings(prev => prev.map(m => ({ ...m, confirmed: true })));
  const updateMapping = (idx, l1, l2) => setMappings(prev => { const n = [...prev]; n[idx] = { ...n[idx], selectedL1: l1, selectedL2: l2, confirmed: true }; return n; });

  const generateResults = () => {
    const rev = parseSpend(revenue);
    const l1Agg = {};
    for (const m of mappings) { if (m.selectedL1) l1Agg[m.selectedL1] = (l1Agg[m.selectedL1] || 0) + m.totalSpend; }
    const comps = Object.entries(l1Agg).map(([l1cat, sp]) => {
      const p = (sp / rev) * 100;
      const bm = l1Summary.find(b => b.industry === industry && b.l1_category === l1cat && b.year === Number(mappingYear));
      const pos = bm ? getPosition(p, bm.leader_spend_pct, bm.median_spend_pct, bm.laggard_spend_pct) : { label: "N/A", color: COLORS.muted, desc: "No benchmark" };
      return { l1: l1cat, spend: sp, pct: Math.round(p * 100) / 100, bm, pos };
    }).sort((a, b) => b.spend - a.spend);
    setResults(comps);
    setStep("results");
    onAgentMsg(`Full analysis: ${industry} (${mappingYear}). ${comps.map(c => `${c.l1}: ${c.pct}% (${c.pos.label})`).join("; ")}. Revenue: $${(rev/1e6).toFixed(0)}M. Provide comprehensive summary and top 3 priorities.`);
  };

  // Progress Stepper Component
  const Stepper = () => (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 0, marginBottom: 28, padding: "0 20px" }}>
      {steps.map((s, i) => (
        <div key={s} style={{ display: "flex", alignItems: "center" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <div style={{ width: 28, height: 28, borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, fontWeight: 700,
              background: i <= stepIdx[step] ? (i === stepIdx[step] ? COLORS.accent : COLORS.green) : "#e2e8f0",
              color: i <= stepIdx[step] ? "white" : COLORS.muted, transition: "all 0.3s" }}>
              {i < stepIdx[step] ? "✓" : i + 1}
            </div>
            <span style={{ fontSize: 12, fontWeight: i === stepIdx[step] ? 700 : 500, color: i <= stepIdx[step] ? COLORS.text : COLORS.muted }}>{s}</span>
          </div>
          {i < steps.length - 1 && <div style={{ width: 40, height: 2, background: i < stepIdx[step] ? COLORS.green : "#e2e8f0", margin: "0 8px", transition: "all 0.3s" }} />}
        </div>
      ))}
    </div>
  );

  if (processing) return (
    <div>
      <Stepper />
      <div style={cardStyle}>
        <div style={{ textAlign: "center", padding: 48 }}>
          <div style={{ fontSize: 36, marginBottom: 16, animation: "spin 1s linear infinite" }}>⟳</div>
          <div style={{ fontSize: 16, fontWeight: 700, color: COLORS.text, marginBottom: 8 }}>Processing {fileName}...</div>
          <div style={{ fontSize: 13, color: COLORS.muted }}>Detecting headers, analyzing columns, and preparing data. Large files may take a moment.</div>
          <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
        </div>
      </div>
    </div>
  );

  if (step === "upload") return (
    <div>
      <Stepper />
      <div style={cardStyle}>
        <div style={titleStyle}>Upload Spend Data</div>
        <div style={{ border: `2px dashed ${COLORS.border}`, borderRadius: 16, padding: 48, textAlign: "center", background: "#fafbfc" }}>
          <div style={{ fontSize: 13, color: COLORS.muted, marginBottom: 20 }}>Supports .xlsx, .xls, .csv — files up to 50,000 rows processed automatically</div>
          <label style={{ padding: "12px 28px", borderRadius: 10, border: "none", cursor: "pointer", fontSize: 14, fontWeight: 600, background: `linear-gradient(135deg, ${COLORS.blue}, #1d4ed8)`, color: "white", display: "inline-block" }}>
            Browse Files<input type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{ display: "none" }} />
          </label>
        </div>
      </div>
    </div>
  );

  if (step === "configure") return (
    <div>
      <Stepper />
      <div style={cardStyle}>
      <div style={titleStyle}>Confirm Columns — {fileName}</div>
      <div style={{ fontSize: 13, color: COLORS.muted, marginBottom: 16 }}>
        {rawData.length.toLocaleString()} rows processed. Header at row {headerRowNum + 1}. {headers.length} columns.
        {fileStats?.droppedCols > 0 && <span style={{ color: COLORS.muted }}> ({fileStats.droppedCols} empty columns dropped)</span>}
      </div>
      <div style={grid3}>
        <div><label style={labelStyle}>Category Column *</label><select style={selectStyle} value={columns.category || ""} onChange={e => setColumns(p => ({ ...p, category: e.target.value }))}><option value="">Select...</option>{headers.map(h => <option key={h}>{h}</option>)}</select></div>
        <div><label style={labelStyle}>Spend Column *</label><select style={selectStyle} value={columns.spend || ""} onChange={e => setColumns(p => ({ ...p, spend: e.target.value }))}><option value="">Select...</option>{headers.map(h => <option key={h}>{h}</option>)}</select></div>
        <div><label style={labelStyle}>Year Column</label><select style={selectStyle} value={columns.year || ""} onChange={e => setColumns(p => ({ ...p, year: e.target.value }))}><option value="">None</option>{headers.map(h => <option key={h}>{h}</option>)}</select></div>
      </div>
      <div style={{ marginTop: 20, overflowX: "auto", borderRadius: 10, border: `1px solid ${COLORS.border}` }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
          <thead><tr>{headers.map(h => <th key={h} style={{ padding: "10px 12px", textAlign: "left", fontWeight: 600, fontSize: 10, textTransform: "uppercase", color: COLORS.muted, borderBottom: `2px solid ${COLORS.border}`, background: [columns.category, columns.spend, columns.year].includes(h) ? "#eff6ff" : "#fafbfc", whiteSpace: "nowrap" }}>{h}</th>)}</tr></thead>
          <tbody>{rawData.slice(0, 10).map((row, i) => <tr key={i}>{headers.map(h => <td key={h} style={{ padding: "8px 12px", borderBottom: "1px solid #f1f5f9", whiteSpace: "nowrap", maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis" }}>{String(row[h]).slice(0, 40)}</td>)}</tr>)}</tbody>
        </table>
      </div>
      <div style={{ display: "flex", gap: 12, marginTop: 20 }}>
        <Btn onClick={() => setStep("upload")}>Back</Btn>
        <Btn variant="accent" style={{ flex: 1 }} onClick={proceedToRevenue} disabled={!columns.category || !columns.spend}>Continue</Btn>
      </div>
    </div>
    </div>
  );

  if (step === "revenue") return (
    <div>
      <Stepper />
      <div style={cardStyle}>
      <div style={titleStyle}>Company Information</div>
      <div style={{ fontSize: 13, color: COLORS.muted, marginBottom: 20 }}>Revenue is needed to calculate spend as % of revenue for benchmarking.</div>
      <div style={grid3}>
        <div><label style={labelStyle}>Total Revenue (USD) *</label><input style={inputStyle} placeholder="e.g. $2.4B" value={revenue} onChange={e => setRevenue(e.target.value)} /></div>
        <div><label style={labelStyle}>Industry *</label><select style={selectStyle} value={industry} onChange={e => setIndustry(e.target.value)}><option value="">Select...</option>{industries.map(i => <option key={i}>{i}</option>)}</select></div>
        <div><label style={labelStyle}>Benchmark Year</label><select style={selectStyle} value={mappingYear} onChange={e => setMappingYear(e.target.value)}>{
          (detectedYears.length > 0 ? detectedYears : [2023, 2024, 2025]).map(y => <option key={y} value={String(y)}>{y}{detectedYears.includes(y) ? " (detected)" : ""}</option>)
        }</select></div>
      </div>
      <div style={{ display: "flex", gap: 12, marginTop: 20 }}>
        <Btn onClick={() => setStep("configure")}>Back</Btn>
        <Btn variant="accent" style={{ flex: 1 }} onClick={proceedToMapping} disabled={!revenue || !industry}>Rationalize & Map</Btn>
      </div>
    </div>
    </div>
  );

  if (step === "mapping") {
    const high = mappings.filter(m => m.confidence >= 80);
    const low = mappings.filter(m => m.confidence < 80);
    const allOk = mappings.every(m => m.confirmed);
    const allL1s = [...new Set(benchmarks.map(b => b.l1_category))].sort();

    return (
      <div>
        <Stepper />
        <div style={cardStyle}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
            <div style={titleStyle}>Category Mapping</div>
            <div style={{ display: "flex", gap: 8 }}><Badge color={COLORS.green}>{high.length} Mapped</Badge><Badge color={low.length ? COLORS.gold : COLORS.green}>{low.length} Review</Badge></div>
          </div>
          {high.length > 0 && (<div style={{ marginBottom: 24 }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: COLORS.green, marginBottom: 10 }}>High Confidence (80%+)</div>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead><tr><th style={{ padding: "10px 14px", textAlign: "left", fontWeight: 600, fontSize: 10, textTransform: "uppercase", color: COLORS.muted, borderBottom: `2px solid ${COLORS.border}` }}>Your Category</th><th style={{ padding: "10px 14px", textAlign: "left", fontWeight: 600, fontSize: 10, textTransform: "uppercase", color: COLORS.muted, borderBottom: `2px solid ${COLORS.border}` }}>Mapped L1</th><th style={{ padding: "10px 14px", textAlign: "left", fontWeight: 600, fontSize: 10, textTransform: "uppercase", color: COLORS.muted, borderBottom: `2px solid ${COLORS.border}` }}>L2</th><th style={{ padding: "10px 14px", textAlign: "left", fontWeight: 600, fontSize: 10, textTransform: "uppercase", color: COLORS.muted, borderBottom: `2px solid ${COLORS.border}` }}>Confidence</th><th style={{ padding: "10px 14px", textAlign: "left", fontWeight: 600, fontSize: 10, textTransform: "uppercase", color: COLORS.muted, borderBottom: `2px solid ${COLORS.border}` }}>Spend</th></tr></thead>
              <tbody>{high.map((m, i) => (<tr key={i}><td style={{ padding: "10px 14px", borderBottom: "1px solid #f1f5f9", fontWeight: 600 }}>{m.raw}</td><td style={{ padding: "10px 14px", borderBottom: "1px solid #f1f5f9" }}>{m.selectedL1}</td><td style={{ padding: "10px 14px", borderBottom: "1px solid #f1f5f9" }}>{m.selectedL2 || "—"}</td><td style={{ padding: "10px 14px", borderBottom: "1px solid #f1f5f9" }}><Badge color={COLORS.green}>{m.confidence}%</Badge></td><td style={{ padding: "10px 14px", borderBottom: "1px solid #f1f5f9" }}>${(m.totalSpend / 1e6).toFixed(1)}M</td></tr>))}</tbody>
            </table>
          </div>)}
          {low.length > 0 && (<div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: COLORS.gold }}>Needs Review</div>
              <Btn variant="accent" style={{ fontSize: 12, padding: "8px 16px" }} onClick={confirmAll}>Confirm All</Btn>
            </div>
            {low.map((m, idx) => {
              const gi = mappings.indexOf(m);
              return (
                <div key={idx} style={{ background: m.confirmed ? "#f0fdf4" : "#fffbeb", border: `1px solid ${m.confirmed ? "#bbf7d0" : "#fde68a"}`, borderRadius: 12, padding: 18, marginBottom: 12 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
                    <div><span style={{ fontWeight: 700 }}>{m.raw}</span> <Badge color={m.confidence >= 55 ? COLORS.gold : COLORS.red}>{m.confidence}%</Badge> <span style={{ fontSize: 11, color: COLORS.muted, marginLeft: 8 }}>{m.reason}</span></div>
                    <div style={{ fontSize: 13, fontWeight: 600 }}>${(m.totalSpend / 1e6).toFixed(1)}M ({m.spendPct.toFixed(2)}%)</div>
                  </div>
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                    {m.matches.slice(0, 5).map((opt, oi) => (
                      <button key={oi} onClick={() => updateMapping(gi, opt.l1, opt.l2)}
                        style={{ padding: "6px 14px", borderRadius: 8, fontSize: 12, cursor: "pointer", border: `1px solid ${m.selectedL1 === opt.l1 && m.confirmed ? COLORS.green : COLORS.border}`, background: m.selectedL1 === opt.l1 && m.confirmed ? "#dcfce7" : "white", fontWeight: m.selectedL1 === opt.l1 && m.confirmed ? 700 : 400 }}>
                        {opt.l1}{opt.l2 ? ` / ${opt.l2}` : ""} ({opt.confidence}%)
                      </button>
                    ))}
                    <select style={{ ...selectStyle, width: "auto", fontSize: 12, padding: "6px 10px" }} onChange={e => { if (e.target.value) updateMapping(gi, e.target.value, null); }} value=""><option value="">Other...</option>{allL1s.map(l => <option key={l}>{l}</option>)}</select>
                  </div>
                </div>
              );
            })}
          </div>)}
        </div>
        <div style={{ display: "flex", gap: 12 }}>
          <Btn onClick={() => setStep("revenue")}>Back</Btn>
          <Btn variant="primary" style={{ flex: 1, padding: 14 }} onClick={generateResults} disabled={!allOk}>
            {allOk ? "Generate Benchmark Report" : `Confirm ${mappings.filter(m => !m.confirmed).length} remaining`}
          </Btn>
        </div>
      </div>
    );
  }

  if (step === "results" && results) {
    const totalSpend = results.reduce((s, r) => s + r.spend, 0);
    const totalPct = results.reduce((s, r) => s + r.pct, 0);
    const barData = results.map(r => ({ name: r.l1.length > 18 ? r.l1.slice(0, 16) + "…" : r.l1, you: r.pct, median: r.bm?.median_spend_pct || 0, leader: r.bm?.leader_spend_pct || 0, laggard: r.bm?.laggard_spend_pct || 0 }));
    const radarData = results.filter(r => r.bm).map(r => ({ cat: r.l1.split(" ")[0], you: r.pct, median: r.bm.median_spend_pct }));
    const posCounts = {};
    results.forEach(r => { posCounts[r.pos.label] = (posCounts[r.pos.label] || 0) + 1; });
    const pieData = Object.entries(posCounts).map(([name, value]) => ({ name, value, fill: name === "Leader" ? COLORS.leader : name === "Above Median" ? COLORS.accent : name === "Below Median" ? COLORS.gold : name === "Laggard" ? COLORS.red : COLORS.muted }));

    return (
      <div>
        <Stepper />
        <div style={{ ...grid3, marginBottom: 24 }}>
          <Metric value={`$${(totalSpend / 1e6).toFixed(0)}M`} label="Total Spend" color={COLORS.navy} />
          <Metric value={`${totalPct.toFixed(1)}%`} label="% of Revenue" color={COLORS.blue} />
          <Metric value={results.length} label="Categories" color={COLORS.accent} />
        </div>
        <div style={cardStyle}>
          <div style={titleStyle}>Performance Distribution</div>
          <div style={{ display: "flex", alignItems: "center", gap: 32 }}>
            <ResponsiveContainer width="50%" height={220}>
              <PieChart><Pie data={pieData} cx="50%" cy="50%" innerRadius={55} outerRadius={90} paddingAngle={3} dataKey="value">{pieData.map((d, i) => <Cell key={i} fill={d.fill} />)}</Pie><Tooltip /></PieChart>
            </ResponsiveContainer>
            <div>{pieData.map((d, i) => (<div key={i} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}><div style={{ width: 14, height: 14, borderRadius: 4, background: d.fill }} /><span style={{ fontSize: 13, fontWeight: 600 }}>{d.name}: {d.value}</span></div>))}</div>
          </div>
        </div>
        <div style={cardStyle}>
          <div style={titleStyle}>Full Category Comparison</div>
          <ResponsiveContainer width="100%" height={Math.max(350, results.length * 40)}>
            <BarChart data={barData} layout="vertical" margin={{ top: 5, right: 30, left: 120, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" /><XAxis type="number" tick={{ fontSize: 11 }} /><YAxis type="category" dataKey="name" tick={{ fontSize: 11 }} width={110} /><Tooltip formatter={v => `${v.toFixed(2)}%`} /><Legend />
              <Bar dataKey="leader" fill={COLORS.leader} name="Leader" barSize={8} /><Bar dataKey="median" fill={COLORS.median} name="Median" barSize={8} /><Bar dataKey="you" fill={COLORS.gold} name="Your Spend" barSize={8} /><Bar dataKey="laggard" fill={COLORS.laggard} name="Laggard" barSize={8} />
            </BarChart>
          </ResponsiveContainer>
        </div>
        {radarData.length >= 3 && (<div style={cardStyle}>
          <div style={titleStyle}>Spend Profile</div>
          <ResponsiveContainer width="100%" height={350}>
            <RadarChart data={radarData}><PolarGrid stroke="#e2e8f0" /><PolarAngleAxis dataKey="cat" tick={{ fontSize: 10 }} /><PolarRadiusAxis tick={{ fontSize: 10 }} />
              <Radar name="You" dataKey="you" stroke={COLORS.gold} fill={COLORS.gold} fillOpacity={0.3} /><Radar name="Median" dataKey="median" stroke={COLORS.median} fill={COLORS.median} fillOpacity={0.15} /><Legend /><Tooltip formatter={v => `${v.toFixed(2)}%`} /></RadarChart>
          </ResponsiveContainer>
        </div>)}
        <div style={cardStyle}>
          <div style={titleStyle}>Detailed Results</div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead><tr>{["Category","Spend","Your %","Leader","Median","Laggard","Position"].map(h => <th key={h} style={{ padding: "12px 14px", textAlign: "left", fontWeight: 600, fontSize: 10, textTransform: "uppercase", letterSpacing: "1px", color: COLORS.muted, borderBottom: `2px solid ${COLORS.border}` }}>{h}</th>)}</tr></thead>
              <tbody>{results.map((r, i) => (<tr key={i}><td style={{ padding: "12px 14px", borderBottom: "1px solid #f1f5f9", fontWeight: 600 }}>{r.l1}</td><td style={{ padding: "12px 14px", borderBottom: "1px solid #f1f5f9" }}>${(r.spend / 1e6).toFixed(1)}M</td><td style={{ padding: "12px 14px", borderBottom: "1px solid #f1f5f9", fontWeight: 700 }}>{r.pct.toFixed(2)}%</td><td style={{ padding: "12px 14px", borderBottom: "1px solid #f1f5f9", color: COLORS.leader }}>{r.bm?.leader_spend_pct ?? "—"}%</td><td style={{ padding: "12px 14px", borderBottom: "1px solid #f1f5f9", color: COLORS.median }}>{r.bm?.median_spend_pct ?? "—"}%</td><td style={{ padding: "12px 14px", borderBottom: "1px solid #f1f5f9", color: COLORS.laggard }}>{r.bm?.laggard_spend_pct ?? "—"}%</td><td style={{ padding: "12px 14px", borderBottom: "1px solid #f1f5f9" }}><Badge color={r.pos.color}>{r.pos.label}</Badge></td></tr>))}</tbody>
            </table>
          </div>
        </div>
        <div style={cardStyle}>
          <div style={titleStyle}>Priority Improvements</div>
          {results.filter(r => r.pos.label === "Laggard" || r.pos.label === "Below Median").slice(0, 3).map((r, i) => (
            <div key={i} style={{ marginBottom: 20, padding: 18, background: `${r.pos.color}08`, borderRadius: 12, borderLeft: `3px solid ${r.pos.color}` }}>
              <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 10 }}>{r.l1} — <span style={{ color: r.pos.color }}>{r.pos.label}</span> ({r.pct.toFixed(2)}% vs {r.bm?.median_spend_pct}% median)</div>
              {(IMPROVEMENTS[r.l1] || []).slice(0, 2).map((tip, ti) => (<div key={ti} style={{ fontSize: 13, color: COLORS.textSecondary, marginBottom: 6, paddingLeft: 14, borderLeft: `2px solid ${COLORS.border}` }}>{tip}</div>))}
            </div>
          ))}
        </div>
      </div>
    );
  }
  return null;
}

// ═══════════════════════════════════════
// AGENT PANEL
// ═══════════════════════════════════════
function AgentPanel({ contextMsgs }) {
  const [messages, setMessages] = useState([
    { role: "assistant", content: "Welcome to ProcureBench!\n\nI'm your procurement benchmarking advisor. Choose an option:\n\n**Quick Lookup** — Compare a single category\n**Excel Upload** — Full portfolio analysis\n\nI'll guide you through every step." }
  ]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const chatRef = useRef(null);
  const processed = useRef(new Set());

  useEffect(() => { if (chatRef.current) chatRef.current.scrollTop = chatRef.current.scrollHeight; }, [messages]);

  useEffect(() => {
    if (!contextMsgs.length) return;
    const latest = contextMsgs[contextMsgs.length - 1];
    if (processed.current.has(latest)) return;
    processed.current.add(latest);
    callAgent(latest, true);
  }, [contextMsgs]);

  const callAgent = async (msg, isSystem = false) => {
    setLoading(true);
    if (!isSystem) setMessages(prev => [...prev, { role: "user", content: msg }]);
    try {
      const hist = messages.filter(m => m.role !== "system").map(m => ({ role: m.role, content: m.content }));
      hist.push({ role: "user", content: isSystem ? `[SYSTEM] ${msg}` : msg });
      const res = await fetch("/api/chat", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ model: "claude-sonnet-4-20250514", max_tokens: 2000, system: AGENT_SYSTEM, messages: hist.slice(-12) }),
      });
      const data = await res.json();
      const reply = data.content?.map(c => c.text || "").join("") || "How can I help?";
      setMessages(prev => [...prev, { role: "assistant", content: reply }]);
    } catch {
      setMessages(prev => [...prev, { role: "assistant", content: "I'm having trouble connecting. What would you like to know?" }]);
    }
    setLoading(false);
  };

  const send = () => { if (!input.trim() || loading) return; const m = input.trim(); setInput(""); callAgent(m); };

  return (
    <div style={{ width: 380, borderLeft: `1px solid ${COLORS.border}`, display: "flex", flexDirection: "column", background: "#fafbfc" }}>
      <div style={{ padding: "18px 22px", borderBottom: `1px solid ${COLORS.border}`, display: "flex", alignItems: "center", gap: 12, background: "white" }}>
        <div style={{ width: 36, height: 36, borderRadius: 10, background: `linear-gradient(135deg, ${COLORS.accent}, ${COLORS.accentDark})`, display: "flex", alignItems: "center", justifyContent: "center", color: "white", fontSize: 14, fontWeight: 800 }}>AI</div>
        <div><div style={{ fontSize: 14, fontWeight: 700, color: COLORS.text }}>ProcureBench Advisor</div><div style={{ fontSize: 11, color: COLORS.accent, fontWeight: 500 }}>Online</div></div>
      </div>
      <div ref={chatRef} style={{ flex: 1, overflow: "auto", padding: 18, display: "flex", flexDirection: "column", gap: 14 }}>
        {messages.map((m, i) => (
          <div key={i} style={{
            maxWidth: "88%", padding: "12px 16px", fontSize: 13, lineHeight: 1.7,
            borderRadius: m.role === "user" ? "16px 16px 4px 16px" : "16px 16px 16px 4px",
            background: m.role === "user" ? `linear-gradient(135deg, ${COLORS.blue}, #1d4ed8)` : "white",
            color: m.role === "user" ? "white" : COLORS.text,
            alignSelf: m.role === "user" ? "flex-end" : "flex-start",
            boxShadow: m.role === "user" ? "none" : "0 1px 4px rgba(0,0,0,0.04)",
            border: m.role === "user" ? "none" : `1px solid ${COLORS.border}`,
          }} dangerouslySetInnerHTML={{ __html: m.content.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>").replace(/\n/g, "<br/>") }} />
        ))}
        {loading && <div style={{ alignSelf: "flex-start", padding: "12px 16px", borderRadius: "16px 16px 16px 4px", background: "white", border: `1px solid ${COLORS.border}`, color: COLORS.muted, fontSize: 13, fontStyle: "italic" }}>Analyzing...</div>}
      </div>
      <div style={{ display: "flex", gap: 8, padding: "14px 18px", borderTop: `1px solid ${COLORS.border}`, background: "white" }}>
        <input style={{ ...inputStyle, flex: 1 }} placeholder="Ask about benchmarks..." value={input} onChange={e => setInput(e.target.value)} onKeyDown={e => e.key === "Enter" && send()} />
        <Btn variant="primary" style={{ padding: "12px 20px" }} onClick={send} disabled={loading || !input.trim()}>Send</Btn>
      </div>
    </div>
  );
}
