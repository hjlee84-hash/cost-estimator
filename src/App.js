import { useState, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── 상수 ───────────────────────────────────────────────
const WAGE_TABLE = {
  2023: 20300, 2024: 21400, 2025: 22600, 2026: 26000,
  2027: 27400, 2028: 28800, 2029: 30200, 2030: 31700, 2031: 33100,
};

const PRESS_TABLE = {
  150: { acquisition: 350000000, power: 22, load: 0.7, area: 25 },
  300: { acquisition: 700000000, power: 45, load: 0.7, area: 40 },
};

const PROCESS_RULES = {
  "MTG PLATE":    { flux: false, inspection: true },
  "LOWER PLATE":  { flux: true,  inspection: false },
  "TUBE PLATE A": { flux: true,  inspection: false },
  "TUBE PLATE B": { flux: true,  inspection: false },
  "UPPER PLATE":  { flux: true,  inspection: false },
  "END PLATE":    { flux: false, inspection: true },
  "TURBULATOR":   { flux: false, inspection: false },
};

const MOLD_TABLE = {
  150: { base: 8000000,  perPart: 500000 },
  300: { base: 15000000, perPart: 800000 },
};

const YEARS = Object.keys(WAGE_TABLE).map(Number);
const DEFAULT_ANNUAL_QTY = { 2025: "", 2026: "", 2027: "", 2028: "", 2029: "", 2030: "" };

function calcExpense(ton) {
  const t = PRESS_TABLE[ton];
  const depreciationPerHour = (t.acquisition / 10) / 2000;
  const powerCostPerHour = t.power * t.load * 120;
  const maintenancePerHour = (t.acquisition * 0.03) / 2000;
  const spaceRentPerHour = (t.area * 50000 * 12) / 2000;
  return depreciationPerHour + powerCostPerHour + maintenancePerHour + spaceRentPerHour;
}

function calcCycleTime(w, l, t) {
  const area = w * l;
  let base = 6;
  if (area > 10000) base = 10;
  if (area > 20000) base = 15;
  if (t >= 3) base += 3;
  return base;
}

function calcPartCost(part, sopYear, totalAnnualQty) {
  const { name, w, l, t, qty } = part;
  if (!w || !l || !t) return null;

  const ton = parseFloat(t) >= 3 ? 300 : 150;
  const wagePerHour = WAGE_TABLE[sopYear] || 26000;
  const cycleTimeSec = calcCycleTime(w, l, t);
  const cycleTimeHour = cycleTimeSec / 3600;
  const pcsPerHour = Math.floor(3600 / cycleTimeSec);
  const laborPerPc = wagePerHour * cycleTimeHour;
  const expensePerHour = calcExpense(ton);
  const expensePerPc = expensePerHour * cycleTimeHour;

  const rule = PROCESS_RULES[name] || { flux: false, inspection: false };
  const washCost = 500;
  const fluxCost = rule.flux ? 800 : 0;
  const inspectionCost = rule.inspection ? 1200 : 0;
  const processCost = washCost + fluxCost + inspectionCost;

  const subTotal = laborPerPc + expensePerPc + processCost;
  const mgmtCost = (laborPerPc + expensePerPc) * 0.25;
  const profit = (laborPerPc + expensePerPc + mgmtCost) * 0.15;
  const unitPrice = subTotal + mgmtCost + profit;

  const totalMoldCost = MOLD_TABLE[ton].base + MOLD_TABLE[ton].perPart;
  const moldAmortizationPerPc = totalMoldCost / Math.max(totalAnnualQty * 5, 10000);

  return {
    ton, wagePerHour, cycleTimeSec, pcsPerHour,
    laborPerPc: Math.round(laborPerPc),
    expensePerPc: Math.round(expensePerPc),
    processCost: Math.round(processCost),
    washCost, fluxCost: Math.round(fluxCost), inspectionCost,
    subTotal: Math.round(subTotal),
    mgmtCost: Math.round(mgmtCost),
    profit: Math.round(profit),
    unitPrice: Math.round(unitPrice),
    totalMoldCost,
    moldAmortizationPerPc: Math.round(moldAmortizationPerPc),
    rule,
  };
}

async function parseBomWithAI(base64Image, mimeType) {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1000,
      messages: [{
        role: "user",
        content: [
          { type: "image", source: { type: "base64", media_type: mimeType, data: base64Image } },
          { type: "text", text: `이 BOM 이미지에서 부품 정보를 추출해줘. W, L, T 값이 있는 부품만 추출하고, 없으면 null로. 반드시 아래 JSON 배열만 응답해. 다른 텍스트 없이.\n[{"no":1,"name":"MTG PLATE","material":"AA3003 H16","w":159,"l":109,"t":4,"qty":1,"weight":"143.0g","remark":""},...]` }
        ]
      }]
    })
  });
  const data = await res.json();
  const text = data.content?.[0]?.text || "[]";
  const clean = text.replace(/```json|```/g, "").trim();
  return JSON.parse(clean);
}

function downloadExcel(project, totalQty) {
  const wb = XLSX.utils.book_new();
  const sopYear = project.sopYear;

  const summaryData = [
    ["프로젝트명", project.name],
    ["SOP 연도", sopYear],
    ["임률 (원/시간)", WAGE_TABLE[sopYear] || ""],
    ["총 연간 수량", totalQty],
    [],
    ["NO", "부품명", "재질", "W", "L", "T", "수량", "톤수", "노무비", "경비", "공정비", "일반관리비", "이윤", "단가", "금형비(총)"],
  ];
  project.parts.forEach(p => {
    const c = calcPartCost(p, sopYear, totalQty);
    if (!c) return;
    summaryData.push([p.no, p.name, p.material, p.w, p.l, p.t, p.qty,
      `${c.ton}톤`, c.laborPerPc, c.expensePerPc, c.processCost,
      c.mgmtCost, c.profit, c.unitPrice, c.totalMoldCost]);
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryData), "요약");

  const qtyData = [["연도", "수량"], ...Object.entries(project.annualQty || {}).map(([y, q]) => [y, parseInt(q) || 0])];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(qtyData), "연도별수량");

  project.parts.forEach(p => {
    const c = calcPartCost(p, sopYear, totalQty);
    if (!c) return;
    const d = [
      ["부품 원가 계산서"],
      ["부품명", p.name, "", "재질", p.material],
      ["W(mm)", p.w, "", "L(mm)", p.l, "", "T(mm)", p.t],
      ["톤수", `${c.ton}톤`, "", "수량(개)", p.qty],
      [],
      ["── 가공비 ──"],
      ["항목", "금액(원)", "산출근거"],
      ["직접노무비", c.laborPerPc, `임률 ${c.wagePerHour}원/h × 사이클 ${c.cycleTimeSec}초`],
      ["경비", c.expensePerPc, "설비감가+전력+유지보수+공간비"],
      ["세척비", c.washCost, "세척 공정"],
      ["FLUX 도포비", c.fluxCost, c.rule.flux ? "FLUX 도포 적용" : "해당없음"],
      ["단품검사비", c.inspectionCost, c.rule.inspection ? "검사 적용" : "해당없음"],
      ["소계", c.subTotal, ""],
      ["일반관리비(25%)", c.mgmtCost, "(노무비+경비) × 25%"],
      ["이윤(15%)", c.profit, "(노무비+경비+관리비) × 15%"],
      ["단가 합계", c.unitPrice, ""],
      [],
      ["── 금형비 ──"],
      ["금형비(총)", c.totalMoldCost, ""],
      ["금형비 상각(개당)", c.moldAmortizationPerPc, `총수량 ${totalQty}개 × 5년 기준`],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(d), p.name.substring(0, 31));
  });

  XLSX.writeFile(wb, `${project.name}_원가계산서.xlsx`);
}

function CostDetail({ part, sopYear, totalQty }) {
  const c = calcPartCost(part, sopYear, totalQty);
  if (!c) return <p style={{ color: "#9ca3af", fontSize: 12 }}>W/L/T 값이 없어 계산 불가</p>;
  const rows = [
    ["직접노무비", c.laborPerPc],
    ["경비", c.expensePerPc],
    ["세척비", c.washCost],
    ["FLUX 도포비", c.fluxCost],
    ["단품검사비", c.inspectionCost],
    ["소계", c.subTotal],
    ["일반관리비 (25%)", c.mgmtCost],
    ["이윤 (15%)", c.profit],
    ["단가", c.unitPrice],
    ["금형비(총)", c.totalMoldCost],
    ["금형비 상각(개당)", c.moldAmortizationPerPc],
  ];
  return (
    <div style={{ fontSize: 12, marginTop: 10, borderTop: "1px solid #e5e7eb", paddingTop: 10 }}>
      <div style={{ display: "flex", gap: 10, marginBottom: 6, color: "#1d4ed8", fontWeight: 700 }}>
        <span>{c.ton}톤 PRESS</span>
        <span>사이클 {c.cycleTimeSec}초</span>
        <span>{c.pcsPerHour} pcs/h</span>
        {c.rule.flux && <span style={{ color: "#16a34a" }}>✓ FLUX</span>}
        {c.rule.inspection && <span style={{ color: "#ea580c" }}>✓ 단품검사</span>}
      </div>
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <tbody>
          {rows.map(([label, val]) => (
            <tr key={label} style={{ borderBottom: label === "단가" || label === "소계" ? "2px solid #d1d5db" : "1px solid #f3f4f6", fontWeight: label === "단가" ? 700 : 400 }}>
              <td style={{ padding: "3px 6px", color: "#6b7280" }}>{label}</td>
              <td style={{ padding: "3px 6px", textAlign: "right", fontFamily: "monospace" }}>{val?.toLocaleString()} 원</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export default function App() {
  const [projects, setProjects] = useState([]);
  const [activeId, setActiveId] = useState(null);
  const [showNew, setShowNew] = useState(false);
  const [newName, setNewName] = useState("");
  const [loading, setLoading] = useState(false);
  const [expandedPart, setExpandedPart] = useState(null);

  const active = projects.find(p => p.id === activeId);
  const totalQty = active
    ? Object.values(active.annualQty || {}).reduce((a, b) => a + (parseInt(b) || 0), 0)
    : 0;

  const addProject = () => {
    if (!newName.trim()) return;
    const id = Date.now().toString();
    setProjects(prev => [...prev, { id, name: newName.trim(), sopYear: 2026, annualQty: { ...DEFAULT_ANNUAL_QTY }, parts: [] }]);
    setActiveId(id);
    setNewName(""); setShowNew(false);
  };

  const updateProject = useCallback((id, updater) => {
    setProjects(prev => prev.map(p => p.id === id ? updater(p) : p));
  }, []);

  const handleImageUpload = async e => {
    const file = e.target.files[0];
    if (!file || !active) return;
    setLoading(true);
    try {
      const b64 = await new Promise((res, rej) => {
        const r = new FileReader();
        r.onload = () => res(r.result.split(",")[1]);
        r.onerror = rej;
        r.readAsDataURL(file);
      });
      const parts = await parseBomWithAI(b64, file.type);
      updateProject(active.id, p => ({ ...p, parts }));
    } catch (err) {
      alert("BOM 인식 실패: " + err.message);
    }
    setLoading(false);
    e.target.value = "";
  };

  const updatePart = (idx, field, val) => {
    updateProject(active.id, p => {
      const parts = [...p.parts];
      parts[idx] = { ...parts[idx], [field]: val };
      return { ...p, parts };
    });
  };

  const s = {
    sidebar: { width: 220, background: "#1e2b3c", color: "#fff", padding: 16, display: "flex", flexDirection: "column", minHeight: "100vh" },
    main: { flex: 1, padding: 24, overflowY: "auto", background: "#f4f6fb", minHeight: "100vh" },
    card: { background: "#fff", borderRadius: 10, padding: 16, boxShadow: "0 1px 4px rgba(0,0,0,0.07)", marginBottom: 14 },
    btn: (bg) => ({ padding: "7px 14px", background: bg, color: "#fff", border: "none", borderRadius: 6, fontSize: 12, fontWeight: 700, cursor: "pointer" }),
    input: { padding: "5px 8px", border: "1px solid #d1d5db", borderRadius: 5, fontSize: 13 },
  };

  return (
    <div style={{ fontFamily: "'Malgun Gothic', sans-serif", display: "flex" }}>
      {/* 사이드바 */}
      <div style={s.sidebar}>
        <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 16, color: "#7eb8f7", letterSpacing: 1 }}>⚙ 원가산출 시스템</div>
        <div style={{ fontSize: 11, color: "#8899aa", marginBottom: 8 }}>프로젝트 목록</div>
        <div style={{ flex: 1, overflowY: "auto" }}>
          {projects.map(p => (
            <div key={p.id} onClick={() => setActiveId(p.id)} style={{ padding: "8px 10px", borderRadius: 6, marginBottom: 6, cursor: "pointer", fontSize: 12, background: activeId === p.id ? "#2563eb" : "#263447", borderLeft: activeId === p.id ? "3px solid #60a5fa" : "3px solid transparent" }}>
              <div style={{ fontWeight: 600 }}>{p.name}</div>
              <div style={{ color: "#8899aa", fontSize: 10 }}>SOP {p.sopYear} | {p.parts.length}부품</div>
              <button onClick={e => { e.stopPropagation(); setProjects(prev => prev.filter(x => x.id !== p.id)); if (activeId === p.id) setActiveId(null); }} style={{ color: "#cc4444", background: "none", border: "none", cursor: "pointer", fontSize: 10, marginTop: 2 }}>삭제</button>
            </div>
          ))}
        </div>
        {showNew ? (
          <div>
            <input autoFocus value={newName} onChange={e => setNewName(e.target.value)} onKeyDown={e => e.key === "Enter" && addProject()} placeholder="프로젝트명" style={{ width: "100%", padding: "6px 8px", borderRadius: 4, border: "1px solid #3a4a5c", background: "#263447", color: "#fff", fontSize: 12, boxSizing: "border-box", marginBottom: 6 }} />
            <div style={{ display: "flex", gap: 4 }}>
              <button onClick={addProject} style={{ flex: 1, padding: "5px 0", background: "#2563eb", color: "#fff", border: "none", borderRadius: 4, cursor: "pointer", fontSize: 11 }}>추가</button>
              <button onClick={() => setShowNew(false)} style={{ flex: 1, padding: "5px 0", background: "#3a4a5c", color: "#ccc", border: "none", borderRadius: 4, cursor: "pointer", fontSize: 11 }}>취소</button>
            </div>
          </div>
        ) : (
          <button onClick={() => setShowNew(true)} style={{ width: "100%", padding: "8px 0", background: "#2563eb", color: "#fff", border: "none", borderRadius: 6, cursor: "pointer", fontSize: 12, marginTop: 8 }}>+ 새 프로젝트</button>
        )}
      </div>

      {/* 메인 */}
      <div style={s.main}>
        {!active ? (
          <div style={{ textAlign: "center", marginTop: 80, color: "#9ca3af" }}>
            <div style={{ fontSize: 48, marginBottom: 16 }}>📁</div>
            <p style={{ fontSize: 16 }}>좌측에서 프로젝트를 선택하거나 새로 만들어주세요</p>
          </div>
        ) : (
          <>
            {/* 헤더 카드 */}
            <div style={s.card}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 }}>
                <div>
                  <h2 style={{ margin: 0, fontSize: 18, color: "#1e2b3c" }}>{active.name}</h2>
                  <p style={{ margin: "4px 0 0", color: "#6b7280", fontSize: 12 }}>임률: {(WAGE_TABLE[active.sopYear] || 0).toLocaleString()}원/시간</p>
                </div>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <label style={{ fontSize: 12, fontWeight: 700 }}>SOP 연도</label>
                    <select value={active.sopYear} onChange={e => updateProject(active.id, p => ({ ...p, sopYear: parseInt(e.target.value) }))} style={s.input}>
                      {YEARS.map(y => <option key={y} value={y}>{y}년</option>)}
                    </select>
                  </div>
                  <label style={{ ...s.btn("#2563eb"), cursor: "pointer" }}>
                    {loading ? "인식 중..." : "📷 BOM 업로드"}
                    <input type="file" accept="image/*" onChange={handleImageUpload} style={{ display: "none" }} disabled={loading} />
                  </label>
                  <button onClick={() => downloadExcel(active, totalQty)} style={s.btn("#16a34a")}>⬇ 엑셀 다운로드</button>
                </div>
              </div>

              {/* 연도별 수량 */}
              <div style={{ marginTop: 16, borderTop: "1px solid #f3f4f6", paddingTop: 14 }}>
                <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 8 }}>📦 연도별 생산 수량 (개/년)</div>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                  {YEARS.filter(y => y >= 2025).map(y => (
                    <div key={y} style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 3 }}>
                      <label style={{ fontSize: 10, color: "#6b7280" }}>{y}년</label>
                      <input type="number" value={active.annualQty?.[y] ?? ""} onChange={e => updateProject(active.id, p => ({ ...p, annualQty: { ...p.annualQty, [y]: e.target.value } }))} style={{ ...s.input, width: 72, textAlign: "center", fontSize: 12 }} placeholder="0" />
                    </div>
                  ))}
                  <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 3 }}>
                    <label style={{ fontSize: 10, color: "#2563eb", fontWeight: 700 }}>총계</label>
                    <div style={{ width: 72, padding: "5px 6px", border: "2px solid #2563eb", borderRadius: 5, fontSize: 12, textAlign: "center", fontWeight: 700, color: "#2563eb", background: "#eff6ff" }}>{totalQty.toLocaleString()}</div>
                  </div>
                </div>
              </div>
            </div>

            {/* 부품 카드 */}
            {active.parts.length === 0 ? (
              <div style={{ textAlign: "center", padding: 60, color: "#9ca3af", background: "#fff", borderRadius: 10 }}>
                <div style={{ fontSize: 36, marginBottom: 12 }}>📋</div>
                <p>BOM 이미지를 업로드하면 자동으로 부품이 추출됩니다</p>
              </div>
            ) : (
              <>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(340px, 1fr))", gap: 14 }}>
                  {active.parts.map((part, idx) => {
                    const cost = calcPartCost(part, active.sopYear, totalQty);
                    return (
                      <div key={idx} style={{ ...s.card, marginBottom: 0, border: expandedPart === idx ? "2px solid #2563eb" : "2px solid transparent" }}>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                            <span style={{ background: "#2563eb", color: "#fff", borderRadius: "50%", width: 22, height: 22, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, fontWeight: 700 }}>{part.no}</span>
                            <span style={{ fontWeight: 700, fontSize: 13 }}>{part.name}</span>
                          </div>
                          {cost && <span style={{ background: "#f0fdf4", color: "#16a34a", fontSize: 11, fontWeight: 700, padding: "2px 8px", borderRadius: 12 }}>{cost.unitPrice.toLocaleString()}원</span>}
                        </div>
                        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 6, marginBottom: 8 }}>
                          {["w", "l", "t", "qty"].map(field => (
                            <div key={field} style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                              <label style={{ fontSize: 9, color: "#6b7280", textTransform: "uppercase" }}>{field}</label>
                              <input type="number" value={part[field] ?? ""} onChange={e => updatePart(idx, field, parseFloat(e.target.value) || "")} style={{ ...s.input, textAlign: "center", fontSize: 12, padding: "4px 5px" }} />
                            </div>
                          ))}
                        </div>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                          <span style={{ fontSize: 10, color: "#6b7280" }}>{part.material}</span>
                          <button onClick={() => setExpandedPart(expandedPart === idx ? null : idx)} style={{ fontSize: 11, color: "#2563eb", background: "none", border: "none", cursor: "pointer", fontWeight: 700 }}>{expandedPart === idx ? "▲ 접기" : "▼ 상세보기"}</button>
                        </div>
                        {expandedPart === idx && <CostDetail part={part} sopYear={active.sopYear} totalQty={totalQty} />}
                      </div>
                    );
                  })}
                </div>

                {/* 전체 요약 */}
                <div style={{ background: "#1e2b3c", color: "#fff", borderRadius: 10, padding: 20, marginTop: 16 }}>
                  <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 12, color: "#7eb8f7" }}>📊 전체 원가 요약</div>
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ width: "100%", fontSize: 12, borderCollapse: "collapse", minWidth: 700 }}>
                      <thead>
                        <tr style={{ borderBottom: "1px solid #3a4a5c" }}>
                          {["NO", "부품명", "톤수", "노무비", "경비", "공정비", "관리비", "이윤", "단가", "금형비"].map(h => (
                            <th key={h} style={{ padding: "6px 8px", textAlign: h === "NO" || h === "부품명" ? "left" : "right", color: "#8899aa" }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {active.parts.map((p, i) => {
                          const c = calcPartCost(p, active.sopYear, totalQty);
                          if (!c) return null;
                          return (
                            <tr key={i} style={{ borderBottom: "1px solid #263447" }}>
                              <td style={{ padding: "6px 8px" }}>{p.no}</td>
                              <td style={{ padding: "6px 8px" }}>{p.name}</td>
                              <td style={{ padding: "6px 8px", textAlign: "right", color: "#60a5fa" }}>{c.ton}T</td>
                              <td style={{ padding: "6px 8px", textAlign: "right" }}>{c.laborPerPc.toLocaleString()}</td>
                              <td style={{ padding: "6px 8px", textAlign: "right" }}>{c.expensePerPc.toLocaleString()}</td>
                              <td style={{ padding: "6px 8px", textAlign: "right" }}>{c.processCost.toLocaleString()}</td>
                              <td style={{ padding: "6px 8px", textAlign: "right" }}>{c.mgmtCost.toLocaleString()}</td>
                              <td style={{ padding: "6px 8px", textAlign: "right" }}>{c.profit.toLocaleString()}</td>
                              <td style={{ padding: "6px 8px", textAlign: "right", fontWeight: 700, color: "#4ade80" }}>{c.unitPrice.toLocaleString()}</td>
                              <td style={{ padding: "6px 8px", textAlign: "right", color: "#fbbf24" }}>{c.totalMoldCost.toLocaleString()}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </>
            )}
          </>
        )}
      </div>
    </div>
  );
}
