import { useState, useEffect } from "react";
import { loadS, saveS } from "./firebase";
import * as XLSX from "xlsx";

const DEFAULT_TEAM_PW = "1234";
const DEFAULT_ADMIN_PW = "admin1234";

const GIFT_TYPES = [
  { id:"입사기념일", label:"입사기념일", amounts:[20000,20000,20000,20000,20000] },
  { id:"생일", label:"생일", amounts:[20000,20000,20000,20000,20000] },
  { id:"자녀출산", label:"자녀출산", amounts:[40000,80000,120000,160000,200000] },
  { id:"자녀돌", label:"자녀돌", amounts:[50000,50000,50000,50000,50000] },
  { id:"퇴사", label:"퇴사(2년이상)", amounts:[null,60000,90000,120000,150000] },
  { id:"결혼_본인", label:"결혼-본인", amounts:[100000,200000,300000,400000,500000] },
  { id:"결혼_자녀", label:"결혼-자녀", amounts:[50000,100000,150000,200000,250000] },
  { id:"수연_본인", label:"수연-본인(회갑)", amounts:[20000,40000,60000,80000,100000] },
  { id:"수연_부모", label:"수연-부모", amounts:[20000,40000,60000,80000,100000] },
  { id:"수연_부모배우자", label:"수연-부모(배우자)", amounts:[20000,40000,60000,80000,100000] },
  { id:"고희_부모", label:"고희-부모", amounts:[20000,40000,60000,80000,100000] },
  { id:"고희_부모배우자", label:"고희-부모(배우자)", amounts:[20000,40000,60000,80000,100000] },
  { id:"팔순_부모", label:"팔순-부모", amounts:[20000,40000,60000,80000,100000] },
  { id:"팔순_부모배우자", label:"팔순-부모(배우자)", amounts:[20000,40000,60000,80000,100000] },
  { id:"조위_본인", label:"조위-본인", amounts:[1000000,1000000,1000000,1000000,1000000] },
  { id:"조위_부모", label:"조위-부모", amounts:[200000,200000,200000,200000,200000] },
  { id:"조위_외조부모", label:"조위-(외)조부모", amounts:[100000,100000,100000,100000,100000] },
  { id:"조위_형제자매", label:"조위-형제자매", amounts:[100000,100000,100000,100000,100000] },
  { id:"조위_자녀", label:"조위-자녀", amounts:[500000,500000,500000,500000,500000] },
  { id:"조위_배우자", label:"조위-배우자", amounts:[500000,500000,500000,500000,500000] },
  { id:"조위_부모배우자", label:"조위-부모(배우자)", amounts:[200000,200000,200000,200000,200000] },
];

const BANKS = ["국민은행","신한은행","우리은행","하나은행","IBK기업은행","농협은행","SC제일은행","씨티은행","카카오뱅크","케이뱅크","토스뱅크","대구은행","부산은행","경남은행","광주은행","전북은행","제주은행","수협은행","산업은행","우체국"];
const BANK_CODES = {
  "국민은행":"004","신한은행":"088","우리은행":"020","하나은행":"081",
  "IBK기업은행":"003","농협은행":"011","SC제일은행":"023","씨티은행":"027",
  "카카오뱅크":"090","케이뱅크":"089","토스뱅크":"092","대구은행":"031",
  "부산은행":"032","경남은행":"039","광주은행":"034","전북은행":"037",
  "제주은행":"035","수협은행":"007","산업은행":"002","우체국":"071"
};
const NO_PROOF = new Set(["입사기념일","생일","퇴사"]);

// ── Helpers ──
function calcYrs(hd) {
  if (!hd) return 1;
  const y = Math.floor((Date.now() - new Date(hd)) / (365.25 * 86400000));
  return Math.max(1, Math.min(y < 1 ? 1 : y, 5));
}
function yrsLabel(y) { return y >= 5 ? "5년차 이상" : y + "년차"; }
function yrsExcel(y) { return y >= 5 ? "5년 이상" : y + "년"; }
function getAmt(gid, yrs) {
  const g = GIFT_TYPES.find(x => x.id === gid);
  if (!g) return 0;
  return g.amounts[Math.min(Math.max(yrs, 1), 5) - 1] || 0;
}
function fmtMoney(n) { return n != null ? Number(n).toLocaleString() + "원" : "-"; }
function nowYM() {
  const d = new Date();
  return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0");
}
function lastYM() {
  const d = new Date();
  d.setMonth(d.getMonth() - 1);
  return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0");
}
function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 5); }
function fmtED(s) {
  if (!s) return "";
  const d = new Date(s);
  return isNaN(d.getTime()) ? s : (d.getMonth() + 1) + "/" + d.getDate();
}
async function loadS(k) {
  try { const r = await window.storage.get(k); return r ? JSON.parse(r.value) : null; }
  catch { return null; }
}
async function saveS(k, v) {
  try { await window.storage.set(k, JSON.stringify(v)); } catch (e) { console.error(e); }
}
// 마감 여부 체크: deadlineDay(매월 n일)가 지났으면 잠금
function isDeadlinePassed(deadlineDay) {
  if (!deadlineDay) return false;
  return new Date().getDate() > Number(deadlineDay);
}

// ── Logo ──
function Logo() {
  return (
    <div style={{ fontFamily: "Arial,sans-serif", fontSize: "20px", fontWeight: "700", color: "#0057D9", letterSpacing: "1px" }}>
      메가스터디교육
    </div>
  );
}

// ── DatePicker ──
function DatePicker({ value, onChange, startYear, endYear }) {
  const sy = startYear || 1950;
  const ey = endYear || new Date().getFullYear();
  function parse(v) {
    if (v && /^\d{4}-\d{2}-\d{2}$/.test(v)) { const p = v.split("-"); return [p[0], p[1], p[2]]; }
    return ["", "", ""];
  }
  const [parts, setParts] = useState(() => parse(value));
  useEffect(() => { setParts(parse(value)); }, [value]);
  const yr = parts[0], mo = parts[1], dy = parts[2];
  const maxD = yr && mo ? new Date(Number(yr), Number(mo), 0).getDate() : 31;
  function upd(i, v) {
    const np = [parts[0], parts[1], parts[2]]; np[i] = v;
    if (np[0] && np[1]) { const md = new Date(Number(np[0]), Number(np[1]), 0).getDate(); if (np[2] && Number(np[2]) > md) np[2] = String(md).padStart(2, "0"); }
    setParts(np);
    if (np[0] && np[1] && np[2]) onChange(np[0] + "-" + np[1] + "-" + np[2]);
  }
  const sc = "border border-gray-200 rounded-xl py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-300";
  const years = [];
  for (let i = ey; i >= sy; i--) years.push(String(i));
  return (
    <div style={{ display: "flex", gap: 4 }}>
      <select value={yr} onChange={e => upd(0, e.target.value)} style={{ flex: 1 }} className={sc + " px-2"}>
        <option value="">연도</option>
        {years.map(y => <option key={y} value={y}>{y}년</option>)}
      </select>
      <select value={mo} onChange={e => upd(1, e.target.value)} style={{ width: 66 }} className={sc + " px-1"}>
        <option value="">월</option>
        {[1,2,3,4,5,6,7,8,9,10,11,12].map(m => <option key={m} value={String(m).padStart(2,"0")}>{m}월</option>)}
      </select>
      <select value={dy} onChange={e => upd(2, e.target.value)} style={{ width: 66 }} className={sc + " px-1"}>
        <option value="">일</option>
        {Array.from({ length: maxD }, (_, i) => i + 1).map(d => <option key={d} value={String(d).padStart(2,"0")}>{d}일</option>)}
      </select>
    </div>
  );
}

// ── Confirm Dialog ──
function ConfirmDlg({ msg, onYes, onNo }) {
  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-2xl p-6 w-full max-w-xs text-center shadow-2xl">
        <div className="text-3xl mb-3">🗑️</div>
        <p className="text-sm text-gray-700 mb-5 leading-relaxed">{msg}</p>
        <div className="flex gap-2">
          <button onClick={onNo} className="flex-1 py-2 bg-gray-100 text-gray-600 rounded-xl text-sm font-semibold">취소</button>
          <button onClick={onYes} className="flex-1 py-2 bg-red-500 text-white rounded-xl text-sm font-semibold">삭제</button>
        </div>
      </div>
    </div>
  );
}

// ── Change Password Modal ──
function ChangePasswordModal({ currentPw, onSave, onClose }) {
  const [cur, setCur] = useState(""), [nw, setNw] = useState(""), [cf, setCf] = useState("");
  const [err, setErr] = useState(""), [done, setDone] = useState(false);
  function submit() {
    if (cur !== currentPw) { setErr("현재 비밀번호가 틀렸습니다."); return; }
    if (nw.length < 4) { setErr("새 비밀번호는 4자 이상이어야 합니다."); return; }
    if (nw !== cf) { setErr("새 비밀번호가 일치하지 않습니다."); return; }
    onSave(nw); setDone(true); setTimeout(() => onClose(), 1200);
  }
  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-3xl p-6 w-full max-w-sm shadow-2xl">
        <div className="flex justify-between items-center mb-5">
          <h3 className="font-bold text-gray-800">🔑 비밀번호 변경</h3>
          <button onClick={onClose} className="text-gray-300 hover:text-gray-500 text-2xl leading-none">✕</button>
        </div>
        {done ? (
          <div className="text-center py-6"><div className="text-4xl mb-2">✅</div><p className="text-sm font-semibold text-green-600">비밀번호가 변경되었습니다.</p></div>
        ) : (
          <div>
            {[["현재 비밀번호", cur, setCur], ["새 비밀번호", nw, setNw], ["새 비밀번호 확인", cf, setCf]].map(function(item) {
              return (
                <div key={item[0]} className="mb-3">
                  <label className="block text-xs font-semibold text-gray-500 mb-1.5">{item[0]}</label>
                  <input type="password" value={item[1]} onChange={e => { item[2](e.target.value); setErr(""); }}
                    className="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-300" />
                </div>
              );
            })}
            {err && <p className="text-red-500 text-xs mb-3">{err}</p>}
            <div className="flex gap-2 mt-4">
              <button onClick={onClose} className="flex-1 py-2.5 bg-gray-100 text-gray-600 rounded-xl text-sm font-bold">취소</button>
              <button onClick={submit} className="flex-1 py-2.5 bg-blue-500 text-white rounded-xl text-sm font-bold hover:bg-blue-600">변경</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ── App Root ──
export default function App() {
  const [screen, setScreen] = useState("login");
  const [teams, setTeams] = useState([]);
  const [apps, setApps] = useState([]);
  const [teamId, setTeamId] = useState(null);
  const [adminPw, setAdminPw] = useState(DEFAULT_ADMIN_PW);
  const [deadlineDay, setDeadlineDay] = useState(null); // 매월 마감일 (일)
  const [ready, setReady] = useState(false);

  useEffect(() => {
    (async () => {
      const t = await loadS("sw:teams"), a = await loadS("sw:apps");
      const ap = await loadS("sw:adminPw"), dd = await loadS("sw:deadlineDay");
      if (t) setTeams(t); if (a) setApps(a);
      if (ap) setAdminPw(ap); if (dd) setDeadlineDay(dd);
      setReady(true);
    })();
  }, []);

  async function updTeams(t) { setTeams(t); await saveS("sw:teams", t); }
  async function updApps(a) { setApps(a); await saveS("sw:apps", a); }
  async function updAdminPw(pw) { setAdminPw(pw); await saveS("sw:adminPw", pw); }
  async function updDeadlineDay(d) { setDeadlineDay(d); await saveS("sw:deadlineDay", d); }

  if (!ready) return <div className="flex items-center justify-center h-screen text-gray-400 text-sm" style={{ fontFamily: "'Noto Sans KR',sans-serif" }}>불러오는 중...</div>;
  const team = teams.find(t => t.id === teamId);

  return (
    <div style={{ fontFamily: "'Noto Sans KR',sans-serif" }}>
      {screen === "login" && <LoginScreen adminPw={adminPw} onStaff={() => setScreen("staffTeam")} onAdmin={() => setScreen("admin")} />}
      {screen === "staffTeam" && <TeamPicker teams={teams} updTeams={updTeams} onEnter={id => { setTeamId(id); setScreen("staffPw"); }} onBack={() => setScreen("login")} />}
      {screen === "staffPw" && team && <TeamPasswordScreen team={team} onEnter={() => setScreen("staff")} onBack={() => setScreen("staffTeam")} />}
      {screen === "staff" && team && <StaffScreen team={team} teams={teams} updTeams={updTeams} apps={apps} updApps={updApps} deadlineDay={deadlineDay} onLogout={() => { setScreen("login"); setTeamId(null); }} />}
      {screen === "admin" && <AdminScreen adminPw={adminPw} updAdminPw={updAdminPw} deadlineDay={deadlineDay} updDeadlineDay={updDeadlineDay} teams={teams} updTeams={updTeams} apps={apps} updApps={updApps} onLogout={() => setScreen("login")} />}
    </div>
  );
}

// ── Login ──
function LoginScreen({ adminPw, onStaff, onAdmin }) {
  const [mode, setMode] = useState("staff"), [pw, setPw] = useState(""), [err, setErr] = useState("");
  function submit() {
    if (mode === "staff") { onStaff(); return; }
    if (pw === adminPw) { onAdmin(); return; }
    setErr("비밀번호가 맞지 않습니다.");
  }
  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 flex items-center justify-center p-4">
      <div className="bg-white rounded-3xl shadow-xl p-8 w-full max-w-sm">
        <div className="text-center mb-7">
          <div className="flex justify-center mb-3"><Logo /></div>
          <h1 className="text-xl font-bold text-gray-800">사우회 경조금 관리</h1>
          <p className="text-xs text-gray-400 mt-1">담당자 또는 관리자로 로그인하세요</p>
        </div>
        <div className="flex gap-2 mb-5">
          {[["staff","👥 팀 담당자"],["admin","🔑 관리자"]].map(([m, l]) => (
            <button key={m} onClick={() => { setMode(m); setErr(""); setPw(""); }}
              className={"flex-1 py-2.5 rounded-xl text-sm font-bold transition-all " + (mode === m ? "bg-blue-500 text-white shadow-md" : "bg-gray-100 text-gray-500")}>{l}</button>
          ))}
        </div>
        {mode === "admin" ? (
          <div className="mb-4">
            <label className="block text-xs font-semibold text-gray-500 mb-1.5">비밀번호</label>
            <input type="password" value={pw} onChange={e => { setPw(e.target.value); setErr(""); }} onKeyDown={e => e.key === "Enter" && submit()}
              className="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-300" placeholder="관리자 비밀번호 입력" />
          </div>
        ) : (
          <div className="mb-4 bg-blue-50 rounded-xl p-4 text-center">
            <p className="text-xs text-blue-600">팀 선택 후 팀별 비밀번호로 접속합니다.</p>
          </div>
        )}
        {err && <p className="text-red-500 text-xs mb-3">{err}</p>}
        <button onClick={submit} className="w-full bg-blue-500 text-white py-3 rounded-xl font-bold text-sm hover:bg-blue-600 shadow-md">
          {mode === "staff" ? "팀 선택하기" : "로그인"}
        </button>
      </div>
    </div>
  );
}

// ── Team Picker ──
function TeamPicker({ teams, updTeams, onEnter, onBack }) {
  const [search, setSearch] = useState(""), [selectedId, setSelectedId] = useState(""), [remember, setRemember] = useState(true);
  const [showManual, setShowManual] = useState(false), [loading, setLoading] = useState(true);
  const [bName, setBName] = useState(""), [bDiv, setBDiv] = useState(""), [bTeam, setBTeam] = useState(""), [bErr, setBErr] = useState("");

  useEffect(() => { (async () => { const last = await loadS("sw:lastTeamId"); if (last && teams.find(t => t.id === last)) setSelectedId(last); setLoading(false); })(); }, []);

  const filtered = teams.filter(t => { const q = search.toLowerCase().trim(); if (!q) return true; return t.팀명.includes(q) || t.본부명.includes(q) || (t.부문명 && t.부문명.includes(q)); });

  async function enter() { if (!selectedId) return; if (remember) await saveS("sw:lastTeamId", selectedId); onEnter(selectedId); }
  async function manualSubmit() {
    if (!bName.trim() || !bTeam.trim()) { setBErr("본부명과 팀명은 필수입니다."); return; }
    const ex = teams.find(t => t.본부명 === bName.trim() && t.부문명 === bDiv.trim() && t.팀명 === bTeam.trim());
    if (ex) { if (remember) await saveS("sw:lastTeamId", ex.id); onEnter(ex.id); return; }
    const nt = { id: uid(), 본부명: bName.trim(), 부문명: bDiv.trim(), 팀명: bTeam.trim(), members: [] };
    await updTeams([...teams, nt]); if (remember) await saveS("sw:lastTeamId", nt.id); onEnter(nt.id);
  }

  if (loading) return <div className="min-h-screen flex items-center justify-center text-gray-400 text-sm">불러오는 중...</div>;
  const selTeam = teams.find(t => t.id === selectedId);

  if (showManual) return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 flex items-center justify-center p-4">
      <div className="bg-white rounded-3xl shadow-xl p-8 w-full max-w-sm">
        <button onClick={() => setShowManual(false)} className="text-xs text-gray-400 hover:text-gray-600 mb-5 block">← 목록으로</button>
        <h2 className="text-lg font-bold text-gray-800 mb-1">팀 직접 입력</h2>
        <p className="text-xs text-gray-400 mb-5">새 팀을 등록하거나 기존 팀에 접근합니다.</p>
        {[["본부명 *", bName, setBName, "예: 교육사업본부"],["부문(실)명", bDiv, setBDiv, "없으면 생략"],["팀명 *", bTeam, setBTeam, "예: 기획팀"]].map(([lbl, val, fn, ph]) => (
          <div key={lbl} className="mb-3"><label className="block text-xs font-semibold text-gray-500 mb-1.5">{lbl}</label>
            <input value={val} onChange={e => { fn(e.target.value); setBErr(""); }} className="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-300" placeholder={ph} />
          </div>
        ))}
        {bErr && <p className="text-red-500 text-xs mb-3">{bErr}</p>}
        <label className="flex items-center gap-2 text-xs text-gray-500 mb-4 cursor-pointer select-none">
          <input type="checkbox" checked={remember} onChange={e => setRemember(e.target.checked)} className="w-3.5 h-3.5" />다음 방문 시 자동 선택
        </label>
        <button onClick={manualSubmit} className="w-full bg-blue-500 text-white py-3 rounded-xl font-bold text-sm hover:bg-blue-600 shadow-md">확인 / 입장</button>
      </div>
    </div>
  );

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 flex items-center justify-center p-4">
      <div className="bg-white rounded-3xl shadow-xl p-6 w-full max-w-sm">
        <button onClick={onBack} className="text-xs text-gray-400 hover:text-gray-600 mb-4 block">← 뒤로</button>
        <h2 className="text-lg font-bold text-gray-800 mb-0.5">팀 선택</h2>
        <p className="text-xs text-gray-400 mb-4">소속 팀을 선택하거나 검색하세요.</p>
        {teams.length === 0 ? (
          <div className="text-center py-8 text-gray-400"><div className="text-3xl mb-2">📋</div><p className="text-xs">등록된 팀이 없습니다.</p>
            <button onClick={() => setShowManual(true)} className="text-blue-500 text-xs mt-2 underline">직접 팀을 등록해주세요.</button></div>
        ) : (
          <div>
            <input value={search} onChange={e => setSearch(e.target.value)} className="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-300 mb-3" placeholder="🔍  팀명, 본부명으로 검색..." />
            <div style={{ maxHeight: 220, overflowY: "auto" }} className="space-y-1 mb-4 pr-1">
              {filtered.length === 0 ? <div className="text-center text-gray-400 text-xs py-6">검색 결과가 없습니다.</div> :
                filtered.map(t => (
                  <button key={t.id} onClick={() => setSelectedId(t.id)}
                    className={"w-full text-left px-3 py-2.5 rounded-xl text-sm transition-all " + (selectedId === t.id ? "bg-blue-500 text-white shadow-sm" : "bg-gray-50 text-gray-700 hover:bg-blue-50")}>
                    <div className="font-semibold">{t.팀명}</div>
                    <div className={"text-xs mt-0.5 " + (selectedId === t.id ? "text-blue-100" : "text-gray-400")}>{[t.본부명, t.부문명].filter(Boolean).join(" · ")}</div>
                  </button>
                ))}
            </div>
          </div>
        )}
        <label className="flex items-center gap-2 text-xs text-gray-500 mb-3 cursor-pointer select-none">
          <input type="checkbox" checked={remember} onChange={e => setRemember(e.target.checked)} className="w-3.5 h-3.5" />다음 방문 시 이 팀을 자동으로 선택
        </label>
        {selTeam && <div className="bg-blue-50 rounded-xl px-3 py-2 text-xs text-blue-700 mb-3">선택: <span className="font-bold">{selTeam.팀명}</span><span className="text-blue-400 ml-1">({[selTeam.본부명, selTeam.부문명].filter(Boolean).join(" · ")})</span></div>}
        <button onClick={enter} disabled={!selectedId} className={"w-full py-3 rounded-xl font-bold text-sm shadow-sm " + (!selectedId ? "bg-gray-200 text-gray-400" : "bg-blue-500 text-white hover:bg-blue-600")}>다음 →</button>
        <button onClick={() => setShowManual(true)} className="w-full text-center text-xs text-gray-400 hover:text-blue-500 mt-3">목록에 팀이 없나요? 직접 입력 →</button>
      </div>
    </div>
  );
}

// ── Team Password Screen ──
function TeamPasswordScreen({ team, onEnter, onBack }) {
  const [pw, setPw] = useState(""), [err, setErr] = useState("");
  const teamPw = team.password || DEFAULT_TEAM_PW;
  function submit() { if (pw === teamPw) onEnter(); else setErr("비밀번호가 맞지 않습니다."); }
  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 flex items-center justify-center p-4">
      <div className="bg-white rounded-3xl shadow-xl p-8 w-full max-w-sm">
        <button onClick={onBack} className="text-xs text-gray-400 hover:text-gray-600 mb-5 block">← 뒤로</button>
        <div className="text-center mb-6">
          <div className="text-4xl mb-2">🔒</div>
          <h2 className="text-lg font-bold text-gray-800">{team.팀명}</h2>
          <p className="text-xs text-gray-400 mt-1">{[team.본부명, team.부문명].filter(Boolean).join(" · ")}</p>
        </div>
        {!team.password && <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 mb-4 text-xs text-amber-700">⚠️ 초기 비밀번호 <span className="font-bold">1234</span>로 설정되어 있습니다. 로그인 후 반드시 변경해주세요.</div>}
        <div className="mb-4">
          <label className="block text-xs font-semibold text-gray-500 mb-1.5">비밀번호</label>
          <input type="password" value={pw} onChange={e => { setPw(e.target.value); setErr(""); }} onKeyDown={e => e.key === "Enter" && submit()}
            className="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-300" placeholder="팀 비밀번호 입력" />
        </div>
        {err && <p className="text-red-500 text-xs mb-3">{err}</p>}
        <button onClick={submit} className="w-full bg-blue-500 text-white py-3 rounded-xl font-bold text-sm hover:bg-blue-600 shadow-md">로그인</button>
      </div>
    </div>
  );
}

// ── Staff Screen ──
function StaffScreen({ team, teams, updTeams, apps, updApps, deadlineDay, onLogout }) {
  const [tab, setTab] = useState("members"), [showPwModal, setShowPwModal] = useState(false);
  const locked = isDeadlinePassed(deadlineDay);
  async function updateTeam(u) { await updTeams(teams.map(t => t.id === u.id ? u : t)); }
  async function changeTeamPw(newPw) { await updateTeam({ ...team, password: newPw }); setShowPwModal(false); }

  return (
    <div className="min-h-screen bg-gray-50">
      <div className="bg-blue-500 text-white px-4 py-3 flex justify-between items-center" style={{ position: "sticky", top: 0, zIndex: 20 }}>
        <div>
          <div className="font-bold text-sm">{[team.본부명, team.부문명, team.팀명].filter(Boolean).join(" · ")}</div>
          <div className="text-xs flex items-center gap-2" style={{ opacity: .9 }}>
            <span>담당자 화면</span>
            {deadlineDay && (
              locked
                ? <span className="bg-red-500 text-white text-xs px-2 py-0.5 rounded-full font-bold">🔒 매월 {deadlineDay}일 마감</span>
                : <span className="bg-blue-400 text-white text-xs px-2 py-0.5 rounded-full">📅 매월 {deadlineDay}일 마감</span>
            )}
          </div>
        </div>
        <div className="flex gap-2">
          <button onClick={() => setShowPwModal(true)} className="text-xs bg-blue-600 hover:bg-blue-700 px-3 py-1.5 rounded-lg">🔑 비번변경</button>
          <button onClick={onLogout} className="text-xs bg-blue-600 hover:bg-blue-700 px-3 py-1.5 rounded-lg">로그아웃</button>
        </div>
      </div>
      <div className="bg-white border-b flex gap-1 px-4 py-2 overflow-x-auto" style={{ position: "sticky", top: 52, zIndex: 10 }}>
        {[["members","👤 팀원 관리"],["apply","✍️ 경조금 신청"],["history","📋 이번달 내역"]].map(([t, l]) => (
          <button key={t} onClick={() => setTab(t)} className={"whitespace-nowrap px-3 py-1.5 rounded-lg text-xs font-bold transition-all " + (tab === t ? "bg-blue-500 text-white" : "text-gray-400 hover:bg-gray-100")}>{l}</button>
        ))}
      </div>
      <div className="p-4 max-w-xl mx-auto">
        {tab === "members" && <MembersTab team={team} updateTeam={updateTeam} />}
        {tab === "apply" && <ApplyTab team={team} apps={apps} updApps={updApps} locked={locked} deadlineDay={deadlineDay} />}
        {tab === "history" && <HistoryTab team={team} apps={apps} updApps={updApps} locked={locked} />}
      </div>
      {showPwModal && <ChangePasswordModal currentPw={team.password || DEFAULT_TEAM_PW} onSave={changeTeamPw} onClose={() => setShowPwModal(false)} />}
    </div>
  );
}

// ── Members Tab ──
function MembersTab({ team, updateTeam }) {
  const [showForm, setShowForm] = useState(false), [editing, setEditing] = useState(null), [confirmDel, setConfirmDel] = useState(null);
  async function save(m) { const members = m.id ? team.members.map(x => x.id === m.id ? m : x) : [...team.members, { ...m, id: uid() }]; await updateTeam({ ...team, members }); setShowForm(false); setEditing(null); }
  async function del(id) { await updateTeam({ ...team, members: team.members.filter(x => x.id !== id) }); setConfirmDel(null); }
  return (
    <div>
      <div className="flex justify-between items-center mb-4">
        <h2 className="font-bold text-gray-800">팀원 목록 <span className="text-gray-400 font-normal">({team.members.length}명)</span></h2>
        <button onClick={() => { setEditing(null); setShowForm(true); }} className="bg-blue-500 text-white text-xs px-3 py-2 rounded-xl font-bold shadow-sm">+ 팀원 추가</button>
      </div>
      {team.members.length === 0 ? (
        <div className="text-center text-gray-400 text-sm py-14 bg-white rounded-2xl shadow-sm"><div className="text-4xl mb-2">👥</div>팀원을 추가해주세요.</div>
      ) : (
        <div className="space-y-2">
          {team.members.map(m => (
            <div key={m.id} className="bg-white rounded-2xl p-4 shadow-sm">
              <div className="flex justify-between items-start gap-3">
                <div className="min-w-0">
                  <div className="font-bold text-sm text-gray-800">{m.이름}{m.계좌명 && m.계좌명 !== m.이름 && <span className="text-gray-400 text-xs font-normal ml-1">({m.계좌명})</span>}</div>
                  <div className="text-xs text-gray-500 mt-0.5">{m.은행명 || "은행 미등록"} · {m.계좌번호 || "계좌 미등록"}</div>
                  <div className="text-xs text-gray-400 mt-0.5">생일: {m.생일 || "-"} · 입사일: {m.입사일 || "-"}</div>
                  {m.입사일 && <div className="text-xs text-blue-500 mt-0.5 font-semibold">{yrsLabel(calcYrs(m.입사일))}</div>}
                </div>
                <div className="flex gap-1 shrink-0">
                  <button onClick={() => { setEditing(m); setShowForm(true); }} className="text-xs bg-gray-100 hover:bg-gray-200 px-2.5 py-1.5 rounded-lg text-gray-600 font-semibold">수정</button>
                  <button onClick={() => setConfirmDel(m.id)} className="text-xs bg-red-50 hover:bg-red-100 px-2.5 py-1.5 rounded-lg text-red-500 font-semibold">삭제</button>
                </div>
              </div>
            </div>
          ))}
        </div>
      )}
      {showForm && <MemberModal member={editing} onSave={save} onClose={() => { setShowForm(false); setEditing(null); }} />}
      {confirmDel && <ConfirmDlg msg="이 팀원을 삭제할까요?" onYes={() => del(confirmDel)} onNo={() => setConfirmDel(null)} />}
    </div>
  );
}

// ── Member Modal ──
function MemberModal({ member, onSave, onClose }) {
  const [form, setForm] = useState(member || { 이름: "", 생일: "", 입사일: "", 은행명: "", 계좌번호: "", 계좌명: "" });
  function upd(k, v) { setForm(f => ({ ...f, [k]: v })); }
  function submit() { if (!form.이름.trim()) return; onSave(form); }
  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-3xl p-6 w-full max-w-sm max-h-screen overflow-y-auto shadow-2xl">
        <div className="flex justify-between items-center mb-5">
          <h3 className="font-bold text-gray-800">{member ? "팀원 수정" : "팀원 추가"}</h3>
          <button onClick={onClose} className="text-gray-300 hover:text-gray-500 text-2xl leading-none">✕</button>
        </div>
        <div className="mb-3"><label className="block text-xs font-semibold text-gray-500 mb-1.5">이름 *</label><input value={form.이름} onChange={e => upd("이름", e.target.value)} className="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-300" /></div>
        <div className="mb-3"><label className="block text-xs font-semibold text-gray-500 mb-1.5">생일</label><DatePicker value={form.생일} onChange={v => upd("생일", v)} startYear={1950} /></div>
        <div className="mb-3"><label className="block text-xs font-semibold text-gray-500 mb-1.5">입사기념일 (입사일)</label><DatePicker value={form.입사일} onChange={v => upd("입사일", v)} startYear={1990} /></div>
        <div className="mb-3"><label className="block text-xs font-semibold text-gray-500 mb-1.5">계좌 은행</label>
          <select value={form.은행명} onChange={e => upd("은행명", e.target.value)} className="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm focus:outline-none">
            <option value="">은행 선택</option>{BANKS.map(b => <option key={b}>{b}</option>)}
          </select>
        </div>
        <div className="mb-3"><label className="block text-xs font-semibold text-gray-500 mb-1.5">계좌번호</label><input value={form.계좌번호} onChange={e => upd("계좌번호", e.target.value)} className="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm focus:outline-none" placeholder="'-' 없이 입력" /></div>
        <div className="mb-5"><label className="block text-xs font-semibold text-gray-500 mb-1.5">계좌명 (예금주)</label><input value={form.계좌명} onChange={e => upd("계좌명", e.target.value)} className="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm focus:outline-none" placeholder="미입력 시 이름으로 사용" /></div>
        <div className="flex gap-2">
          <button onClick={onClose} className="flex-1 py-2.5 bg-gray-100 text-gray-600 rounded-xl text-sm font-bold">취소</button>
          <button onClick={submit} disabled={!form.이름.trim()} className={"flex-1 py-2.5 rounded-xl text-sm font-bold " + (!form.이름.trim() ? "bg-gray-200 text-gray-400" : "bg-blue-500 text-white hover:bg-blue-600")}>저장</button>
        </div>
      </div>
    </div>
  );
}

// ── Apply Tab ──
function ApplyTab({ team, apps, updApps, locked, deadlineDay }) {
  const cy = new Date().getFullYear();
  const [memberId, setMemberId] = useState(""), [giftId, setGiftId] = useState(""), [eventDate, setEventDate] = useState("");
  const [yrs, setYrs] = useState(1), [overrideAmt, setOverrideAmt] = useState(false), [amt, setAmt] = useState(0);
  const [memo, setMemo] = useState(""), [proofFile, setProofFile] = useState(null), [proofLater, setProofLater] = useState(false), [success, setSuccess] = useState(false);
  const [autoApplying, setAutoApplying] = useState(false);

  const member = team.members.find(m => m.id === memberId);
  const giftType = GIFT_TYPES.find(g => g.id === giftId);
  const isInvalid = giftId === "퇴사" && yrs === 1;
  const needsProof = giftId && !NO_PROOF.has(giftId);

  // 경조사 기준일이 이번 달 이상이면 신청 불가
  const thisMonthNum = new Date().getMonth() + 1;
  const thisYearNum = new Date().getFullYear();
  const lastMonthNum = thisMonthNum === 1 ? 12 : thisMonthNum - 1;
  const eventDateInvalid = (function() {
    if (!eventDate) return false;
    const d = new Date(eventDate);
    if (isNaN(d.getTime())) return false;
    const ey = d.getFullYear(), em = d.getMonth() + 1;
    return ey > thisYearNum || (ey === thisYearNum && em >= thisMonthNum);
  })();

  const canSubmit = memberId && giftId && amt && !isInvalid && !eventDateInvalid && (!needsProof || proofFile || proofLater);

  useEffect(() => { if (member) setYrs(calcYrs(member.입사일)); }, [memberId]);
  useEffect(() => { if (giftId && !overrideAmt) setAmt(getAmt(giftId, yrs)); }, [giftId, yrs, overrideAmt]);

  // 지난달 입사기념일/생일 대상자
  const prevMonth = new Date(); prevMonth.setMonth(prevMonth.getMonth() - 1);
  const prevM = prevMonth.getMonth() + 1;
  const prevY = prevMonth.getFullYear();

  const autoTargets = [];
  team.members.forEach(function(m) {
    const curMonth = nowYM();
    if (m.입사일) {
      const hd = new Date(m.입사일);
      if (hd.getMonth() + 1 === prevM) {
        // 이미 신청됐는지 확인
        const already = apps.find(a => a.teamId === team.id && a.month === curMonth && a.수령인 === m.이름 && a.경조사 === "입사기념일");
        if (!already) autoTargets.push({ member: m, type: "입사기념일", date: prevY + "-" + String(prevM).padStart(2,"0") + "-" + String(hd.getDate()).padStart(2,"0") });
      }
    }
    if (m.생일) {
      const bd = new Date(m.생일);
      if (bd.getMonth() + 1 === prevM) {
        const already = apps.find(a => a.teamId === team.id && a.month === curMonth && a.수령인 === m.이름 && a.경조사 === "생일");
        if (!already) autoTargets.push({ member: m, type: "생일", date: prevY + "-" + String(prevM).padStart(2,"0") + "-" + String(bd.getDate()).padStart(2,"0") });
      }
    }
  });

  async function applyAuto(target) {
    setAutoApplying(true);
    const m = target.member;
    const yrsVal = calcYrs(m.입사일);
    const amtVal = getAmt(target.type, yrsVal);
    const na = {
      id: uid(), month: nowYM(), teamId: team.id,
      팀명: team.팀명, 본부명: team.본부명, 부문명: team.부문명,
      수령인: m.이름, 계좌명: m.계좌명 || m.이름,
      은행명: m.은행명 || "", 계좌번호: m.계좌번호 || "",
      경조사: target.type, 경조사발생일: target.date,
      연차: yrsVal, 금액: amtVal, 메모: "",
      증빙파일: null, 증빙나중제출: false,
      createdAt: new Date().toISOString(),
    };
    await updApps([...apps, na]);
    setAutoApplying(false);
  }

  async function applyAllAuto() {
    setAutoApplying(true);
    const newApps = [...apps];
    autoTargets.forEach(function(target) {
      const m = target.member;
      const yrsVal = calcYrs(m.입사일);
      const amtVal = getAmt(target.type, yrsVal);
      newApps.push({
        id: uid(), month: nowYM(), teamId: team.id,
        팀명: team.팀명, 본부명: team.본부명, 부문명: team.부문명,
        수령인: m.이름, 계좌명: m.계좌명 || m.이름,
        은행명: m.은행명 || "", 계좌번호: m.계좌번호 || "",
        경조사: target.type, 경조사발생일: target.date,
        연차: yrsVal, 금액: getAmt(target.type, yrsVal), 메모: "",
        증빙파일: null, 증빙나중제출: false,
        createdAt: new Date().toISOString(),
      });
    });
    await updApps(newApps);
    setAutoApplying(false);
  }

  function handleFile(e) { const f = e.target.files[0]; if (!f) return; if (f.size > 4 * 1024 * 1024) { alert("4MB 이하만 가능합니다."); return; } const r = new FileReader(); r.onload = ev => setProofFile({ name: f.name, type: f.type, size: f.size, data: ev.target.result }); r.readAsDataURL(f); }
  function reset() { setMemberId(""); setGiftId(""); setEventDate(""); setAmt(0); setMemo(""); setOverrideAmt(false); setProofFile(null); setProofLater(false); }
  async function submit() {
    if (!canSubmit) return;
    const na = { id: uid(), month: nowYM(), teamId: team.id, 팀명: team.팀명, 본부명: team.본부명, 부문명: team.부문명, 수령인: member.이름, 계좌명: member.계좌명 || member.이름, 은행명: member.은행명 || "", 계좌번호: member.계좌번호 || "", 경조사: giftType ? giftType.label : giftId, 경조사발생일: eventDate, 연차: yrs, 금액: amt, 메모: memo, 증빙파일: proofFile ? { name: proofFile.name, type: proofFile.type, size: proofFile.size, data: proofFile.data } : null, 증빙나중제출: proofLater, createdAt: new Date().toISOString() };
    await updApps([...apps, na]);
    setSuccess(true); setTimeout(() => { reset(); setSuccess(false); }, 2000);
  }

  if (locked) {
    return (
      <div>
        <h2 className="font-bold text-gray-800 mb-4">경조금 신청</h2>
        <div className="bg-red-50 border border-red-200 rounded-2xl p-8 text-center">
          <div className="text-4xl mb-3">🔒</div>
          <div className="font-bold text-red-700 text-base mb-1">신청 마감</div>
          <div className="text-sm text-red-500">매월 <span className="font-bold">{deadlineDay}일</span> 이후에는 신청이 마감됩니다.</div>
          <div className="text-xs text-gray-400 mt-2">신청이 필요한 경우 관리자에게 문의하세요.</div>
        </div>
      </div>
    );
  }

  return (
    <div>
      <h2 className="font-bold text-gray-800 mb-4">경조금 신청</h2>

      {/* 지난달 자동신청 대상 */}
      {autoTargets.length > 0 && (
        <div className="bg-green-50 border border-green-200 rounded-2xl p-4 mb-4">
          <div className="flex justify-between items-center mb-2">
            <div className="text-xs font-bold text-green-700">🎂 지난달({prevM}월) 자동신청 대상</div>
            <button onClick={applyAllAuto} disabled={autoApplying}
              className="text-xs bg-green-500 hover:bg-green-600 text-white px-3 py-1.5 rounded-lg font-bold">
              전체 자동신청
            </button>
          </div>
          <div className="space-y-2">
            {autoTargets.map(function(target, i) {
              const yrsVal = calcYrs(target.member.입사일);
              const amtVal = getAmt(target.type, yrsVal);
              return (
                <div key={i} className="flex items-center justify-between bg-white rounded-xl px-3 py-2 border border-green-100">
                  <div>
                    <span className="text-sm font-semibold text-gray-800">{target.member.이름}</span>
                    <span className="text-xs text-green-600 ml-2">{target.type}</span>
                    <span className="text-xs text-gray-400 ml-1">({fmtED(target.date)})</span>
                    <span className="text-xs text-blue-600 font-semibold ml-2">{fmtMoney(amtVal)}</span>
                  </div>
                  <button onClick={() => applyAuto(target)} disabled={autoApplying}
                    className="text-xs bg-green-100 hover:bg-green-200 text-green-700 px-2.5 py-1 rounded-lg font-semibold">
                    신청
                  </button>
                </div>
              );
            })}
          </div>
        </div>
      )}

      <div className="bg-white rounded-2xl p-5 shadow-sm space-y-4">
        <div>
          <label className="block text-xs font-semibold text-gray-500 mb-1.5">팀원 선택 *</label>
          <select value={memberId} onChange={e => { setMemberId(e.target.value); setGiftId(""); setAmt(0); setOverrideAmt(false); setProofFile(null); setProofLater(false); }} className="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm focus:outline-none">
            <option value="">팀원을 선택하세요</option>
            {team.members.map(m => <option key={m.id} value={m.id}>{m.이름}</option>)}
          </select>
        </div>
        {member && (
          <div className="bg-blue-50 rounded-xl p-3 text-xs text-blue-700 space-y-0.5">
            <div className="font-bold">{member.이름}{member.계좌명 && member.계좌명 !== member.이름 ? " (예금주: " + member.계좌명 + ")" : ""}</div>
            <div>{member.은행명 || "은행 미등록"} · {member.계좌번호 || "계좌 미등록"}</div>
            <div>입사일: {member.입사일 || "미등록"} · <span className="font-semibold">{yrsLabel(yrs)}</span></div>
          </div>
        )}
        <div>
          <label className="block text-xs font-semibold text-gray-500 mb-1.5">연차 <span className="font-normal text-gray-400">(자동계산, 수정 가능)</span></label>
          <div className="flex items-center gap-3">
            <input type="number" min="1" max="5" value={yrs} onChange={e => { setYrs(Math.max(1, Math.min(5, Number(e.target.value) || 1))); setOverrideAmt(false); }} className="w-20 border border-gray-200 rounded-xl px-3 py-2.5 text-sm text-center focus:outline-none" />
            <span className="text-sm font-semibold text-blue-600">{yrsLabel(yrs)}</span>
          </div>
        </div>
        <div>
          <label className="block text-xs font-semibold text-gray-500 mb-1.5">경조사 종류 *</label>
          <select value={giftId} onChange={e => { setGiftId(e.target.value); setOverrideAmt(false); setProofFile(null); setProofLater(false); }} className="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm focus:outline-none">
            <option value="">경조사를 선택하세요</option>
            {GIFT_TYPES.map(function(g) { const a = g.amounts[Math.min(yrs, 5) - 1]; return <option key={g.id} value={g.id}>{g.label} — {a != null ? a.toLocaleString() + "원" : "해당없음"}</option>; })}
          </select>
          {isInvalid && <p className="text-red-500 text-xs mt-1.5">⚠️ 퇴사 경조금은 2년 이상 근속자에게만 지급됩니다.</p>}
        </div>
        <div>
          <label className="block text-xs font-semibold text-gray-500 mb-1.5">경조사 기준일</label>
          <DatePicker value={eventDate} onChange={setEventDate} startYear={cy - 3} endYear={cy} />
          {eventDateInvalid && (
            <div className="mt-2 bg-red-50 border border-red-200 rounded-xl px-3 py-2.5 text-xs text-red-600 font-semibold">
              ⛔ 이번 신청은 지난 달({lastMonthNum}월) 기준의 경조금을 신청해주세요.
            </div>
          )}
        </div>
        <div>
          <div className="flex justify-between items-center mb-1.5">
            <label className="text-xs font-semibold text-gray-500">지급 금액</label>
            <label className="flex items-center gap-1.5 text-xs text-gray-400 cursor-pointer select-none">
              <input type="checkbox" checked={overrideAmt} onChange={e => { setOverrideAmt(e.target.checked); if (!e.target.checked) setAmt(getAmt(giftId, yrs)); }} className="w-3 h-3" />금액 직접 입력
            </label>
          </div>
          {overrideAmt ? <input type="number" value={amt} onChange={e => setAmt(Number(e.target.value))} className="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm focus:outline-none" /> : <div className="border border-blue-200 bg-blue-50 rounded-xl px-4 py-2.5 text-sm font-bold text-blue-700">{fmtMoney(amt)}</div>}
        </div>
        {needsProof && (
          <div className="border border-orange-200 bg-orange-50 rounded-xl p-4">
            <div className="text-xs font-bold text-orange-700 mb-1">📎 증빙서류 첨부</div>
            <div className="text-xs text-orange-500 mb-3">{giftType ? giftType.label : ""} 경조사는 증빙서류가 필요합니다.</div>
            {!proofFile ? (
              <label className="block w-full border-2 border-dashed border-orange-300 rounded-xl p-4 text-center cursor-pointer hover:bg-orange-100">
                <div className="text-2xl mb-1">📁</div><div className="text-xs font-semibold text-orange-700">파일 선택 (사진, PDF)</div><div className="text-xs text-orange-400 mt-0.5">최대 4MB</div>
                <input type="file" accept="image/*,.pdf" onChange={handleFile} className="hidden" />
              </label>
            ) : (
              <div className="flex items-center gap-2 bg-white border border-green-200 rounded-xl p-3">
                <span className="text-lg">✅</span>
                <div className="min-w-0 flex-1"><div className="text-xs font-semibold text-green-700 truncate">{proofFile.name}</div><div className="text-xs text-gray-400">{(proofFile.size / 1024).toFixed(0)}KB</div></div>
                <button onClick={() => setProofFile(null)} className="text-red-400 text-sm">✕</button>
              </div>
            )}
            <label className="flex items-center gap-2 text-xs text-gray-500 cursor-pointer mt-3 select-none">
              <input type="checkbox" checked={proofLater} onChange={e => { setProofLater(e.target.checked); if (e.target.checked) setProofFile(null); }} className="w-3.5 h-3.5" />증빙서류를 추후에 제출하겠습니다
            </label>
          </div>
        )}
        <div><label className="block text-xs font-semibold text-gray-500 mb-1.5">메모 (선택)</label><input value={memo} onChange={e => setMemo(e.target.value)} className="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm focus:outline-none" placeholder="비고사항" /></div>
        <button onClick={submit} disabled={!canSubmit}
          className={"w-full py-3 rounded-xl font-bold text-sm transition-all shadow-sm " + (success ? "bg-green-500 text-white" : !canSubmit ? "bg-gray-200 text-gray-400 cursor-not-allowed" : "bg-blue-500 text-white hover:bg-blue-600")}>
          {success ? "✅ 신청 완료!" : "신청하기"}
        </button>
        {needsProof && !proofFile && !proofLater && memberId && giftId && <p className="text-orange-500 text-xs text-center -mt-2">증빙서류를 첨부하거나 '추후 제출'을 선택해야 신청됩니다.</p>}
      </div>
    </div>
  );
}

// ── History Tab ──
function HistoryTab({ team, apps, updApps, locked }) {
  const month = nowYM();
  const myApps = apps.filter(a => a.teamId === team.id && a.month === month);
  const total = myApps.reduce((s, a) => s + (a.금액 || 0), 0);
  const [confirmDel, setConfirmDel] = useState(null);
  async function del(id) { await updApps(apps.filter(a => a.id !== id)); setConfirmDel(null); }
  return (
    <div>
      <h2 className="font-bold text-gray-800 mb-4">이번달 신청내역 <span className="text-gray-400 font-normal text-sm">({month})</span></h2>
      {myApps.length === 0 ? (
        <div className="text-center text-gray-400 text-sm py-14 bg-white rounded-2xl shadow-sm"><div className="text-4xl mb-2">📋</div>이번달 신청 내역이 없습니다.</div>
      ) : (
        <div className="space-y-2">
          {myApps.map(a => (
            <div key={a.id} className="bg-white rounded-2xl p-4 shadow-sm flex justify-between items-start">
              <div>
                <div className="font-bold text-sm text-gray-800">{a.수령인}</div>
                <div className="text-xs text-gray-500 mt-0.5">{a.경조사}{a.경조사발생일 ? " · " + fmtED(a.경조사발생일) : ""} · {yrsExcel(a.연차 || 1)}</div>
                <div className="text-sm font-bold text-blue-600 mt-1">{fmtMoney(a.금액)}</div>
                <div className="text-xs text-gray-400 mt-0.5">{a.은행명} {a.계좌번호}</div>
                {a.증빙파일 && <div className="text-xs text-green-600 mt-0.5">📎 {a.증빙파일.name}</div>}
                {a.증빙나중제출 && !a.증빙파일 && <div className="text-xs text-orange-500 mt-0.5">⏳ 증빙 추후제출</div>}
              </div>
              {!locked && <button onClick={() => setConfirmDel(a.id)} className="text-xs text-red-400 bg-red-50 px-2.5 py-1.5 rounded-lg hover:bg-red-100 font-semibold shrink-0">삭제</button>}
            </div>
          ))}
          <div className="bg-blue-50 rounded-2xl p-4 flex justify-between items-center">
            <span className="text-sm text-gray-600 font-bold">합계 ({myApps.length}건)</span>
            <span className="font-bold text-blue-700 text-lg">{fmtMoney(total)}</span>
          </div>
        </div>
      )}
      {confirmDel && <ConfirmDlg msg="이 신청 내역을 삭제할까요?" onYes={() => del(confirmDel)} onNo={() => setConfirmDel(null)} />}
    </div>
  );
}

// ── Admin Screen ──
function AdminScreen({ adminPw, updAdminPw, deadlineDay, updDeadlineDay, teams, updTeams, apps, updApps, onLogout }) {
  const [month, setMonth] = useState(nowYM()), [confirmDel, setConfirmDel] = useState(null), [tab, setTab] = useState("list");
  const [showPwModal, setShowPwModal] = useState(false);
  const filtered = apps.filter(a => a.month === month);
  const total = filtered.reduce((s, a) => s + (a.금액 || 0), 0);
  const pendingProof = filtered.filter(a => a.증빙나중제출 && !a.증빙파일).length;
  async function del(id) { await updApps(apps.filter(a => a.id !== id)); setConfirmDel(null); }

  function dlCompiled() {
    try {
      const rows = filtered.map(function(a) {
        return { "본부": a.본부명||"", "팀명": a.팀명||"", "경조사 기준일": fmtED(a.경조사발생일)||"", "신청자명": a.수령인||"", "근무연차": yrsExcel(a.연차||1), "경조사종류": a.경조사||"", "금융기관명": a.은행명||"", "입금계좌번호": String(a.계좌번호||""), "지급금액": a.금액||0, "증빙": a.증빙파일?"제출완료":a.증빙나중제출?"추후제출":"-", "메모": a.메모||"" };
      });
      const ws = XLSX.utils.json_to_sheet(rows); ws["!cols"] = [{wch:14},{wch:14},{wch:11},{wch:10},{wch:10},{wch:16},{wch:12},{wch:20},{wch:10},{wch:10},{wch:16}];
      const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "경조금취합"); XLSX.writeFile(wb, "경조금취합양식_" + month + ".xlsx");
    } catch(e) { console.error(e); }
  }

  function dlBank() {
    try {
      const merged = {};
      filtered.forEach(function(a) {
        const key = String(a.계좌번호 || "");
        if (!merged[key]) { merged[key] = { "입금은행코드": BANK_CODES[a.은행명]||"", "입금계좌번호": key, "이체금액": 0, "예상예금주": a.계좌명||a.수령인||"", "보내는분 통장표시내용": "", "받는분 통장표시내용": "사우회경조금", "CMS/모집인코드": "" }; }
        merged[key]["이체금액"] += (a.금액 || 0);
      });
      const rows = Object.values(merged);
      const ws = XLSX.utils.json_to_sheet(rows); ws["!cols"] = [{wch:12},{wch:20},{wch:10},{wch:12},{wch:18},{wch:18},{wch:12}];
      const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "은행이체"); XLSX.writeFile(wb, "은행이체양식_" + month + ".xlsx");
    } catch(e) { console.error(e); }
  }

  const teamSum = {}, typeSum = {};
  filtered.forEach(function(a) {
    const tk = a.팀명||"미등록"; if (!teamSum[tk]) teamSum[tk] = { cnt:0, tot:0 }; teamSum[tk].cnt++; teamSum[tk].tot += (a.금액||0);
    const gk = a.경조사||"기타"; if (!typeSum[gk]) typeSum[gk] = { cnt:0, tot:0 }; typeSum[gk].cnt++; typeSum[gk].tot += (a.금액||0);
  });

  return (
    <div className="min-h-screen bg-gray-50">
      <div className="bg-blue-500 text-white px-4 py-3 flex justify-between items-center" style={{ position: "sticky", top: 0, zIndex: 20 }}>
        <div><div className="font-bold text-sm">🔑 관리자 화면</div><div className="text-xs" style={{ opacity: .75 }}>전체 경조금 내역</div></div>
        <div className="flex gap-2">
          <button onClick={() => setShowPwModal(true)} className="text-xs bg-blue-600 hover:bg-blue-700 px-3 py-1.5 rounded-lg">🔑 비번변경</button>
          <button onClick={onLogout} className="text-xs bg-blue-600 hover:bg-blue-700 px-3 py-1.5 rounded-lg">로그아웃</button>
        </div>
      </div>
      <div className="p-4 max-w-3xl mx-auto space-y-4">
        <div className="flex flex-wrap gap-3 items-center justify-between">
          <div className="flex items-center gap-2">
            <label className="text-xs font-semibold text-gray-500">조회 월</label>
            <input type="month" value={month} onChange={e => setMonth(e.target.value)} className="border border-gray-200 rounded-xl px-3 py-2 text-sm focus:outline-none" />
          </div>
          <div className="flex gap-2">
            <button onClick={dlCompiled} className="bg-indigo-500 text-white px-3 py-2 rounded-xl text-xs font-bold hover:bg-indigo-600 shadow-sm">📋 경조금취합양식</button>
            <button onClick={dlBank} className="bg-green-500 text-white px-3 py-2 rounded-xl text-xs font-bold hover:bg-green-600 shadow-sm">🏦 은행이체양식</button>
          </div>
        </div>

        {/* 마감일 설정 */}
        <div className="bg-white rounded-2xl p-4 shadow-sm">
          <div className="flex items-center justify-between flex-wrap gap-3">
            <div>
              <div className="text-sm font-bold text-gray-700">📅 신청 마감일 설정</div>
              <div className="text-xs text-gray-400 mt-0.5">매월 설정한 날짜 이후 담당자 신청이 자동 잠금됩니다.</div>
            </div>
            <div className="flex items-center gap-2">
              <span className="text-xs text-gray-500">매월</span>
              <select value={deadlineDay || ""} onChange={e => updDeadlineDay(e.target.value || null)}
                className="border border-gray-200 rounded-xl px-3 py-2 text-sm focus:outline-none">
                <option value="">마감일 없음</option>
                {Array.from({ length: 31 }, (_, i) => i + 1).map(d => <option key={d} value={d}>{d}일</option>)}
              </select>
              <span className="text-xs text-gray-500">이후 잠금</span>
              {deadlineDay && (
                isDeadlinePassed(deadlineDay)
                  ? <span className="text-xs bg-red-100 text-red-600 px-2 py-1 rounded-lg font-bold">🔒 현재 잠금 중</span>
                  : <span className="text-xs bg-green-100 text-green-600 px-2 py-1 rounded-lg font-bold">✅ 신청 가능</span>
              )}
            </div>
          </div>
        </div>

        <div className="grid grid-cols-3 gap-3">
          {[
            [String(filtered.length)+"건","신청 건수","text-blue-600"],
            [fmtMoney(total),"총 이체금액","text-blue-600"],
            [pendingProof>0?pendingProof+"건":"없음","증빙 추후제출",pendingProof>0?"text-orange-500":"text-gray-400"],
          ].map(function(item) {
            return (<div key={item[1]} className="bg-white rounded-2xl p-4 shadow-sm text-center"><div className={"text-xl font-bold "+item[2]}>{item[0]}</div><div className="text-xs text-gray-400 mt-0.5">{item[1]}</div></div>);
          })}
        </div>

        <div className="flex gap-1 bg-white rounded-2xl p-1 shadow-sm overflow-x-auto">
          {[["list","📋 내역"],["team","팀별"],["type","경조사별"],["teams","🏢 팀 관리"]].map(([t, l]) => (
            <button key={t} onClick={() => setTab(t)} className={"flex-1 py-2 rounded-xl text-xs font-bold whitespace-nowrap " + (tab===t?"bg-blue-500 text-white shadow-sm":"text-gray-400 hover:bg-gray-50")}>{l}</button>
          ))}
        </div>

        {tab === "list" && (
          <div className="bg-white rounded-2xl shadow-sm overflow-hidden">
            {filtered.length === 0 ? (
              <div className="text-center text-gray-400 text-sm py-12"><div className="text-4xl mb-2">📭</div>해당 월에 신청 내역이 없습니다.</div>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full" style={{ fontSize: "11px" }}>
                  <thead><tr className="bg-gray-50 border-b">{["팀명","수령인","기준일","경조사","연차","금액","은행","계좌번호","증빙",""].map(h => <th key={h} className="px-2.5 py-2.5 text-left font-semibold text-gray-500 whitespace-nowrap">{h}</th>)}</tr></thead>
                  <tbody>
                    {filtered.map(a => (
                      <tr key={a.id} className="border-b hover:bg-gray-50">
                        <td className="px-2.5 py-2.5 text-gray-600 whitespace-nowrap">{a.팀명}</td>
                        <td className="px-2.5 py-2.5 font-bold text-gray-800 whitespace-nowrap">{a.수령인}</td>
                        <td className="px-2.5 py-2.5 text-gray-500 whitespace-nowrap">{fmtED(a.경조사발생일)}</td>
                        <td className="px-2.5 py-2.5 text-gray-600 whitespace-nowrap">{a.경조사}</td>
                        <td className="px-2.5 py-2.5 text-gray-500 whitespace-nowrap">{yrsExcel(a.연차||1)}</td>
                        <td className="px-2.5 py-2.5 font-bold text-blue-600 text-right whitespace-nowrap">{fmtMoney(a.금액)}</td>
                        <td className="px-2.5 py-2.5 text-gray-500 whitespace-nowrap">{a.은행명}</td>
                        <td className="px-2.5 py-2.5 text-gray-500">{a.계좌번호}</td>
                        <td className="px-2.5 py-2.5 whitespace-nowrap">{a.증빙파일?<span className="text-green-600 font-semibold">✅</span>:a.증빙나중제출?<span className="text-orange-500">⏳</span>:<span className="text-gray-300">-</span>}</td>
                        <td className="px-2.5 py-2.5"><button onClick={() => setConfirmDel(a.id)} className="text-red-300 hover:text-red-500 text-base leading-none">✕</button></td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot><tr className="bg-blue-50"><td colSpan={5} className="px-2.5 py-2.5 text-xs font-bold text-gray-600">합계 {filtered.length}건</td><td className="px-2.5 py-2.5 text-xs font-bold text-blue-700 text-right">{fmtMoney(total)}</td><td colSpan={4}></td></tr></tfoot>
                </table>
              </div>
            )}
          </div>
        )}

        {tab === "team" && (
          <div className="bg-white rounded-2xl shadow-sm p-4"><div className="text-xs font-bold text-gray-500 mb-3">팀별 집계</div>
            {Object.keys(teamSum).length === 0 ? <div className="text-xs text-gray-400 py-4 text-center">내역 없음</div> : (
              <div className="space-y-2">{Object.entries(teamSum).sort((a,b) => b[1].tot-a[1].tot).map(([k,v]) => (<div key={k} className="flex justify-between items-center py-2 border-b last:border-0"><span className="text-sm font-semibold text-gray-700">{k}</span><span className="text-xs text-gray-400">{v.cnt}건 <span className="font-bold text-blue-600 text-sm ml-1">{fmtMoney(v.tot)}</span></span></div>))}</div>
            )}
          </div>
        )}
        {tab === "type" && (
          <div className="bg-white rounded-2xl shadow-sm p-4"><div className="text-xs font-bold text-gray-500 mb-3">경조사별 집계</div>
            {Object.keys(typeSum).length === 0 ? <div className="text-xs text-gray-400 py-4 text-center">내역 없음</div> : (
              <div className="space-y-2">{Object.entries(typeSum).sort((a,b) => b[1].tot-a[1].tot).map(([k,v]) => (<div key={k} className="flex justify-between items-center py-2 border-b last:border-0"><span className="text-sm font-semibold text-gray-700">{k}</span><span className="text-xs text-gray-400">{v.cnt}건 <span className="font-bold text-blue-600 text-sm ml-1">{fmtMoney(v.tot)}</span></span></div>))}</div>
            )}
          </div>
        )}
        {tab === "teams" && <TeamsManageTab teams={teams} updTeams={updTeams} />}
      </div>
      {confirmDel && <ConfirmDlg msg="이 신청 내역을 삭제할까요?" onYes={() => del(confirmDel)} onNo={() => setConfirmDel(null)} />}
      {showPwModal && <ChangePasswordModal currentPw={adminPw} onSave={async pw => { await updAdminPw(pw); setShowPwModal(false); }} onClose={() => setShowPwModal(false)} />}
    </div>
  );
}

// ── Teams Manage Tab ──
function TeamsManageTab({ teams, updTeams }) {
  const [showForm, setShowForm] = useState(false), [form, setForm] = useState({ 본부명:"", 부문명:"", 팀명:"" }), [search, setSearch] = useState(""), [confirmDel, setConfirmDel] = useState(null);
  async function addTeam() {
    if (!form.본부명.trim() || !form.팀명.trim()) return;
    const ex = teams.find(t => t.본부명===form.본부명.trim()&&t.부문명===form.부문명.trim()&&t.팀명===form.팀명.trim());
    if (ex) { alert("이미 존재하는 팀입니다."); return; }
    await updTeams([...teams, { id:uid(), 본부명:form.본부명.trim(), 부문명:form.부문명.trim(), 팀명:form.팀명.trim(), members:[] }]);
    setForm({ 본부명:"", 부문명:"", 팀명:"" }); setShowForm(false);
  }
  async function delTeam(id) { await updTeams(teams.filter(t => t.id !== id)); setConfirmDel(null); }
  const filtered = teams.filter(t => { const q=search.toLowerCase().trim(); if(!q)return true; return t.팀명.includes(q)||t.본부명.includes(q)||(t.부문명&&t.부문명.includes(q)); });
  return (
    <div>
      <div className="flex justify-between items-center mb-4">
        <h2 className="font-bold text-gray-800">팀 목록 관리 <span className="text-gray-400 font-normal">({teams.length}팀)</span></h2>
        <button onClick={() => setShowForm(true)} className="bg-blue-500 text-white text-xs px-3 py-2 rounded-xl font-bold shadow-sm">+ 팀 추가</button>
      </div>
      <input value={search} onChange={e => setSearch(e.target.value)} className="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm focus:outline-none mb-3" placeholder="🔍  팀명, 본부명으로 검색..." />
      {filtered.length === 0 ? <div className="text-center text-gray-400 text-sm py-10 bg-white rounded-2xl">등록된 팀이 없습니다.</div> : (
        <div className="space-y-2">{filtered.map(t => (<div key={t.id} className="bg-white rounded-2xl p-4 shadow-sm flex justify-between items-center"><div><div className="font-semibold text-sm text-gray-800">{t.팀명}</div><div className="text-xs text-gray-400 mt-0.5">{[t.본부명,t.부문명].filter(Boolean).join(" · ")}</div><div className="text-xs text-blue-400 mt-0.5">{t.members.length}명 등록</div></div><button onClick={() => setConfirmDel(t.id)} className="text-xs text-red-400 bg-red-50 hover:bg-red-100 px-2.5 py-1.5 rounded-lg font-semibold">삭제</button></div>))}</div>
      )}
      {showForm && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-3xl p-6 w-full max-w-sm shadow-2xl">
            <div className="flex justify-between items-center mb-5"><h3 className="font-bold text-gray-800">팀 추가</h3><button onClick={() => setShowForm(false)} className="text-gray-300 hover:text-gray-500 text-2xl leading-none">✕</button></div>
            {[["본부명 *","본부명","예: 교육사업본부"],["부문(실)명","부문명","없으면 생략"],["팀명 *","팀명","예: 기획팀"]].map(([lbl,k,ph]) => (
              <div key={k} className="mb-3"><label className="block text-xs font-semibold text-gray-500 mb-1.5">{lbl}</label><input value={form[k]} onChange={e => setForm(f => ({ ...f, [k]: e.target.value }))} className="w-full border border-gray-200 rounded-xl px-4 py-2.5 text-sm focus:outline-none" placeholder={ph} /></div>
            ))}
            <div className="flex gap-2 mt-4">
              <button onClick={() => setShowForm(false)} className="flex-1 py-2.5 bg-gray-100 text-gray-600 rounded-xl text-sm font-bold">취소</button>
              <button onClick={addTeam} disabled={!form.본부명.trim()||!form.팀명.trim()} className={"flex-1 py-2.5 rounded-xl text-sm font-bold "+(!form.본부명.trim()||!form.팀명.trim()?"bg-gray-200 text-gray-400":"bg-blue-500 text-white hover:bg-blue-600")}>추가</button>
            </div>
          </div>
        </div>
      )}
      {confirmDel && <ConfirmDlg msg="이 팀을 삭제할까요?" onYes={() => delTeam(confirmDel)} onNo={() => setConfirmDel(null)} />}
    </div>
  );
}