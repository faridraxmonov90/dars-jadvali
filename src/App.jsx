import React, { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { storageGet, storageSet } from "./storageApi";
import { Settings, GraduationCap, BookOpen, Pencil, Ruler, Calculator, School, Bell, CalendarDays, CalendarOff, ClipboardList, AlertTriangle, Upload, CheckCircle2, ChevronLeft, FileDown } from "lucide-react";

/* ---------- palette ---------- */
const C = {
  navy: "#1E2A45",
  navy2: "#2B3A5C",
  bg: "#F3F5F9",
  card: "#FFFFFF",
  line: "#E2E6EE",
  gold: "#C98A3E",
  goldDark: "#A66F2E",
  teal: "#2F8F82",
  red: "#C0503F",
  temp: "#FCEBD5",
  tempBorder: "#E3A75B",
  darkBlue: "#0B1F4D",
  textMain: "#1B2333",
  textSoft: "#6B7385",
};

const STORAGE_KEY = "dars-jadvali-db";

const uid = () => Math.random().toString(36).slice(2, 10);
const fmtDate = (iso) => {
  const d = new Date(iso);
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `${dd}.${mm}.${d.getFullYear()}`;
};

const defaultDb = () => ({
  adminPassword: "1234",
  teachers: [],
  subjects: [],
  days: [
    { id: uid(), name: "Dushanba" },
    { id: uid(), name: "Seshanba" },
    { id: uid(), name: "Chorshanba" },
    { id: uid(), name: "Payshanba" },
    { id: uid(), name: "Juma" },
    { id: uid(), name: "Shanba" },
  ],
  slots: [
    { id: uid(), type: "dars", name: "1-dars", start: "08:30", end: "09:15" },
    { id: uid(), type: "tanaffus", name: "Tanaffus", start: "09:15", end: "09:25" },
    { id: uid(), type: "dars", name: "2-dars", start: "09:25", end: "10:10" },
    { id: uid(), type: "tanaffus", name: "Tanaffus", start: "10:10", end: "10:20" },
    { id: uid(), type: "dars", name: "3-dars", start: "10:20", end: "11:05" },
    { id: uid(), type: "tanaffus", name: "Katta tanaffus", start: "11:05", end: "11:25" },
    { id: uid(), type: "dars", name: "4-dars", start: "11:25", end: "12:10" },
    { id: uid(), type: "dars", name: "5-dars", start: "12:20", end: "13:05" },
  ],
  classes: [],
  rooms: [],
  schedule: {}, // classId -> dayId -> slotId -> { split, entries:[{id,subjectId,teacherId,roomId,comment}], tempEdits:[{id,date,entries:[...]}] }
  absenceRequests: [], // {id, teacherId, teacherName, startDate, endDate, reason, createdAt, read}
  tasks: [], // {id, teacherId, title, description, createdAt, read}
  yearlyPlans: [], // {id, teacherId, classId, subjectId, importedAt, rows:[{id,order,label,topics[],darsTuri,sinfIshi,uygaVazifa,soat,izoh,chorak,isBSB,isCHSB,done,doneDate,doneTopics[]}]}
  bsbNotifications: [], // {id, teacherId, teacherName, classId, className, subjectId, subjectName, type, date, read}
  dailyTopicAlerts: [], // {id, date, items:[{teacherId,teacherName,classId,className,subjectId,subjectName}], createdAt, read}
});

const clone = (x) => JSON.parse(JSON.stringify(x));

/* ---------- small ui bits ---------- */
function Modal({ title, onClose, children, wide }) {
  return (
    <div
      style={{ background: "rgba(20,24,38,0.55)" }}
      className="fixed inset-0 z-50 flex items-center justify-center p-4"
      onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}
    >
      <div
        style={{ background: C.card, borderColor: C.line }}
        className={`w-full ${wide ? "max-w-2xl" : "max-w-md"} rounded-xl border shadow-xl max-h-[90vh] overflow-y-auto`}
      >
        <div style={{ borderColor: C.line }} className="flex items-center justify-between px-5 py-4 border-b sticky top-0 bg-white rounded-t-xl">
          <h3 style={{ color: C.navy }} className="font-semibold text-lg">{title}</h3>
          <button onClick={onClose} style={{ color: C.textSoft }} className="text-xl leading-none hover:text-black px-1">×</button>
        </div>
        <div className="p-5">{children}</div>
      </div>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <label className="block mb-3">
      <span style={{ color: C.textSoft }} className="block text-xs font-medium mb-1 uppercase tracking-wide">{label}</span>
      {children}
    </label>
  );
}

const inputCls = "w-full rounded-md border px-3 py-2 text-sm outline-none focus:ring-2";
const inputStyle = { borderColor: C.line, color: C.textMain, ["--tw-ring-color"]: C.gold };

function Btn({ children, onClick, kind = "primary", type = "button", small }) {
  const styles = {
    primary: { background: C.gold, color: "#fff" },
    ghost: { background: "transparent", color: C.navy, border: `1px solid ${C.line}` },
    danger: { background: "#fff", color: C.red, border: `1px solid ${C.red}` },
    dark: { background: C.navy, color: "#fff" },
  };
  return (
    <button
      type={type}
      onClick={onClick}
      style={styles[kind]}
      className={`rounded-md font-medium ${small ? "px-2.5 py-1 text-xs" : "px-4 py-2 text-sm"} hover:opacity-90 transition`}
    >
      {children}
    </button>
  );
}

function parseNameRows(raw) {
  const keyWords = ["nomi", "fan", "fan nomi", "name", "subject"];
  return raw.map((row) => {
    let name = "";
    for (const k of Object.keys(row)) {
      if (keyWords.includes(String(k).trim().toLowerCase())) { name = String(row[k] ?? "").trim(); break; }
    }
    if (!name) name = String(Object.values(row)[0] ?? "").trim();
    return { name, valid: !!name };
  });
}

function downloadNameTemplate(filename, sheetName, header, examples) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([[header], ...examples.map((e) => [e])]);
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
}

function NamedImportControls({ items, onImportMany, label, fileHeader, templateFile, templateSheet, templateExamples }) {
  const [rows, setRows] = useState(null);
  const [err, setErr] = useState("");

  const handleFile = (e) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    setErr("");
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });
        if (raw.length === 0) { setErr("Faylda ma’lumot topilmadi."); return; }
        setRows(parseNameRows(raw));
      } catch {
        setErr("Faylni o‘qib bo‘lmadi. .xlsx yoki .csv formatdagi faylni tanlang.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const confirm = () => {
    const existing = new Set(items.map((it) => it.name.toLowerCase()));
    const toAdd = [];
    rows.forEach((r) => {
      if (!r.valid) return;
      if (existing.has(r.name.toLowerCase())) return;
      existing.add(r.name.toLowerCase());
      toAdd.push(r.name);
    });
    onImportMany(toAdd);
    setRows(null);
  };

  return (
    <>
      <div className="flex flex-wrap gap-2 mb-4 justify-end">
        <Btn kind="ghost" onClick={() => downloadNameTemplate(templateFile, templateSheet, fileHeader, templateExamples)}>⬇ Namuna faylni yuklab olish</Btn>
        <label style={{ background: C.teal, color: "#fff" }} className="rounded-md font-medium px-4 py-2 text-sm hover:opacity-90 transition cursor-pointer">
          📥 Excel orqali import qilish
          <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleFile} />
        </label>
      </div>
      {err && <p style={{ color: C.red }} className="text-xs mb-3">{err}</p>}
      {rows && (
        <Modal title={`Excel’dan import — ${label}`} onClose={() => setRows(null)}>
          <p style={{ color: C.textSoft }} className="text-sm mb-3">Jami {rows.length} qator topildi. Ustun: <b>{fileHeader}</b> (bitta ustunga bittadan nom).</p>
          <div className="max-h-72 overflow-y-auto border rounded-md" style={{ borderColor: C.line }}>
            <table className="w-full text-xs">
              <thead>
                <tr style={{ background: C.bg }}>
                  <th className="text-left px-2 py-1.5">Nomi</th>
                  <th className="text-left px-2 py-1.5">Holat</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((r, i) => {
                  const dup = items.some((it) => it.name.toLowerCase() === r.name.toLowerCase());
                  return (
                    <tr key={i} style={{ borderTop: `1px solid ${C.line}` }}>
                      <td className="px-2 py-1.5">{r.name || "—"}</td>
                      <td className="px-2 py-1.5">
                        {!r.valid ? <span style={{ color: C.red }}>✕ Bo‘sh qator</span> : dup ? <span style={{ color: C.goldDark }}>⚠ Mavjud, o‘tkazib yuboriladi</span> : <span style={{ color: C.teal }}>✓ Qo‘shiladi</span>}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div className="flex justify-end gap-2 mt-4">
            <Btn kind="ghost" onClick={() => setRows(null)}>Bekor qilish</Btn>
            <Btn onClick={confirm}>Import qilish</Btn>
          </div>
        </Modal>
      )}
    </>
  );
}

/* ---------- generic named-list editor (subjects / classes / rooms / days) ---------- */
function NamedListEditor({ items, onAdd, onRename, onDelete, onMove, label, placeholder }) {
  const [val, setVal] = useState("");
  const [editing, setEditing] = useState(null);
  const [editVal, setEditVal] = useState("");

  return (
    <div>
      <div className="flex gap-2 mb-4">
        <input
          value={val}
          onChange={(e) => setVal(e.target.value)}
          placeholder={placeholder}
          className={inputCls}
          style={inputStyle}
          onKeyDown={(e) => { if (e.key === "Enter" && val.trim()) { onAdd(val.trim()); setVal(""); } }}
        />
        <Btn onClick={() => { if (val.trim()) { onAdd(val.trim()); setVal(""); } }}>Qo‘shish</Btn>
      </div>
      {items.length === 0 && <p style={{ color: C.textSoft }} className="text-sm">{label} hali kiritilmagan.</p>}
      <ul className="space-y-2">
        {items.map((it, i) => (
          <li key={it.id} style={{ borderColor: C.line }} className="flex items-center gap-2 border rounded-md px-3 py-2">
            {onMove && (
              <div className="flex flex-col mr-1">
                <button disabled={i === 0} onClick={() => onMove(it.id, -1)} style={{ color: i === 0 ? "#CCC" : C.textSoft }} className="text-xs leading-none">▲</button>
                <button disabled={i === items.length - 1} onClick={() => onMove(it.id, 1)} style={{ color: i === items.length - 1 ? "#CCC" : C.textSoft }} className="text-xs leading-none">▼</button>
              </div>
            )}
            {editing === it.id ? (
              <>
                <input value={editVal} onChange={(e) => setEditVal(e.target.value)} className={inputCls + " flex-1"} style={inputStyle} />
                <Btn small onClick={() => { onRename(it.id, editVal.trim() || it.name); setEditing(null); }}>Saqlash</Btn>
                <Btn small kind="ghost" onClick={() => setEditing(null)}>Bekor</Btn>
              </>
            ) : (
              <>
                <span style={{ color: C.textMain }} className="flex-1 text-sm">{it.name}</span>
                <Btn small kind="ghost" onClick={() => { setEditing(it.id); setEditVal(it.name); }}>✎ Tahrirlash</Btn>
                <Btn small kind="danger" onClick={() => onDelete(it.id)}>O‘chirish</Btn>
              </>
            )}
          </li>
        ))}
      </ul>
    </div>
  );
}

/* ---------- teachers editor ---------- */
const HEADER_MAP = {
  ism: "firstName", "first name": "firstName", firstname: "firstName", name: "firstName",
  familiya: "lastName", "last name": "lastName", lastname: "lastName", surname: "lastName",
  login: "username", username: "username",
  parol: "password", password: "password",
  fan: "subjects", fanlar: "subjects", subjects: "subjects", subject: "subjects",
};

function parseTeacherRows(raw, subjects) {
  return raw.map((row) => {
    const norm = {};
    Object.keys(row).forEach((k) => {
      const key = HEADER_MAP[String(k).trim().toLowerCase()];
      if (key) norm[key] = String(row[k] ?? "").trim();
    });
    const names = (norm.subjects || "").split(/[,;/]/).map((s) => s.trim()).filter(Boolean);
    const subjectIds = names.map((name) => subjects.find((s) => s.name.toLowerCase() === name.toLowerCase())?.id).filter(Boolean);
    const unmatched = names.filter((name) => !subjects.some((s) => s.name.toLowerCase() === name.toLowerCase()));
    const valid = !!(norm.firstName && norm.username && norm.password);
    return {
      firstName: norm.firstName || "",
      lastName: norm.lastName || "",
      username: norm.username || "",
      password: norm.password || "",
      subjectIds,
      subjectText: norm.subjects || "",
      unmatched,
      valid,
    };
  });
}

function downloadTeacherTemplate() {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ["Ism", "Familiya", "Login", "Parol", "Fanlar"],
    ["Aziza", "Karimova", "aziza.k", "parol123", "Matematika, Fizika"],
  ]);
  XLSX.utils.book_append_sheet(wb, ws, "O‘qituvchilar");
  XLSX.writeFile(wb, "oqituvchilar-namuna.xlsx");
}

function TeachersEditor({ teachers, subjects, onSave, onDelete, onImportMany }) {
  const [modal, setModal] = useState(null); // {id?, firstName, lastName, username, password, subjectIds}
  const [showPw, setShowPw] = useState({});
  const [importRows, setImportRows] = useState(null); // parsed rows pending confirmation
  const [importErr, setImportErr] = useState("");

  const empty = { id: null, firstName: "", lastName: "", username: "", password: "", subjectIds: [] };

  const handleFile = (e) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    setImportErr("");
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });
        if (raw.length === 0) { setImportErr("Faylda ma’lumot topilmadi."); return; }
        const rows = parseTeacherRows(raw, subjects);
        setImportRows(rows);
      } catch {
        setImportErr("Faylni o‘qib bo‘lmadi. .xlsx yoki .csv formatdagi faylni tanlang.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const confirmImport = () => {
    const existingUsernames = new Set(teachers.map((t) => t.username.toLowerCase()));
    const toAdd = [];
    importRows.forEach((r) => {
      if (!r.valid) return;
      if (existingUsernames.has(r.username.toLowerCase())) return;
      existingUsernames.add(r.username.toLowerCase());
      toAdd.push({ id: uid(), firstName: r.firstName, lastName: r.lastName, username: r.username, password: r.password, subjectIds: r.subjectIds });
    });
    onImportMany(toAdd);
    setImportRows(null);
  };

  return (
    <div>
      <div className="flex flex-wrap justify-end gap-2 mb-4">
        <Btn kind="ghost" onClick={downloadTeacherTemplate}>⬇ Namuna faylni yuklab olish</Btn>
        <label style={{ background: C.teal, color: "#fff" }} className="rounded-md font-medium px-4 py-2 text-sm hover:opacity-90 transition cursor-pointer">
          📥 Excel orqali import qilish
          <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleFile} />
        </label>
        <Btn onClick={() => setModal({ ...empty })}>+ O‘qituvchi qo‘shish</Btn>
      </div>
      {importErr && <p style={{ color: C.red }} className="text-xs mb-3">{importErr}</p>}
      {teachers.length === 0 && <p style={{ color: C.textSoft }} className="text-sm">O‘qituvchilar hali kiritilmagan.</p>}
      <div className="space-y-2">
        {teachers.map((t) => (
          <div key={t.id} style={{ borderColor: C.line }} className="flex items-center gap-3 border rounded-md px-3 py-2 text-sm">
            <div className="flex-1">
              <div style={{ color: C.textMain }} className="font-medium">{t.firstName} {t.lastName}</div>
              <div style={{ color: C.textSoft }} className="text-xs">
                Login: {t.username} · Parol: {showPw[t.id] ? t.password : "••••••"}{" "}
                <button className="underline" onClick={() => setShowPw((s) => ({ ...s, [t.id]: !s[t.id] }))}>{showPw[t.id] ? "yashirish" : "ko‘rsatish"}</button>
              </div>
              {t.subjectIds?.length > 0 && (
                <div style={{ color: C.teal }} className="text-xs mt-0.5">{t.subjectIds.map((sid) => subjects.find((s) => s.id === sid)?.name).filter(Boolean).join(", ")}</div>
              )}
            </div>
            <Btn small kind="ghost" onClick={() => setModal({ ...t })}>✎ Tahrirlash</Btn>
            <Btn small kind="danger" onClick={() => onDelete(t.id)}>O‘chirish</Btn>
          </div>
        ))}
      </div>

      {modal && (
        <Modal title={modal.id ? "O‘qituvchini tahrirlash" : "Yangi o‘qituvchi"} onClose={() => setModal(null)}>
          <Field label="Ism"><input className={inputCls} style={inputStyle} value={modal.firstName} onChange={(e) => setModal({ ...modal, firstName: e.target.value })} /></Field>
          <Field label="Familiya"><input className={inputCls} style={inputStyle} value={modal.lastName} onChange={(e) => setModal({ ...modal, lastName: e.target.value })} /></Field>
          <Field label="Login"><input className={inputCls} style={inputStyle} value={modal.username} onChange={(e) => setModal({ ...modal, username: e.target.value })} /></Field>
          <Field label="Parol"><input className={inputCls} style={inputStyle} value={modal.password} onChange={(e) => setModal({ ...modal, password: e.target.value })} /></Field>
          <Field label="Fanlar (ixtiyoriy)">
            <div className="flex flex-wrap gap-2">
              {subjects.map((s) => {
                const active = modal.subjectIds?.includes(s.id);
                return (
                  <button
                    key={s.id}
                    type="button"
                    onClick={() => {
                      const ids = modal.subjectIds || [];
                      setModal({ ...modal, subjectIds: active ? ids.filter((x) => x !== s.id) : [...ids, s.id] });
                    }}
                    style={active ? { background: C.gold, color: "#fff", borderColor: C.gold } : { borderColor: C.line, color: C.textMain }}
                    className="text-xs px-2.5 py-1 rounded-full border"
                  >
                    {s.name}
                  </button>
                );
              })}
              {subjects.length === 0 && <span style={{ color: C.textSoft }} className="text-xs">Avval fan qo‘shing.</span>}
            </div>
          </Field>
          <div className="flex justify-end gap-2 mt-4">
            <Btn kind="ghost" onClick={() => setModal(null)}>Bekor qilish</Btn>
            <Btn onClick={() => { if (!modal.firstName || !modal.username || !modal.password) return; onSave({ ...modal, id: modal.id || uid() }); setModal(null); }}>Saqlash</Btn>
          </div>
        </Modal>
      )}

      {importRows && (
        <Modal title="Excel’dan import — tekshirish" onClose={() => setImportRows(null)} wide>
          <p style={{ color: C.textSoft }} className="text-sm mb-3">
            Jami {importRows.length} qator topildi. Ustunlar: <b>Ism, Familiya, Login, Parol, Fanlar</b> (Fanlar ixtiyoriy, vergul bilan ajratiladi).
          </p>
          <div className="max-h-72 overflow-y-auto border rounded-md" style={{ borderColor: C.line }}>
            <table className="w-full text-xs">
              <thead>
                <tr style={{ background: C.bg }}>
                  <th className="text-left px-2 py-1.5">Ism Familiya</th>
                  <th className="text-left px-2 py-1.5">Login</th>
                  <th className="text-left px-2 py-1.5">Fanlar</th>
                  <th className="text-left px-2 py-1.5">Holat</th>
                </tr>
              </thead>
              <tbody>
                {importRows.map((r, i) => {
                  const dup = teachers.some((t) => t.username.toLowerCase() === r.username.toLowerCase());
                  return (
                    <tr key={i} style={{ borderTop: `1px solid ${C.line}` }}>
                      <td className="px-2 py-1.5">{r.firstName} {r.lastName}</td>
                      <td className="px-2 py-1.5">{r.username}</td>
                      <td className="px-2 py-1.5">
                        {r.subjectText}
                        {r.unmatched.length > 0 && (
                          <div style={{ color: C.red }} className="text-[11px] mt-0.5">Topilmadi: {r.unmatched.join(", ")}</div>
                        )}
                      </td>
                      <td className="px-2 py-1.5">
                        {!r.valid ? <span style={{ color: C.red }}>✕ Ma’lumot yetarli emas</span> : dup ? <span style={{ color: C.goldDark }}>⚠ Login band, o‘tkazib yuboriladi</span> : <span style={{ color: C.teal }}>✓ Qo‘shiladi</span>}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div className="flex justify-end gap-2 mt-4">
            <Btn kind="ghost" onClick={() => setImportRows(null)}>Bekor qilish</Btn>
            <Btn onClick={confirmImport}>Import qilish</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ---------- time slots editor ---------- */
function SlotsEditor({ slots, onSave, onDelete, onMove }) {
  const [modal, setModal] = useState(null);
  const empty = { id: null, type: "dars", name: "", start: "", end: "" };

  return (
    <div>
      <div className="flex justify-end mb-4">
        <Btn onClick={() => setModal({ ...empty })}>+ Vaqt oralig‘i qo‘shish</Btn>
      </div>
      <div className="space-y-2">
        {slots.map((s, i) => (
          <div key={s.id} style={{ borderColor: C.line, background: s.type === "tanaffus" ? "#FAFAFA" : "#fff" }} className="flex items-center gap-3 border rounded-md px-3 py-2 text-sm">
            <div className="flex flex-col mr-1">
              <button disabled={i === 0} onClick={() => onMove(s.id, -1)} style={{ color: i === 0 ? "#CCC" : C.textSoft }} className="text-xs leading-none">▲</button>
              <button disabled={i === slots.length - 1} onClick={() => onMove(s.id, 1)} style={{ color: i === slots.length - 1 ? "#CCC" : C.textSoft }} className="text-xs leading-none">▼</button>
            </div>
            <span style={{ background: s.type === "dars" ? C.navy : C.textSoft, color: "#fff" }} className="text-[10px] uppercase font-semibold px-2 py-0.5 rounded-full">{s.type === "dars" ? "Dars" : "Tanaffus"}</span>
            <span style={{ color: C.textMain }} className="font-medium flex-1">{s.name}</span>
            <span style={{ color: C.textSoft }}>{s.start} – {s.end}</span>
            <Btn small kind="ghost" onClick={() => setModal({ ...s })}>✎</Btn>
            <Btn small kind="danger" onClick={() => onDelete(s.id)}>O‘chirish</Btn>
          </div>
        ))}
        {slots.length === 0 && <p style={{ color: C.textSoft }} className="text-sm">Vaqt jadvali hali kiritilmagan.</p>}
      </div>

      {modal && (
        <Modal title={modal.id ? "Vaqt oralig‘ini tahrirlash" : "Yangi vaqt oralig‘i"} onClose={() => setModal(null)}>
          <Field label="Turi">
            <div className="flex gap-2">
              <button type="button" onClick={() => setModal({ ...modal, type: "dars" })} style={modal.type === "dars" ? { background: C.navy, color: "#fff" } : { border: `1px solid ${C.line}`, color: C.textMain }} className="px-3 py-1.5 rounded-md text-sm">Dars</button>
              <button type="button" onClick={() => setModal({ ...modal, type: "tanaffus" })} style={modal.type === "tanaffus" ? { background: C.navy, color: "#fff" } : { border: `1px solid ${C.line}`, color: C.textMain }} className="px-3 py-1.5 rounded-md text-sm">Tanaffus</button>
            </div>
          </Field>
          <Field label="Nomi (masalan: 1-dars)"><input className={inputCls} style={inputStyle} value={modal.name} onChange={(e) => setModal({ ...modal, name: e.target.value })} /></Field>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Boshlanish"><input type="time" className={inputCls} style={inputStyle} value={modal.start} onChange={(e) => setModal({ ...modal, start: e.target.value })} /></Field>
            <Field label="Tugash"><input type="time" className={inputCls} style={inputStyle} value={modal.end} onChange={(e) => setModal({ ...modal, end: e.target.value })} /></Field>
          </div>
          <div className="flex justify-end gap-2 mt-4">
            <Btn kind="ghost" onClick={() => setModal(null)}>Bekor qilish</Btn>
            <Btn onClick={() => { if (!modal.name) return; onSave({ ...modal, id: modal.id || uid() }); setModal(null); }}>Saqlash</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ---------- lesson entry modal (the ➕) ---------- */
const teachersFor = (db, subjectId) => {
  if (!subjectId) return db.teachers;
  const filtered = db.teachers.filter((t) => t.subjectIds?.includes(subjectId));
  return filtered.length > 0 ? filtered : db.teachers;
};

function LessonModal({ db, ctx, onClose, onSave }) {
  const [step, setStep] = useState("ask"); // ask | form
  const [split, setSplit] = useState(false);
  const [groups, setGroups] = useState([{ subjectId: "", teacherId: "", roomId: "", comment: "" }]);

  const chooseSplit = (isSplit) => {
    setSplit(isSplit);
    setGroups(isSplit ? [{ subjectId: "", teacherId: "", roomId: "", comment: "" }, { subjectId: "", teacherId: "", roomId: "", comment: "" }] : [{ subjectId: "", teacherId: "", roomId: "", comment: "" }]);
    setStep("form");
  };

  const updateGroup = (i, field, val) => setGroups((gs) => gs.map((g, idx) => (idx === i ? { ...g, [field]: val } : g)));
  const addGroup = () => groups.length < 4 && setGroups((gs) => [...gs, { subjectId: "", teacherId: "", roomId: "", comment: "" }]);
  const removeGroup = (i) => groups.length > 2 && setGroups((gs) => gs.filter((_, idx) => idx !== i));

  const valid = groups.every((g) => g.subjectId && g.teacherId && g.roomId);

  return (
    <Modal title="Darsni qayd etish" onClose={onClose} wide={step === "form"}>
      {step === "ask" && (
        <div className="text-center py-4">
          <p style={{ color: C.textMain }} className="mb-5 text-sm">Bu dars guruhlarga bo‘linadimi?</p>
          <div className="flex justify-center gap-3">
            <Btn kind="dark" onClick={() => chooseSplit(false)}>Yo‘q, bitta dars</Btn>
            <Btn onClick={() => chooseSplit(true)}>Ha, bo‘linadi</Btn>
          </div>
        </div>
      )}
      {step === "form" && (
        <div>
          <div className={split ? "grid sm:grid-cols-2 gap-4" : ""}>
            {groups.map((g, i) => (
              <div key={i} style={{ borderColor: C.line }} className="border rounded-md p-3 mb-3">
                {split && (
                  <div className="flex items-center justify-between mb-2">
                    <span style={{ color: C.navy }} className="text-xs font-semibold uppercase">{i + 1}-guruh</span>
                    {groups.length > 2 && <button onClick={() => removeGroup(i)} style={{ color: C.red }} className="text-xs">Olib tashlash</button>}
                  </div>
                )}
                <Field label="Fan">
                  <select className={inputCls} style={inputStyle} value={g.subjectId} onChange={(e) => { updateGroup(i, "subjectId", e.target.value); updateGroup(i, "teacherId", ""); }}>
                    <option value="">Tanlang</option>
                    {db.subjects.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                  </select>
                </Field>
                <Field label="O‘qituvchi">
                  <select className={inputCls} style={inputStyle} value={g.teacherId} onChange={(e) => updateGroup(i, "teacherId", e.target.value)}>
                    <option value="">Tanlang</option>
                    {teachersFor(db, g.subjectId).map((t) => <option key={t.id} value={t.id}>{t.firstName} {t.lastName}</option>)}
                  </select>
                  {g.subjectId && teachersFor(db, g.subjectId).length < db.teachers.length && (
                    <span style={{ color: C.textSoft }} className="text-[11px] block mt-1">Faqat shu fandan dars beradigan o‘qituvchilar ko‘rsatilmoqda.</span>
                  )}
                </Field>
                <Field label="Xona">
                  <select className={inputCls} style={inputStyle} value={g.roomId} onChange={(e) => updateGroup(i, "roomId", e.target.value)}>
                    <option value="">Tanlang</option>
                    {db.rooms.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}
                  </select>
                </Field>
                <Field label="Izoh">
                  <textarea className={inputCls} style={inputStyle} rows={2} value={g.comment} onChange={(e) => updateGroup(i, "comment", e.target.value)} placeholder="Ixtiyoriy izoh" />
                </Field>
              </div>
            ))}
          </div>
          {split && groups.length < 4 && <button onClick={addGroup} style={{ color: C.gold }} className="text-sm font-medium mb-3">+ Yana guruh qo‘shish</button>}
          <div className="flex justify-between items-center mt-2">
            <button onClick={() => setStep("ask")} style={{ color: C.textSoft }} className="text-sm">← Orqaga</button>
            <div className="flex gap-2">
              <Btn kind="ghost" onClick={onClose}>Bekor qilish</Btn>
              <Btn onClick={() => valid && onSave({ split, entries: groups.map((g) => ({ id: uid(), ...g })) })}>Saqlash</Btn>
            </div>
          </div>
        </div>
      )}
    </Modal>
  );
}

/* ---------- temporary-edit (🔁) modal ---------- */
function TempEditModal({ db, cell, onClose, onAdd, onRemove }) {
  const [date, setDate] = useState("");
  const [forms, setForms] = useState(cell.entries.map((e) => ({ ...e })));

  const update = (i, field, val) => setForms((fs) => fs.map((f, idx) => (idx === i ? { ...f, [field]: val } : f)));
  const valid = date && forms.every((f) => f.subjectId && f.teacherId && f.roomId);

  return (
    <Modal title="Vaqtincha tahrirlash" onClose={onClose} wide>
      {cell.tempEdits?.length > 0 && (
        <div className="mb-5">
          <p style={{ color: C.textSoft }} className="text-xs uppercase font-medium mb-2">Mavjud vaqtinchalik o‘zgarishlar</p>
          <div className="space-y-2">
            {cell.tempEdits.map((te) => (
              <div key={te.id} style={{ borderColor: C.tempBorder, background: C.temp }} className="border rounded-md px-3 py-2 flex items-center justify-between text-sm">
                <div style={{ color: C.textMain }}>
                  <span className="font-medium">{te.date}</span> —{" "}
                  {te.entries.map((e, i) => {
                    const subj = db.subjects.find((s) => s.id === e.subjectId)?.name || "?";
                    const teach = db.teachers.find((t) => t.id === e.teacherId);
                    return <span key={i}>{i > 0 ? " / " : ""}{subj} ({teach ? teach.firstName : "?"})</span>;
                  })}
                </div>
                <button onClick={() => onRemove(te.id)} style={{ color: C.red }} className="text-xs font-medium">O‘chirish</button>
              </div>
            ))}
          </div>
        </div>
      )}
      <p style={{ color: C.textSoft }} className="text-xs uppercase font-medium mb-2">Yangi vaqtinchalik o‘zgarish qo‘shish</p>
      <Field label="Sana"><input type="date" className={inputCls} style={inputStyle} value={date} onChange={(e) => setDate(e.target.value)} /></Field>
      <div className={forms.length > 1 ? "grid sm:grid-cols-2 gap-4" : ""}>
        {forms.map((f, i) => (
          <div key={i} style={{ borderColor: C.line }} className="border rounded-md p-3 mb-3">
            {forms.length > 1 && <span style={{ color: C.navy }} className="text-xs font-semibold uppercase block mb-2">{i + 1}-guruh</span>}
            <Field label="Fan">
              <select className={inputCls} style={inputStyle} value={f.subjectId} onChange={(e) => { update(i, "subjectId", e.target.value); update(i, "teacherId", ""); }}>
                <option value="">Tanlang</option>
                {db.subjects.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            </Field>
            <Field label="O‘qituvchi">
              <select className={inputCls} style={inputStyle} value={f.teacherId} onChange={(e) => update(i, "teacherId", e.target.value)}>
                <option value="">Tanlang</option>
                {teachersFor(db, f.subjectId).map((t) => <option key={t.id} value={t.id}>{t.firstName} {t.lastName}</option>)}
              </select>
            </Field>
            <Field label="Xona">
              <select className={inputCls} style={inputStyle} value={f.roomId} onChange={(e) => update(i, "roomId", e.target.value)}>
                <option value="">Tanlang</option>
                {db.rooms.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}
              </select>
            </Field>
            <Field label="Izoh"><textarea className={inputCls} style={inputStyle} rows={2} value={f.comment} onChange={(e) => update(i, "comment", e.target.value)} /></Field>
          </div>
        ))}
      </div>
      <div className="flex justify-end gap-2 mt-2">
        <Btn kind="ghost" onClick={onClose}>Yopish</Btn>
        <Btn onClick={() => valid && onAdd({ id: uid(), date, entries: forms })}>Qo‘shish</Btn>
      </div>
    </Modal>
  );
}

/* ---------- schedule grid ---------- */
function ScheduleView({ db, mutate, onClose }) {
  const [classId, setClassId] = useState(db.classes[0]?.id || "");
  const [lessonCtx, setLessonCtx] = useState(null); // {dayId, slotId}
  const [tempCtx, setTempCtx] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const [importGroups, setImportGroups] = useState(null); // preview groups after parsing
  const [importInvalid, setImportInvalid] = useState([]);
  const [importErr, setImportErr] = useState("");
  const [overwrite, setOverwrite] = useState(false);

  const displayName = (id, arr) => arr.find((x) => x.id === id)?.name || "?";
  const displayTeacher = (id) => { const t = db.teachers.find((x) => x.id === id); return t ? `${t.firstName} ${t.lastName}` : "?"; };

  const UZ_WEEKDAYS = ["Yakshanba", "Dushanba", "Seshanba", "Chorshanba", "Payshanba", "Juma", "Shanba"];
  const todayStr = new Date().toISOString().slice(0, 10);
  const todayWeekdayName = UZ_WEEKDAYS[new Date().getDay()];
  const todayDay = db.days.find((d) => d.name.trim().toLowerCase() === todayWeekdayName.toLowerCase());
  const absentTeacherIdsToday = new Set(
    (db.absenceRequests || []).filter((r) => todayStr >= r.startDate && todayStr <= r.endDate).map((r) => r.teacherId)
  );

  const normKey = (k) => String(k).toLowerCase().replace(/[’'ʻ‘`]/g, "").replace(/\s+/g, " ").trim();
  const SCHEDULE_HEADERS = {
    sinf: "className", class: "className",
    kun: "dayName", day: "dayName",
    "dars soati": "slotName", dars: "slotName", soat: "slotName", slot: "slotName", vaqt: "slotName",
    fan: "subjectName", subject: "subjectName",
    oqituvchi: "teacherName", teacher: "teacherName",
    xona: "roomName", room: "roomName",
    izoh: "comment", comment: "comment",
    guruh: "group", group: "group",
  };

  const parseScheduleFile = (raw) => raw.map((row) => {
    const norm = {};
    Object.keys(row).forEach((k) => {
      const key = SCHEDULE_HEADERS[normKey(k)];
      if (key) norm[key] = String(row[k] ?? "").trim();
    });
    const cls = db.classes.find((c) => c.name.toLowerCase() === (norm.className || "").toLowerCase());
    const day = db.days.find((d) => d.name.toLowerCase() === (norm.dayName || "").toLowerCase());
    const slot = db.slots.find((s) => s.type === "dars" && s.name.toLowerCase() === (norm.slotName || "").toLowerCase());
    const subject = db.subjects.find((s) => s.name.toLowerCase() === (norm.subjectName || "").toLowerCase());
    const teacher = db.teachers.find((t) => `${t.firstName} ${t.lastName}`.toLowerCase() === (norm.teacherName || "").toLowerCase());
    const room = db.rooms.find((r) => r.name.toLowerCase() === (norm.roomName || "").toLowerCase());
    const reasons = [];
    if (!cls) reasons.push("Sinf");
    if (!day) reasons.push("Kun");
    if (!slot) reasons.push("Dars soati");
    if (!subject) reasons.push("Fan");
    if (!teacher) reasons.push("O‘qituvchi");
    if (!room) reasons.push("Xona");
    return {
      norm,
      classId: cls?.id, dayId: day?.id, slotId: slot?.id,
      subjectId: subject?.id, teacherId: teacher?.id, roomId: room?.id,
      comment: norm.comment || "", group: norm.group || "",
      valid: reasons.length === 0, reasons,
    };
  });

  const groupValidRows = (rows) => {
    const map = new Map();
    rows.filter((r) => r.valid).forEach((r) => {
      const key = `${r.classId}|${r.dayId}|${r.slotId}`;
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(r);
    });
    return [...map.entries()].map(([key, list]) => {
      list.sort((a, b) => a.group.localeCompare(b.group, undefined, { numeric: true }));
      const [gClassId, gDayId, gSlotId] = key.split("|");
      const existing = db.schedule?.[gClassId]?.[gDayId]?.[gSlotId];
      const occupied = !!(existing && existing.entries?.length > 0);
      return { key, classId: gClassId, dayId: gDayId, slotId: gSlotId, rows: list, occupied };
    });
  };

  const handleScheduleFile = (e) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    setImportErr("");
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });
        if (raw.length === 0) { setImportErr("Faylda ma’lumot topilmadi."); return; }
        const rows = parseScheduleFile(raw);
        setImportGroups(groupValidRows(rows));
        setImportInvalid(rows.filter((r) => !r.valid));
      } catch {
        setImportErr("Faylni o‘qib bo‘lmadi. .xlsx yoki .csv formatdagi faylni tanlang.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const downloadScheduleTemplate = () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      ["Sinf", "Kun", "Dars soati", "Fan", "O‘qituvchi", "Xona", "Izoh", "Guruh"],
      ["5-A", "Dushanba", "1-dars", "Matematika", "Aziza Karimova", "204-xona", "", ""],
      ["5-A", "Dushanba", "2-dars", "Ingliz tili", "Jasur Yusupov", "101-xona", "", "1"],
      ["5-A", "Dushanba", "2-dars", "Rus tili", "Olga Petrova", "102-xona", "", "2"],
    ]);
    XLSX.utils.book_append_sheet(wb, ws, "Jadval");
    XLSX.writeFile(wb, "dars-jadvali-namuna.xlsx");
  };

  const applyScheduleImport = () => {
    mutate((d) => {
      importGroups.forEach((g) => {
        if (g.occupied && !overwrite) return;
        if (!d.schedule[g.classId]) d.schedule[g.classId] = {};
        if (!d.schedule[g.classId][g.dayId]) d.schedule[g.classId][g.dayId] = {};
        d.schedule[g.classId][g.dayId][g.slotId] = {
          split: g.rows.length > 1,
          entries: g.rows.map((r) => ({ id: uid(), subjectId: r.subjectId, teacherId: r.teacherId, roomId: r.roomId, comment: r.comment })),
          tempEdits: [],
        };
      });
      return d;
    });
    setImportGroups(null);
    setImportInvalid([]);
  };

  const getCell = (dayId, slotId) => db.schedule?.[classId]?.[dayId]?.[slotId];

  const setCell = (dayId, slotId, value) => {
    mutate((d) => {
      if (!d.schedule[classId]) d.schedule[classId] = {};
      if (!d.schedule[classId][dayId]) d.schedule[classId][dayId] = {};
      d.schedule[classId][dayId][slotId] = value;
      return d;
    });
  };

  const removeCell = (dayId, slotId) => {
    mutate((d) => {
      if (d.schedule?.[classId]?.[dayId]?.[slotId]) delete d.schedule[classId][dayId][slotId];
      return d;
    });
  };

  const name = (arr, id) => arr.find((x) => x.id === id)?.name || arr.find((x) => x.id === id)?.firstName || "";
  const teacherName = (id) => { const t = db.teachers.find((x) => x.id === id); return t ? `${t.firstName} ${t.lastName}` : "?"; };

  const darsSlots = db.slots; // includes breaks, rendered differently

  if (db.classes.length === 0) {
    return (
      <div className="p-8 text-center">
        <p style={{ color: C.textSoft }}>Avval admin panelda kamida bitta sinf qo‘shing.</p>
        <div className="mt-4"><Btn kind="ghost" onClick={onClose}>Orqaga</Btn></div>
      </div>
    );
  }

  return (
    <div>
      <div className="flex flex-wrap items-center gap-3 mb-5">
        <Btn kind="ghost" onClick={onClose}>← Admin panelga</Btn>
        <span style={{ color: C.textSoft }} className="text-sm">Sinf:</span>
        <select className={inputCls} style={{ ...inputStyle, width: "auto" }} value={classId} onChange={(e) => setClassId(e.target.value)}>
          {db.classes.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
        </select>
        <div className="flex-1" />
        <Btn kind="ghost" onClick={downloadScheduleTemplate}>⬇ Namuna faylni yuklab olish</Btn>
        <label style={{ background: C.teal, color: "#fff" }} className="rounded-md font-medium px-4 py-2 text-sm hover:opacity-90 transition cursor-pointer">
          📥 Jadvalni Excel orqali import qilish
          <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleScheduleFile} />
        </label>
      </div>
      {importErr && <p style={{ color: C.red }} className="text-xs mb-3">{importErr}</p>}

      {importGroups && (
        <Modal title="Jadvalni Excel’dan import qilish — tekshirish" onClose={() => { setImportGroups(null); setImportInvalid([]); }} wide>
          <p style={{ color: C.textSoft }} className="text-sm mb-3">
            Ustunlar: <b>Sinf, Kun, Dars soati, Fan, O‘qituvchi, Xona, Izoh, Guruh</b>. Bir xil Sinf+Kun+Dars soati bilan bir necha qator kiritilsa (turli Guruh raqami bilan), dars bo‘lingan deb hisoblanadi.
          </p>
          <label className="flex items-center gap-2 mb-3 text-sm" style={{ color: C.textMain }}>
            <input type="checkbox" checked={overwrite} onChange={(e) => setOverwrite(e.target.checked)} />
            Band bo‘lgan darslarni yangi ma’lumot bilan almashtirish
          </label>
          <div className="max-h-64 overflow-y-auto border rounded-md mb-3" style={{ borderColor: C.line }}>
            <table className="w-full text-xs">
              <thead>
                <tr style={{ background: C.bg }}>
                  <th className="text-left px-2 py-1.5">Sinf</th>
                  <th className="text-left px-2 py-1.5">Kun</th>
                  <th className="text-left px-2 py-1.5">Dars soati</th>
                  <th className="text-left px-2 py-1.5">Darslar</th>
                  <th className="text-left px-2 py-1.5">Holat</th>
                </tr>
              </thead>
              <tbody>
                {importGroups.map((g) => (
                  <tr key={g.key} style={{ borderTop: `1px solid ${C.line}` }}>
                    <td className="px-2 py-1.5">{displayName(g.classId, db.classes)}</td>
                    <td className="px-2 py-1.5">{displayName(g.dayId, db.days)}</td>
                    <td className="px-2 py-1.5">{displayName(g.slotId, db.slots)}</td>
                    <td className="px-2 py-1.5">
                      {g.rows.map((r, i) => (
                        <div key={i}>{displayName(r.subjectId, db.subjects)} — {displayTeacher(r.teacherId)} ({displayName(r.roomId, db.rooms)})</div>
                      ))}
                    </td>
                    <td className="px-2 py-1.5">
                      {!g.occupied ? <span style={{ color: C.teal }}>✓ Qo‘shiladi</span> : overwrite ? <span style={{ color: C.goldDark }}>⚠ Band, almashtiriladi</span> : <span style={{ color: C.red }}>✕ Band, o‘tkazib yuboriladi</span>}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {importInvalid.length > 0 && (
            <div className="mb-3">
              <p style={{ color: C.red }} className="text-xs font-medium mb-1">{importInvalid.length} ta qatorda xatolik topildi (import qilinmaydi):</p>
              <div className="max-h-32 overflow-y-auto border rounded-md text-xs" style={{ borderColor: C.line }}>
                {importInvalid.map((r, i) => (
                  <div key={i} style={{ borderTop: i > 0 ? `1px solid ${C.line}` : "none" }} className="px-2 py-1.5" >
                    {r.norm.className || "?"} / {r.norm.dayName || "?"} / {r.norm.slotName || "?"} — <span style={{ color: C.red }}>topilmadi: {r.reasons.join(", ")}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
          <div className="flex justify-end gap-2">
            <Btn kind="ghost" onClick={() => { setImportGroups(null); setImportInvalid([]); }}>Bekor qilish</Btn>
            <Btn onClick={applyScheduleImport}>Import qilish</Btn>
          </div>
        </Modal>
      )}

      <div className="overflow-x-auto rounded-lg border" style={{ borderColor: C.line }}>
        <table className="w-full border-collapse text-sm min-w-[800px]">
          <thead>
            <tr>
              <th style={{ background: C.navy, color: "#fff", borderColor: C.line }} className="border px-3 py-2 text-left w-40">Vaqt</th>
              {db.days.map((d) => (
                <th key={d.id} style={{ background: C.navy, color: "#fff", borderColor: C.line }} className="border px-3 py-2 text-center">{d.name}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {darsSlots.map((slot) => {
              if (slot.type === "tanaffus") {
                return (
                  <tr key={slot.id} style={{ background: "#EEF0F4" }}>
                    <td style={{ borderColor: C.line, color: C.textSoft }} className="border px-3 py-1.5 text-xs font-medium">{slot.name} ({slot.start}–{slot.end})</td>
                    <td colSpan={db.days.length} style={{ borderColor: C.line, color: C.textSoft }} className="border px-3 py-1.5 text-center text-xs">tanaffus</td>
                  </tr>
                );
              }
              return (
                <tr key={slot.id}>
                  <td style={{ borderColor: C.line, color: C.textMain }} className="border px-3 py-2 align-top">
                    <div className="font-medium">{slot.name}</div>
                    <div style={{ color: C.textSoft }} className="text-xs">{slot.start}–{slot.end}</div>
                  </td>
                  {db.days.map((day) => {
                    const cell = getCell(day.id, slot.id);
                    const hasTemp = cell?.tempEdits?.length > 0;
                    const isToday = todayDay && day.id === todayDay.id;
                    const absentToday = isToday && cell?.entries?.some((e) => absentTeacherIdsToday.has(e.teacherId));
                    return (
                      <td
                        key={day.id}
                        style={{ borderColor: absentToday ? C.red : C.line, background: absentToday ? undefined : hasTemp ? C.temp : "#fff" }}
                        className={`border p-1.5 align-top relative min-w-[150px] ${absentToday ? "cell-blink" : ""}`}
                      >
                        {!cell || cell.entries.length === 0 ? (
                          <button
                            onClick={() => setLessonCtx({ dayId: day.id, slotId: slot.id })}
                            style={{ color: C.gold, borderColor: C.line }}
                            className="w-full h-14 flex items-center justify-center text-xl rounded-md border border-dashed hover:bg-gray-50"
                            title="Dars qo‘shish"
                          >
                            ➕
                          </button>
                        ) : (
                          <div style={{ borderColor: absentToday ? C.red : hasTemp ? C.tempBorder : C.line }} className="rounded-md border p-2 relative">
                            <div className="absolute top-1 right-1 flex gap-1">
                              <button title="Vaqtincha tahrirlash" onClick={() => setTempCtx({ dayId: day.id, slotId: slot.id })} className="text-xs hover:scale-110 transition">🔁</button>
                              <button title="O‘chirish" onClick={() => setConfirmDel({ dayId: day.id, slotId: slot.id })} className="text-xs hover:scale-110 transition">🚮</button>
                            </div>
                            {absentToday && (
                              <div style={{ color: C.red }} className="flex items-center gap-1 text-[10px] font-semibold mb-1">
                                <AlertTriangle size={12} strokeWidth={2} /> Bugun kelmaydi — almashtiring
                              </div>
                            )}
                            <div className={cell.split ? "space-y-1.5 pr-8" : "pr-8"}>
                              {cell.entries.map((e, i) => (
                                <div key={e.id} style={{ borderTop: i > 0 && cell.split ? `1px dashed ${C.line}` : "none", paddingTop: i > 0 && cell.split ? 4 : 0 }}>
                                  <div style={{ color: C.navy }} className="font-semibold text-xs">{name(db.subjects, e.subjectId)}</div>
                                  <div style={{ color: absentTeacherIdsToday.has(e.teacherId) && isToday ? C.red : C.textMain }} className="text-xs font-medium">{teacherName(e.teacherId)}</div>
                                  <div style={{ color: C.textSoft }} className="text-[11px]">{name(db.rooms, e.roomId)}</div>
                                  {e.comment && <div style={{ color: C.teal }} className="text-[11px] italic">💬 {e.comment}</div>}
                                </div>
                              ))}
                            </div>
                            {hasTemp && <div style={{ color: C.goldDark }} className="text-[10px] font-medium mt-1">⏱ {cell.tempEdits.length} ta vaqtinchalik o‘zgarish</div>}
                          </div>
                        )}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {lessonCtx && (
        <LessonModal
          db={db}
          ctx={lessonCtx}
          onClose={() => setLessonCtx(null)}
          onSave={(val) => { setCell(lessonCtx.dayId, lessonCtx.slotId, { ...val, tempEdits: [] }); setLessonCtx(null); }}
        />
      )}

      {tempCtx && getCell(tempCtx.dayId, tempCtx.slotId) && (
        <TempEditModal
          db={db}
          cell={getCell(tempCtx.dayId, tempCtx.slotId)}
          onClose={() => setTempCtx(null)}
          onAdd={(te) => {
            mutate((d) => {
              const c = d.schedule[classId][tempCtx.dayId][tempCtx.slotId];
              c.tempEdits = [...(c.tempEdits || []), te];
              return d;
            });
          }}
          onRemove={(id) => {
            mutate((d) => {
              const c = d.schedule[classId][tempCtx.dayId][tempCtx.slotId];
              c.tempEdits = c.tempEdits.filter((t) => t.id !== id);
              return d;
            });
          }}
        />
      )}

      {confirmDel && (
        <Modal title="Darsni o‘chirish" onClose={() => setConfirmDel(null)}>
          <p style={{ color: C.textMain }} className="text-sm mb-4">Ushbu darsni o‘chirmoqchimisiz? Bu amalni ortga qaytarib bo‘lmaydi.</p>
          <div className="flex justify-end gap-2">
            <Btn kind="ghost" onClick={() => setConfirmDel(null)}>Bekor qilish</Btn>
            <Btn kind="danger" onClick={() => { removeCell(confirmDel.dayId, confirmDel.slotId); setConfirmDel(null); }}>O‘chirish</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ---------- absence request modal (teacher side) ---------- */
function AbsenceModal({ onClose, onSubmit }) {
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [reason, setReason] = useState("");
  const [err, setErr] = useState("");

  const submit = () => {
    if (!startDate || !endDate || !reason.trim()) { setErr("Barcha maydonlarni to‘ldiring"); return; }
    if (endDate < startDate) { setErr("Tugash sanasi boshlanish sanasidan oldin bo‘lishi mumkin emas"); return; }
    onSubmit({ startDate, endDate, reason: reason.trim() });
  };

  return (
    <Modal title="Ishda bo‘lmayman" onClose={onClose}>
      <p style={{ color: C.textSoft }} className="text-sm mb-4">Ishga kelolmaydigan kunlaringizni va sababini kiriting — so‘rov adminga bildirishnoma sifatida yuboriladi.</p>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Sanadan"><input type="date" className={inputCls} style={inputStyle} value={startDate} onChange={(e) => setStartDate(e.target.value)} /></Field>
        <Field label="Sanagacha"><input type="date" className={inputCls} style={inputStyle} value={endDate} onChange={(e) => setEndDate(e.target.value)} /></Field>
      </div>
      <Field label="Sababi">
        <textarea className={inputCls} style={inputStyle} rows={3} value={reason} onChange={(e) => setReason(e.target.value)} placeholder="Masalan: sog‘liq holati, oilaviy sabab..." />
      </Field>
      {err && <p style={{ color: C.red }} className="text-xs mb-2">{err}</p>}
      <div className="flex justify-end gap-2 mt-2">
        <Btn kind="ghost" onClick={onClose}>Bekor qilish</Btn>
        <Btn onClick={submit}>Yuborish</Btn>
      </div>
    </Modal>
  );
}

/* ---------- task assignment (admin -> teacher) ---------- */
function TaskAssignEditor({ teachers, tasks, onAssign, onDelete, onSeenReply }) {
  const [modal, setModal] = useState(null); // {teacherId, title, description}
  const empty = { teacherId: "", title: "", description: "" };

  const teacherName = (id) => { const t = teachers.find((x) => x.id === id); return t ? `${t.firstName} ${t.lastName}` : "?"; };

  return (
    <div>
      <div className="flex justify-end mb-4">
        <Btn onClick={() => setModal({ ...empty })}>+ Vazifa berish</Btn>
      </div>
      {tasks.length === 0 && <p style={{ color: C.textSoft }} className="text-sm">Hozircha vazifalar yo‘q.</p>}
      <div className="space-y-2">
        {[...tasks].sort((a, b) => (b.createdAt || "").localeCompare(a.createdAt || "")).map((t) => (
          <div key={t.id} style={{ borderColor: t.reply && !t.replySeen ? C.tempBorder : C.line, background: t.reply && !t.replySeen ? C.temp : "#fff" }} className="border rounded-md px-3 py-2.5 text-sm">
            <div className="flex items-start justify-between gap-2">
              <div className="flex-1">
                <div style={{ color: C.navy }} className="font-semibold">{t.title}</div>
                <div style={{ color: C.textMain }} className="text-xs mt-0.5">{teacherName(t.teacherId)}</div>
                {t.description && <div style={{ color: C.textSoft }} className="text-xs mt-1">{t.description}</div>}
                {t.reply && (
                  <div style={{ borderColor: C.line, background: "#fff" }} className="border rounded-md px-2.5 py-2 mt-2">
                    <div style={{ color: C.textSoft }} className="text-[11px] uppercase font-medium mb-0.5">O‘qituvchi javobi</div>
                    <div style={{ color: C.textMain }} className="text-xs whitespace-pre-wrap">{t.reply}</div>
                  </div>
                )}
              </div>
              <div className="flex items-center gap-2 shrink-0">
                <span style={{ color: t.read ? C.teal : C.goldDark }} className="text-xs font-medium">{t.read ? "✓ O‘qildi" : "⏳ Ochilmagan"}</span>
                {t.reply && !t.replySeen && <Btn small onClick={() => onSeenReply(t.id)}>Javobni ko‘rdim</Btn>}
                <Btn small kind="danger" onClick={() => onDelete(t.id)}>O‘chirish</Btn>
              </div>
            </div>
          </div>
        ))}
      </div>

      {modal && (
        <Modal title="Yangi vazifa" onClose={() => setModal(null)}>
          <Field label="O‘qituvchi">
            <select className={inputCls} style={inputStyle} value={modal.teacherId} onChange={(e) => setModal({ ...modal, teacherId: e.target.value })}>
              <option value="">Tanlang</option>
              {teachers.map((t) => <option key={t.id} value={t.id}>{t.firstName} {t.lastName}</option>)}
            </select>
          </Field>
          <Field label="Vazifa sarlavhasi"><input className={inputCls} style={inputStyle} value={modal.title} onChange={(e) => setModal({ ...modal, title: e.target.value })} /></Field>
          <Field label="Tavsif"><textarea className={inputCls} style={inputStyle} rows={4} value={modal.description} onChange={(e) => setModal({ ...modal, description: e.target.value })} placeholder="Vazifa haqida batafsil ma’lumot" /></Field>
          <div className="flex justify-end gap-2 mt-4">
            <Btn kind="ghost" onClick={() => setModal(null)}>Bekor qilish</Btn>
            <Btn onClick={() => { if (!modal.teacherId || !modal.title.trim()) return; onAssign({ id: uid(), teacherId: modal.teacherId, title: modal.title.trim(), description: modal.description.trim(), createdAt: new Date().toISOString(), read: false, reply: "", repliedAt: null, replySeen: true }); setModal(null); }}>Yuborish</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}

function loadJsPDF() {
  if (window.jspdf && window.jspdf.jsPDF) return Promise.resolve(window.jspdf.jsPDF);
  return new Promise((resolve, reject) => {
    const existing = document.querySelector('script[data-jspdf-lib]');
    if (existing) {
      existing.addEventListener("load", () => resolve(window.jspdf.jsPDF));
      existing.addEventListener("error", reject);
      return;
    }
    const script = document.createElement("script");
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
    script.setAttribute("data-jspdf-lib", "true");
    script.onload = () => resolve(window.jspdf.jsPDF);
    script.onerror = reject;
    document.body.appendChild(script);
  });
}

function downloadMissingTopicsPDF(alertData) {
  loadJsPDF()
    .then((JsPDF) => {
      const doc = new JsPDF();
      doc.setFontSize(14);
      doc.text(`Kunlik mavzu nazorati - ${fmtDate(alertData.date)}`, 14, 18);
      doc.setFontSize(10);
      doc.text("Bugungi darsi bo'lib, mavzusi belgilanmagan o'qituvchi/sinf/fanlar:", 14, 26);
      let y = 38;
      doc.setFontSize(11);
      doc.setFont(undefined, "bold");
      doc.text("O'qituvchi", 14, y);
      doc.text("Sinf", 100, y);
      doc.text("Fan", 140, y);
      doc.setFont(undefined, "normal");
      y += 4;
      doc.setLineWidth(0.3);
      doc.line(14, y, 196, y);
      y += 8;
      alertData.items.forEach((it) => {
        if (y > 280) { doc.addPage(); y = 20; }
        doc.text(String(it.teacherName || "-"), 14, y);
        doc.text(`${it.className}-sinf`, 100, y);
        doc.text(String(it.subjectName || "-"), 140, y);
        y += 8;
      });
      doc.save(`kunlik-mavzu-nazorati-${alertData.date}.pdf`);
    })
    .catch(() => window.alert("PDF kutubxonasini yuklab bo'lmadi. Internet aloqasini tekshiring."));
}

/* ---------- admin: manage imported yearly plans ---------- */
function AdminYearlyPlansEditor({ db, mutate }) {
  const [filterClass, setFilterClass] = useState("");
  const [filterSubject, setFilterSubject] = useState("");
  const [confirmDel, setConfirmDel] = useState(null);

  const teacherName = (id) => { const t = db.teachers.find((x) => x.id === id); return t ? `${t.firstName} ${t.lastName}` : "?"; };
  const className = (id) => db.classes.find((c) => c.id === id)?.name || "?";
  const subjectName = (id) => db.subjects.find((s) => s.id === id)?.name || "?";

  const plans = (db.yearlyPlans || []).filter((p) => (!filterClass || p.classId === filterClass) && (!filterSubject || p.subjectId === filterSubject));

  return (
    <div>
      <div className="grid sm:grid-cols-2 gap-3 mb-4">
        <Field label="Sinf bo‘yicha filtr">
          <select className={inputCls} style={inputStyle} value={filterClass} onChange={(e) => setFilterClass(e.target.value)}>
            <option value="">Barchasi</option>
            {db.classes.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
        </Field>
        <Field label="Fan bo‘yicha filtr">
          <select className={inputCls} style={inputStyle} value={filterSubject} onChange={(e) => setFilterSubject(e.target.value)}>
            <option value="">Barchasi</option>
            {db.subjects.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        </Field>
      </div>

      {plans.length === 0 && <p style={{ color: C.textSoft }} className="text-sm">Mos ish reja topilmadi.</p>}
      <div className="space-y-2">
        {plans.map((p) => {
          const doneCount = p.rows.filter((r) => r.done).length;
          return (
            <div key={p.id} style={{ borderColor: C.line }} className="border rounded-md px-3 py-2.5 text-sm flex items-center justify-between gap-2">
              <div>
                <div style={{ color: C.navy }} className="font-semibold">{className(p.classId)}-sinf · {subjectName(p.subjectId)}</div>
                <div style={{ color: C.textSoft }} className="text-xs mt-0.5">O‘qituvchi: {teacherName(p.teacherId)} · {doneCount}/{p.rows.length} bajarilgan · Import: {fmtDate(p.importedAt)}</div>
              </div>
              <Btn small kind="danger" onClick={() => setConfirmDel(p)}>O‘chirish</Btn>
            </div>
          );
        })}
      </div>

      {confirmDel && (
        <Modal title="Ish rejani o‘chirish" onClose={() => setConfirmDel(null)}>
          <p style={{ color: C.textMain }} className="text-sm mb-4">
            <b>{className(confirmDel.classId)}-sinf · {subjectName(confirmDel.subjectId)}</b> ({teacherName(confirmDel.teacherId)}) uchun butun ish reja va barcha belgilangan mavzular butunlay o‘chiriladi. Bu amalni ortga qaytarib bo‘lmaydi. Davom etasizmi?
          </p>
          <div className="flex justify-end gap-2">
            <Btn kind="ghost" onClick={() => setConfirmDel(null)}>Bekor qilish</Btn>
            <Btn kind="danger" onClick={() => { mutate((d) => { d.yearlyPlans = d.yearlyPlans.filter((x) => x.id !== confirmDel.id); return d; }); setConfirmDel(null); }}>Ha, o‘chirish</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ---------- admin panel ---------- */
function AdminPanel({ db, mutate, onOpenSchedule, onLogout }) {
  const [tab, setTab] = useState("teachers");
  const [notifOpen, setNotifOpen] = useState(false);
  const [bsbOpen, setBsbOpen] = useState(false);
  const [pwModal, setPwModal] = useState(null); // {current, next, confirm, err}
  const unreadCount = (db.absenceRequests || []).filter((r) => !r.read).length + (db.tasks || []).filter((t) => t.reply && !t.replySeen).length + (db.dailyTopicAlerts || []).filter((a) => !a.read).length;
  const unreadBsbCount = (db.bsbNotifications || []).filter((n) => !n.read).length;
  const tabs = [
    ["teachers", "O‘qituvchilar"],
    ["tasks", "Vazifalar"],
    ["yearlyPlans", "Ish rejalar"],
    ["subjects", "Fanlar"],
    ["days", "Hafta kunlari"],
    ["slots", "Dars soatlari"],
    ["classes", "Sinflar"],
    ["rooms", "Xonalar"],
  ];

  return (
    <div>
      <div className="flex items-center justify-between mb-6">
        <div>
          <h1 style={{ color: C.navy }} className="text-xl font-bold">Admin panel</h1>
          <p style={{ color: C.textSoft }} className="text-sm">Maktab dars jadvalini boshqarish</p>
        </div>
        <div className="flex items-center gap-2">
          <button onClick={() => setPwModal({ current: "", next: "", confirm: "", err: "" })} title="Parolni o‘zgartirish" style={{ background: C.navy, color: "#fff" }} className="w-11 h-11 rounded-lg flex items-center justify-center hover:opacity-90">
            <Settings size={20} strokeWidth={1.8} />
          </button>
          <button onClick={() => setBsbOpen(true)} title="BSB/CHSB nazorati" style={{ background: C.navy, color: "#fff" }} className="w-11 h-11 rounded-lg flex items-center justify-center hover:opacity-90 relative">
            <ClipboardList size={20} strokeWidth={1.8} />
            {unreadBsbCount > 0 && (
              <span style={{ background: C.red }} className="absolute -top-1.5 -right-1.5 text-white text-[10px] font-semibold w-5 h-5 rounded-full flex items-center justify-center">{unreadBsbCount}</span>
            )}
          </button>
          <button onClick={() => setNotifOpen(true)} title="Bildirishnomalar" style={{ background: C.navy, color: "#fff" }} className="w-11 h-11 rounded-lg flex items-center justify-center hover:opacity-90 relative">
            <Bell size={20} strokeWidth={1.8} />
            {unreadCount > 0 && (
              <span style={{ background: C.red }} className="absolute -top-1.5 -right-1.5 text-white text-[10px] font-semibold w-5 h-5 rounded-full flex items-center justify-center">{unreadCount}</span>
            )}
          </button>
          <button onClick={onOpenSchedule} title="Dars jadvali" style={{ background: C.navy, color: "#fff" }} className="w-11 h-11 rounded-lg flex items-center justify-center hover:opacity-90">
            <CalendarDays size={20} strokeWidth={1.8} />
          </button>
          <Btn kind="ghost" onClick={onLogout}>Chiqish</Btn>
        </div>
      </div>

      {pwModal && (
        <Modal title="Admin parolini o‘zgartirish" onClose={() => setPwModal(null)}>
          <Field label="Joriy parol"><input type="password" className={inputCls} style={inputStyle} value={pwModal.current} onChange={(e) => setPwModal({ ...pwModal, current: e.target.value })} /></Field>
          <Field label="Yangi parol"><input type="password" className={inputCls} style={inputStyle} value={pwModal.next} onChange={(e) => setPwModal({ ...pwModal, next: e.target.value })} /></Field>
          <Field label="Yangi parolni tasdiqlang"><input type="password" className={inputCls} style={inputStyle} value={pwModal.confirm} onChange={(e) => setPwModal({ ...pwModal, confirm: e.target.value })} /></Field>
          {pwModal.err && <p style={{ color: C.red }} className="text-xs mb-2">{pwModal.err}</p>}
          <div className="flex justify-end gap-2 mt-4">
            <Btn kind="ghost" onClick={() => setPwModal(null)}>Bekor qilish</Btn>
            <Btn
              onClick={() => {
                if (pwModal.current !== db.adminPassword) { setPwModal({ ...pwModal, err: "Joriy parol noto‘g‘ri" }); return; }
                if (!pwModal.next.trim()) { setPwModal({ ...pwModal, err: "Yangi parolni kiriting" }); return; }
                if (pwModal.next !== pwModal.confirm) { setPwModal({ ...pwModal, err: "Yangi parollar mos emas" }); return; }
                mutate((d) => { d.adminPassword = pwModal.next.trim(); return d; });
                setPwModal(null);
              }}
            >
              Saqlash
            </Btn>
          </div>
        </Modal>
      )}

      {notifOpen && (
        <Modal title="Bildirishnomalar" onClose={() => setNotifOpen(false)} wide>
          <h4 style={{ color: C.navy }} className="font-semibold text-sm mb-2">Ishda bo‘lmaslik so‘rovlari</h4>
          {(db.absenceRequests || []).length === 0 && <p style={{ color: C.textSoft }} className="text-sm mb-4">Hozircha so‘rovlar yo‘q.</p>}
          <div className="space-y-2 mb-5">
            {[...(db.absenceRequests || [])].sort((a, b) => (b.createdAt || "").localeCompare(a.createdAt || "")).map((r) => (
              <div key={r.id} style={{ borderColor: r.read ? C.line : C.tempBorder, background: r.read ? "#fff" : C.temp }} className="border rounded-md px-3 py-2.5 text-sm">
                <div className="flex items-start justify-between gap-2">
                  <div>
                    <div style={{ color: C.navy }} className="font-semibold">{r.teacherName}</div>
                    <div style={{ color: C.textMain }} className="text-xs mt-0.5">{r.startDate} — {r.endDate}</div>
                    <div style={{ color: C.textSoft }} className="text-xs mt-1 italic">{r.reason}</div>
                  </div>
                  <div className="flex gap-1.5 shrink-0">
                    {!r.read && <Btn small onClick={() => mutate((d) => { const req = d.absenceRequests.find((x) => x.id === r.id); if (req) req.read = true; return d; })}>Ko‘rib chiqildi</Btn>}
                    <Btn small kind="danger" onClick={() => mutate((d) => { d.absenceRequests = d.absenceRequests.filter((x) => x.id !== r.id); return d; })}>O‘chirish</Btn>
                  </div>
                </div>
              </div>
            ))}
          </div>

          <h4 style={{ color: C.navy }} className="font-semibold text-sm mb-2">Vazifalarga javoblar</h4>
          {(db.tasks || []).filter((t) => t.reply).length === 0 && <p style={{ color: C.textSoft }} className="text-sm">Hozircha javoblar yo‘q.</p>}
          <div className="space-y-2">
            {[...(db.tasks || [])].filter((t) => t.reply).sort((a, b) => (b.repliedAt || "").localeCompare(a.repliedAt || "")).map((t) => {
              const teacher = db.teachers.find((x) => x.id === t.teacherId);
              return (
                <div key={t.id} style={{ borderColor: t.replySeen ? C.line : C.tempBorder, background: t.replySeen ? "#fff" : C.temp }} className="border rounded-md px-3 py-2.5 text-sm">
                  <div className="flex items-start justify-between gap-2">
                    <div>
                      <div style={{ color: C.navy }} className="font-semibold">{teacher ? `${teacher.firstName} ${teacher.lastName}` : "?"} — {t.title}</div>
                      <div style={{ color: C.textMain }} className="text-xs mt-1 whitespace-pre-wrap">{t.reply}</div>
                    </div>
                    <div className="flex gap-1.5 shrink-0">
                      {!t.replySeen && <Btn small onClick={() => mutate((d) => { const task = d.tasks.find((x) => x.id === t.id); if (task) task.replySeen = true; return d; })}>Ko‘rdim</Btn>}
                      <Btn small kind="danger" onClick={() => mutate((d) => { const task = d.tasks.find((x) => x.id === t.id); if (task) { task.reply = ""; task.repliedAt = null; task.replySeen = true; } return d; })}>O‘chirish</Btn>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>

          <h4 style={{ color: C.navy }} className="font-semibold text-sm mb-2 mt-5">Kunlik mavzu nazorati (soat 16:00)</h4>
          {(db.dailyTopicAlerts || []).length === 0 && <p style={{ color: C.textSoft }} className="text-sm">Hozircha bildirishnoma yo‘q.</p>}
          <div className="space-y-2">
            {[...(db.dailyTopicAlerts || [])].sort((a, b) => (b.date || "").localeCompare(a.date || "")).map((a) => (
              <div key={a.id} style={{ borderColor: a.read ? C.line : C.tempBorder, background: a.read ? "#fff" : C.temp }} className="border rounded-md px-3 py-2.5 text-sm">
                <div className="flex items-center justify-between gap-2">
                  <div>
                    <div style={{ color: C.navy }} className="font-semibold">{fmtDate(a.date)}</div>
                    <div style={{ color: C.textSoft }} className="text-xs mt-0.5">{a.items.length} ta darsda mavzu belgilanmagan</div>
                  </div>
                  <div className="flex items-center gap-2 shrink-0">
                    <Btn small kind="ghost" onClick={() => downloadMissingTopicsPDF(a)}><span className="inline-flex items-center gap-1"><FileDown size={14} /> PDF</span></Btn>
                    {!a.read && <Btn small onClick={() => mutate((d) => { const x = d.dailyTopicAlerts.find((y) => y.id === a.id); if (x) x.read = true; return d; })}>Ko‘rdim</Btn>}
                    <Btn small kind="danger" onClick={() => mutate((d) => { d.dailyTopicAlerts = d.dailyTopicAlerts.filter((y) => y.id !== a.id); return d; })}>O‘chirish</Btn>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </Modal>
      )}

      {bsbOpen && (
        <Modal title="BSB/CHSB nazorati" onClose={() => setBsbOpen(false)} wide>
          {(db.bsbNotifications || []).length === 0 && <p style={{ color: C.textSoft }} className="text-sm">Hozircha bildirishnoma yo‘q.</p>}
          <div className="space-y-2">
            {[...(db.bsbNotifications || [])].sort((a, b) => (b.date || "").localeCompare(a.date || "")).map((n) => (
              <div key={n.id} style={{ borderColor: n.read ? C.line : C.tempBorder, background: n.read ? "#fff" : C.temp }} className="border rounded-md px-3 py-2.5 text-sm">
                <div className="flex items-start justify-between gap-2">
                  <div>
                    <div style={{ color: C.navy }} className="font-semibold">{n.teacherName}</div>
                    <div style={{ color: C.textMain }} className="text-xs mt-0.5">{n.className}-sinf · {n.subjectName}</div>
                    <div className="mt-1">
                      <span style={{ background: C.goldDark, color: "#fff" }} className="text-[10px] font-bold px-2 py-0.5 rounded-full">{n.type}</span>
                      <span style={{ color: C.textSoft }} className="text-xs ml-2">{fmtDate(n.date)}</span>
                    </div>
                  </div>
                  <div className="flex gap-1.5 shrink-0">
                    {!n.read && <Btn small onClick={() => mutate((d) => { const x = d.bsbNotifications.find((y) => y.id === n.id); if (x) x.read = true; return d; })}>Ko‘rdim</Btn>}
                    <Btn small kind="danger" onClick={() => mutate((d) => { d.bsbNotifications = d.bsbNotifications.filter((y) => y.id !== n.id); return d; })}>O‘chirish</Btn>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </Modal>
      )}

      <div className="flex flex-wrap gap-2 mb-6">
        {tabs.map(([id, label]) => (
          <button
            key={id}
            onClick={() => setTab(id)}
            style={tab === id ? { background: C.gold, color: "#fff" } : { background: "#fff", color: C.textMain, border: `1px solid ${C.line}` }}
            className="px-3.5 py-1.5 rounded-full text-sm font-medium"
          >
            {label}
          </button>
        ))}
      </div>

      <div style={{ background: C.card, borderColor: C.line }} className="rounded-xl border p-5">
        {tab === "teachers" && (
          <TeachersEditor
            teachers={db.teachers}
            subjects={db.subjects}
            onSave={(t) => mutate((d) => { const i = d.teachers.findIndex((x) => x.id === t.id); if (i >= 0) d.teachers[i] = t; else d.teachers.push(t); return d; })}
            onDelete={(id) => mutate((d) => { d.teachers = d.teachers.filter((x) => x.id !== id); return d; })}
            onImportMany={(list) => mutate((d) => { d.teachers = [...d.teachers, ...list]; return d; })}
          />
        )}
        {tab === "tasks" && (
          <TaskAssignEditor
            teachers={db.teachers}
            tasks={db.tasks || []}
            onAssign={(t) => mutate((d) => { d.tasks = [...(d.tasks || []), t]; return d; })}
            onDelete={(id) => mutate((d) => { d.tasks = d.tasks.filter((x) => x.id !== id); return d; })}
            onSeenReply={(id) => mutate((d) => { const t = d.tasks.find((x) => x.id === id); if (t) t.replySeen = true; return d; })}
          />
        )}
        {tab === "yearlyPlans" && <AdminYearlyPlansEditor db={db} mutate={mutate} />}
        {tab === "subjects" && (
          <div>
            <NamedImportControls
              items={db.subjects}
              label="Fanlar"
              fileHeader="Fan nomi"
              templateFile="fanlar-namuna.xlsx"
              templateSheet="Fanlar"
              templateExamples={["Matematika", "Fizika", "Ona tili"]}
              onImportMany={(names) => mutate((d) => { d.subjects = [...d.subjects, ...names.map((n) => ({ id: uid(), name: n }))]; return d; })}
            />
            <NamedListEditor
              items={db.subjects}
              label="Fanlar"
              placeholder="Yangi fan nomi"
              onAdd={(name) => mutate((d) => { d.subjects.push({ id: uid(), name }); return d; })}
              onRename={(id, name) => mutate((d) => { const s = d.subjects.find((x) => x.id === id); if (s) s.name = name; return d; })}
              onDelete={(id) => mutate((d) => { d.subjects = d.subjects.filter((x) => x.id !== id); return d; })}
            />
          </div>
        )}
        {tab === "days" && (
          <NamedListEditor
            items={db.days}
            label="Hafta kunlari"
            placeholder="Yangi kun nomi"
            onAdd={(name) => mutate((d) => { d.days.push({ id: uid(), name }); return d; })}
            onRename={(id, name) => mutate((d) => { const s = d.days.find((x) => x.id === id); if (s) s.name = name; return d; })}
            onDelete={(id) => mutate((d) => { d.days = d.days.filter((x) => x.id !== id); return d; })}
            onMove={(id, dir) => mutate((d) => { const i = d.days.findIndex((x) => x.id === id); const j = i + dir; if (j < 0 || j >= d.days.length) return d; [d.days[i], d.days[j]] = [d.days[j], d.days[i]]; return d; })}
          />
        )}
        {tab === "slots" && (
          <SlotsEditor
            slots={db.slots}
            onSave={(s) => mutate((d) => { const i = d.slots.findIndex((x) => x.id === s.id); if (i >= 0) d.slots[i] = s; else d.slots.push(s); return d; })}
            onDelete={(id) => mutate((d) => { d.slots = d.slots.filter((x) => x.id !== id); return d; })}
            onMove={(id, dir) => mutate((d) => { const i = d.slots.findIndex((x) => x.id === id); const j = i + dir; if (j < 0 || j >= d.slots.length) return d; [d.slots[i], d.slots[j]] = [d.slots[j], d.slots[i]]; return d; })}
          />
        )}
        {tab === "classes" && (
          <div>
            <NamedImportControls
              items={db.classes}
              label="Sinflar"
              fileHeader="Sinf nomi"
              templateFile="sinflar-namuna.xlsx"
              templateSheet="Sinflar"
              templateExamples={["5-A", "5-B", "6-A"]}
              onImportMany={(names) => mutate((d) => { d.classes = [...d.classes, ...names.map((n) => ({ id: uid(), name: n }))]; return d; })}
            />
            <NamedListEditor
              items={db.classes}
              label="Sinflar"
              placeholder="Masalan: 5-A"
              onAdd={(name) => mutate((d) => { d.classes.push({ id: uid(), name }); return d; })}
              onRename={(id, name) => mutate((d) => { const s = d.classes.find((x) => x.id === id); if (s) s.name = name; return d; })}
              onDelete={(id) => mutate((d) => { d.classes = d.classes.filter((x) => x.id !== id); return d; })}
            />
          </div>
        )}
        {tab === "rooms" && (
          <NamedListEditor
            items={db.rooms}
            label="Xonalar"
            placeholder="Masalan: 204-xona"
            onAdd={(name) => mutate((d) => { d.rooms.push({ id: uid(), name }); return d; })}
            onRename={(id, name) => mutate((d) => { const s = d.rooms.find((x) => x.id === id); if (s) s.name = name; return d; })}
            onDelete={(id) => mutate((d) => { d.rooms = d.rooms.filter((x) => x.id !== id); return d; })}
          />
        )}
      </div>
    </div>
  );
}

/* ---------- yearly lesson plan (teacher side) ---------- */
const UZ_WEEKDAYS = ["Yakshanba", "Dushanba", "Seshanba", "Chorshanba", "Payshanba", "Juma", "Shanba"];
const getTodayDay = (db) => {
  const name = UZ_WEEKDAYS[new Date().getDay()];
  return db.days.find((d) => d.name.trim().toLowerCase() === name.toLowerCase());
};

function computeMissingTopicItems(db) {
  const todayDay = getTodayDay(db);
  if (!todayDay) return [];
  const todayStr = new Date().toISOString().slice(0, 10);
  const seen = new Set();
  const items = [];
  Object.entries(db.schedule || {}).forEach(([classId, byDay]) => {
    const bySlot = byDay[todayDay.id];
    if (!bySlot) return;
    Object.values(bySlot).forEach((cell) => {
      (cell.entries || []).forEach((e) => {
        const key = `${e.teacherId}|${classId}|${e.subjectId}`;
        if (seen.has(key)) return;
        seen.add(key);
        const plan = (db.yearlyPlans || []).find((p) => p.teacherId === e.teacherId && p.classId === classId && p.subjectId === e.subjectId);
        const markedToday = plan && plan.rows.some((r) => r.done && r.doneDate && r.doneDate.slice(0, 10) === todayStr);
        if (!markedToday) {
          const teacher = db.teachers.find((t) => t.id === e.teacherId);
          const cls = db.classes.find((c) => c.id === classId);
          const subj = db.subjects.find((s) => s.id === e.subjectId);
          items.push({
            teacherId: e.teacherId, teacherName: teacher ? `${teacher.firstName} ${teacher.lastName}` : "?",
            classId, className: cls?.name || "?", subjectId: e.subjectId, subjectName: subj?.name || "?",
          });
        }
      });
    });
  });
  return items;
}

function getTeacherClassSubjectPairs(db, teacherId) {
  const map = new Map();
  Object.entries(db.schedule || {}).forEach(([classId, byDay]) => {
    Object.values(byDay).forEach((bySlot) => {
      Object.values(bySlot).forEach((cell) => {
        (cell.entries || []).forEach((e) => {
          if (e.teacherId === teacherId) {
            const key = `${classId}|${e.subjectId}`;
            if (!map.has(key)) map.set(key, { classId, subjectId: e.subjectId });
          }
        });
      });
    });
  });
  return [...map.values()].map((p) => ({
    ...p,
    className: db.classes.find((c) => c.id === p.classId)?.name || "?",
    subjectName: db.subjects.find((s) => s.id === p.subjectId)?.name || "?",
  }));
}

function YearlyPlanView({ db, mutate, teacher, onClose }) {
  const pairs = getTeacherClassSubjectPairs(db, teacher.id);
  const myPlans = (db.yearlyPlans || []).filter((p) => p.teacherId === teacher.id);

  const [mode, setMode] = useState("list"); // list | add | detail
  const [activePlanId, setActivePlanId] = useState(null);
  const [addClassIds, setAddClassIds] = useState([]);
  const [addSubjectId, setAddSubjectId] = useState("");
  const [err, setErr] = useState("");
  const [confirmReimport, setConfirmReimport] = useState(null); // {file, classIds, subjectId} or {file, classId, subjectId} for single reimport
  const [editingRowId, setEditingRowId] = useState(null);
  const [editValue, setEditValue] = useState("");

  const activePlan = myPlans.find((p) => p.id === activePlanId);

  const todayStr = new Date().toISOString().slice(0, 10);
  const doneTodayCount = activePlan ? activePlan.rows.filter((r) => r.done && r.doneDate && r.doneDate.slice(0, 10) === todayStr).length : 0;
  const dailyLimitReached = doneTodayCount >= 2;

  const className = (id) => db.classes.find((c) => c.id === id)?.name || "?";
  const subjectName = (id) => db.subjects.find((s) => s.id === id)?.name || "?";

  const todayDay = getTodayDay(db);
  const planNeedsAttention = (plan) => {
    if (!todayDay) return false;
    const bySlot = db.schedule?.[plan.classId]?.[todayDay.id];
    if (!bySlot) return false;
    const hasLessonToday = Object.values(bySlot).some((cell) => cell.entries?.some((e) => e.teacherId === teacher.id && e.subjectId === plan.subjectId));
    if (!hasLessonToday) return false;
    const markedToday = plan.rows.some((r) => r.done && r.doneDate && r.doneDate.slice(0, 10) === todayStr);
    return !markedToday;
  };

  const parseRowsFromFile = (file, onSuccess) => {
    setErr("");
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const arr = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        const dataRows = arr.slice(1).filter((r) => r.some((c) => String(c).trim() !== ""));
        if (dataRows.length === 0) { setErr("Faylda dars qatorlari topilmadi."); return; }
        const baseRows = dataRows.map((r, i) => {
          const label = String(r[0] ?? "").trim();
          const topicRaw = String(r[1] ?? "").trim();
          const isBSB = /\bBSB\b/i.test(topicRaw);
          const isCHSB = /\bCHSB\b/i.test(topicRaw);
          return {
            order: i + 1, label, topic: topicRaw,
            darsTuri: String(r[2] ?? "").trim(), sinfIshi: String(r[3] ?? "").trim(), uygaVazifa: String(r[4] ?? "").trim(),
            soat: r[5], izoh: String(r[6] ?? "").trim(), chorak: r[7],
            isBSB, isCHSB,
          };
        });
        onSuccess(baseRows);
      } catch {
        setErr("Faylni o‘qib bo‘lmadi. Namunaga mos .xlsx faylni tanlang.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const cloneRows = (baseRows) => baseRows.map((r) => ({ ...r, id: uid(), done: false, doneDate: null }));

  // create/replace plans for one or more classes at once (used by the "add" flow)
  const importForClasses = (file, classIds, subjectId) => {
    parseRowsFromFile(file, (baseRows) => {
      let firstNewId = null;
      mutate((d) => {
        classIds.forEach((classId) => {
          d.yearlyPlans = (d.yearlyPlans || []).filter((p) => !(p.teacherId === teacher.id && p.classId === classId && p.subjectId === subjectId));
          const newId = uid();
          if (!firstNewId) firstNewId = newId;
          d.yearlyPlans.push({ id: newId, teacherId: teacher.id, classId, subjectId, importedAt: new Date().toISOString(), rows: cloneRows(baseRows) });
        });
        return d;
      });
      setConfirmReimport(null);
      setAddClassIds([]);
      setAddSubjectId("");
      if (classIds.length === 1) { setActivePlanId(firstNewId); setMode("detail"); } else { setMode("list"); }
    });
  };

  // replace a single already-open plan (used by "Qayta import" inside a plan)
  const reimportSinglePlan = (file, classId, subjectId) => {
    parseRowsFromFile(file, (baseRows) => {
      let newId = uid();
      mutate((d) => {
        d.yearlyPlans = (d.yearlyPlans || []).filter((p) => !(p.teacherId === teacher.id && p.classId === classId && p.subjectId === subjectId));
        d.yearlyPlans.push({ id: newId, teacherId: teacher.id, classId, subjectId, importedAt: new Date().toISOString(), rows: cloneRows(baseRows) });
        return d;
      });
      setConfirmReimport(null);
      setActivePlanId(newId);
      setMode("detail");
    });
  };

  const toggleAddClass = (classId) => {
    setAddClassIds((ids) => (ids.includes(classId) ? ids.filter((x) => x !== classId) : [...ids, classId]));
  };

  const handleAddFile = (e) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file || addClassIds.length === 0 || !addSubjectId) return;
    const conflicts = addClassIds.filter((id) => myPlans.some((p) => p.classId === id && p.subjectId === addSubjectId));
    if (conflicts.length > 0) { setConfirmReimport({ file, classIds: addClassIds, subjectId: addSubjectId, conflictNames: conflicts.map(className) }); return; }
    importForClasses(file, addClassIds, addSubjectId);
  };

  const handleReimportFile = (plan) => (e) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    setConfirmReimport({ file, classId: plan.classId, subjectId: plan.subjectId });
  };

  const markDone = (row) => {
    if (dailyLimitReached || !activePlan) return;
    mutate((d) => {
      const p = d.yearlyPlans.find((pl) => pl.id === activePlan.id);
      const r = p.rows.find((x) => x.id === row.id);
      r.done = true;
      r.doneDate = new Date().toISOString();
      if (row.isBSB || row.isCHSB) {
        d.bsbNotifications = d.bsbNotifications || [];
        d.bsbNotifications.push({
          id: uid(), teacherId: teacher.id, teacherName: `${teacher.firstName} ${teacher.lastName}`,
          classId: activePlan.classId, className: className(activePlan.classId), subjectId: activePlan.subjectId, subjectName: subjectName(activePlan.subjectId),
          type: row.isBSB ? "BSB" : "CHSB", date: new Date().toISOString(), read: false,
        });
      }
      return d;
    });
  };

  const editRowTopic = (rowId, newTopic) => {
    mutate((d) => {
      const p = d.yearlyPlans.find((pl) => pl.id === activePlan.id);
      const r = p.rows.find((x) => x.id === rowId);
      r.topic = newTopic;
      r.isBSB = /\bBSB\b/i.test(newTopic);
      r.isCHSB = /\bCHSB\b/i.test(newTopic);
      return d;
    });
  };

  /* ---- LIST MODE ---- */
  if (mode === "list") {
    return (
      <div>
        <div className="flex items-center justify-between mb-5">
          <div className="flex items-center gap-3">
            <button onClick={onClose} style={{ color: C.textSoft }} className="flex items-center gap-1 text-sm"><ChevronLeft size={16} /> Orqaga</button>
            <h2 style={{ color: C.navy }} className="font-bold text-lg">Yillik ish rejalarim</h2>
          </div>
          <Btn onClick={() => { setAddClassIds([]); setAddSubjectId(""); setErr(""); setMode("add"); }}>+ Yangi reja qo‘shish</Btn>
        </div>

        {myPlans.length === 0 ? (
          <div style={{ borderColor: C.line }} className="border border-dashed rounded-xl p-8 text-center">
            <Upload size={28} className="mx-auto mb-2" style={{ color: C.textSoft }} />
            <p style={{ color: C.textMain }} className="text-sm">Hali birorta yillik ish rejasi import qilinmagan.</p>
          </div>
        ) : (
          <div className="grid sm:grid-cols-2 gap-3">
            {myPlans.map((p) => {
              const doneCount = p.rows.filter((r) => r.done).length;
              const needsAttention = planNeedsAttention(p);
              return (
                <button
                  key={p.id}
                  onClick={() => { setActivePlanId(p.id); setMode("detail"); }}
                  style={{ background: C.card, borderColor: needsAttention ? C.gold : C.line }}
                  className={`text-left border rounded-xl p-4 hover:shadow-sm transition ${needsAttention ? "plan-blink" : ""}`}
                >
                  <div style={{ color: C.navy }} className="font-semibold">{className(p.classId)}-sinf · {subjectName(p.subjectId)}</div>
                  <div style={{ color: C.textSoft }} className="text-xs mt-1">{doneCount}/{p.rows.length} dars bajarildi</div>
                  <div style={{ color: C.textSoft }} className="text-[11px] mt-1">Import: {fmtDate(p.importedAt)}</div>
                  {needsAttention && (
                    <div style={{ color: C.goldDark }} className="flex items-center gap-1 text-[11px] font-semibold mt-2">
                      <AlertTriangle size={12} strokeWidth={2} /> Bugun darsingiz bor — mavzuni belgilang
                    </div>
                  )}
                </button>
              );
            })}
          </div>
        )}
      </div>
    );
  }

  /* ---- ADD MODE ---- */
  if (mode === "add") {
    return (
      <div>
        <div className="flex items-center gap-3 mb-5">
          <button onClick={() => setMode("list")} style={{ color: C.textSoft }} className="flex items-center gap-1 text-sm"><ChevronLeft size={16} /> Orqaga</button>
          <h2 style={{ color: C.navy }} className="font-bold text-lg">Yangi ish rejasi qo‘shish</h2>
        </div>
        <div style={{ background: C.card, borderColor: C.line }} className="border rounded-xl p-4 mb-5">
          <Field label="Fan">
            <select className={inputCls} style={inputStyle} value={addSubjectId} onChange={(e) => setAddSubjectId(e.target.value)}>
              <option value="">Tanlang</option>
              {db.subjects.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
          </Field>
          <div className="mt-3">
            <span style={{ color: C.textSoft }} className="block text-xs font-medium mb-1.5 uppercase tracking-wide">Sinf(lar) — parallel sinflar uchun bir nechtasini belgilashingiz mumkin</span>
            <div className="flex flex-wrap gap-2">
              {db.classes.map((c) => {
                const active = addClassIds.includes(c.id);
                return (
                  <button
                    key={c.id}
                    type="button"
                    onClick={() => toggleAddClass(c.id)}
                    style={active ? { background: C.gold, color: "#fff", borderColor: C.gold } : { borderColor: C.line, color: C.textMain }}
                    className="text-xs px-2.5 py-1 rounded-full border"
                  >
                    {c.name}
                  </button>
                );
              })}
              {db.classes.length === 0 && <span style={{ color: C.textSoft }} className="text-xs">Avval sinf qo‘shing.</span>}
            </div>
          </div>
          {pairs.length > 0 && (
            <div className="flex flex-wrap gap-2 mt-3">
              <span style={{ color: C.textSoft }} className="text-xs self-center mr-1">Tezkor tanlash:</span>
              {pairs.map((p) => (
                <button
                  key={`${p.classId}-${p.subjectId}`}
                  onClick={() => { toggleAddClass(p.classId); setAddSubjectId(p.subjectId); }}
                  style={addClassIds.includes(p.classId) && addSubjectId === p.subjectId ? { background: C.gold, color: "#fff", borderColor: C.gold } : { borderColor: C.line, color: C.textMain }}
                  className="text-xs px-2.5 py-1 rounded-full border"
                >
                  {p.className} · {p.subjectName}
                </button>
              ))}
            </div>
          )}
        </div>

        {addClassIds.length === 0 || !addSubjectId ? (
          <p style={{ color: C.textSoft }} className="text-sm">Davom etish uchun kamida bitta sinf va fanni tanlang.</p>
        ) : (
          <div style={{ borderColor: C.line }} className="border border-dashed rounded-xl p-8 text-center">
            <Upload size={28} className="mx-auto mb-2" style={{ color: C.textSoft }} />
            <p style={{ color: C.textMain }} className="text-sm mb-3">
              {addClassIds.map(className).join(", ")}-sinf(lar) · {subjectName(addSubjectId)} uchun bitta .xlsx faylni tanlang — barcha tanlangan sinflarga alohida-alohida saqlanadi.
            </p>
            <label style={{ background: C.teal, color: "#fff" }} className="inline-block rounded-md font-medium px-4 py-2 text-sm hover:opacity-90 transition cursor-pointer">
              📥 .xlsx faylni import qilish
              <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleAddFile} />
            </label>
            {err && <p style={{ color: C.red }} className="text-xs mt-3">{err}</p>}
          </div>
        )}

        {confirmReimport && confirmReimport.classIds && (
          <Modal title="Rejani almashtirish" onClose={() => setConfirmReimport(null)}>
            <p style={{ color: C.textMain }} className="text-sm mb-4">
              Quyidagi sinflar uchun bu fandan reja allaqachon mavjud: <b>{confirmReimport.conflictNames.join(", ")}</b>. Davom etilsa, ularning avvalgi reja va belgilari yangi fayl bilan almashtiriladi. Boshqa tanlangan sinflar uchun yangi reja qo‘shiladi. Davom etasizmi?
            </p>
            <div className="flex justify-end gap-2">
              <Btn kind="ghost" onClick={() => setConfirmReimport(null)}>Bekor qilish</Btn>
              <Btn kind="danger" onClick={() => importForClasses(confirmReimport.file, confirmReimport.classIds, confirmReimport.subjectId)}>Ha, davom etish</Btn>
            </div>
          </Modal>
        )}
      </div>
    );
  }

  /* ---- DETAIL MODE ---- */
  if (!activePlan) { setMode("list"); return null; }
  return (
    <div>
      <div className="flex items-center justify-between mb-5">
        <div className="flex items-center gap-3">
          <button onClick={() => setMode("list")} style={{ color: C.textSoft }} className="flex items-center gap-1 text-sm"><ChevronLeft size={16} /> Rejalarim</button>
          <h2 style={{ color: C.navy }} className="font-bold text-lg">{className(activePlan.classId)}-sinf · {subjectName(activePlan.subjectId)}</h2>
        </div>
        <label style={{ borderColor: C.line, color: C.textMain }} className="border rounded-md px-3 py-1.5 text-xs cursor-pointer hover:bg-gray-50">
          🔄 Qayta import qilish
          <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleReimportFile(activePlan)} />
        </label>
      </div>
      {err && <p style={{ color: C.red }} className="text-xs mb-3">{err}</p>}
      <p style={{ color: dailyLimitReached ? C.goldDark : C.textSoft }} className="text-xs mb-3">
        Bugun belgilangan mavzular: {doneTodayCount}/2{dailyLimitReached ? " — bugungi limit tugadi." : ""}
      </p>
      <div className="space-y-2">
        {activePlan.rows.map((row) => {
          const special = row.isBSB || row.isCHSB;
          const isEditing = editingRowId === row.id;
          return (
            <div
              key={row.id}
              style={{ borderColor: row.done ? "#8FBF8A" : special ? C.tempBorder : C.line, background: row.done ? "#EAF6E9" : special ? C.temp : "#fff" }}
              className="border rounded-md px-3 py-2.5 text-sm"
            >
              <div className="flex items-start justify-between gap-2 mb-1.5">
                <span style={{ color: C.darkBlue }} className="text-sm font-semibold">{row.label}{row.chorak ? ` · ${row.chorak}-chorak` : ""}</span>
                <div className="flex items-center gap-2">
                  {special && <span style={{ background: C.goldDark, color: "#fff" }} className="text-[10px] font-bold px-2 py-0.5 rounded-full">{row.isBSB ? "BSB" : "CHSB"}</span>}
                  {row.done && <span style={{ color: C.teal }} className="text-xs font-medium flex items-center gap-1"><CheckCircle2 size={14} /> Bajarildi</span>}
                  {!isEditing && (
                    <button title="Mavzuni tahrirlash" onClick={() => { setEditingRowId(row.id); setEditValue(row.topic); }} style={{ color: C.textSoft }} className="hover:text-black">
                      <Pencil size={14} />
                    </button>
                  )}
                </div>
              </div>

              {isEditing ? (
                <div className="mb-2">
                  <textarea className={inputCls} style={inputStyle} rows={3} value={editValue} onChange={(e) => setEditValue(e.target.value)} />
                  <div className="flex justify-end gap-2 mt-2">
                    <Btn small kind="ghost" onClick={() => setEditingRowId(null)}>Bekor qilish</Btn>
                    <Btn small onClick={() => { editRowTopic(row.id, editValue.trim()); setEditingRowId(null); }}>Saqlash</Btn>
                  </div>
                </div>
              ) : (
                <p style={{ color: C.textMain }} className="whitespace-pre-wrap mb-2">{row.topic}</p>
              )}

              {!isEditing && (!row.done ? (
                <div className="flex justify-end">
                  <Btn small onClick={() => markDone(row)}>{dailyLimitReached ? "Bugungi limit tugadi" : "Bugun bajarildi deb belgilash"}</Btn>
                </div>
              ) : (
                <div style={{ color: C.darkBlue }} className="text-sm font-semibold">{fmtDate(row.doneDate)}</div>
              ))}
            </div>
          );
        })}
      </div>

      {confirmReimport && confirmReimport.classId && (
        <Modal title="Rejani almashtirish" onClose={() => setConfirmReimport(null)}>
          <p style={{ color: C.textMain }} className="text-sm mb-4">Ushbu sinf va fan uchun mavjud reja va bajarilgan belgilar butunlay yangi fayl bilan almashtiriladi. Davom etasizmi?</p>
          <div className="flex justify-end gap-2">
            <Btn kind="ghost" onClick={() => setConfirmReimport(null)}>Bekor qilish</Btn>
            <Btn kind="danger" onClick={() => reimportSinglePlan(confirmReimport.file, confirmReimport.classId, confirmReimport.subjectId)}>Ha, almashtirish</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ---------- teacher cabinet ---------- */
function TeacherCabinet({ db, teacher, mutate, onLogout }) {
  const [absenceOpen, setAbsenceOpen] = useState(false);
  const [tasksOpen, setTasksOpen] = useState(false);
  const [planOpen, setPlanOpen] = useState(false);
  const [openTaskId, setOpenTaskId] = useState(null);
  const [replyDraft, setReplyDraft] = useState({});
  const [editingReplyId, setEditingReplyId] = useState(null);
  const myTasks = [...(db.tasks || [])].filter((t) => t.teacherId === teacher.id).sort((a, b) => (b.createdAt || "").localeCompare(a.createdAt || ""));
  const unreadTaskCount = myTasks.filter((t) => !t.read).length;
  const myRequests = [...(db.absenceRequests || [])].filter((r) => r.teacherId === teacher.id).sort((a, b) => (b.createdAt || "").localeCompare(a.createdAt || ""));

  if (planOpen) return <YearlyPlanView db={db} mutate={mutate} teacher={teacher} onClose={() => setPlanOpen(false)} />;

  const rows = [];
  db.days.forEach((day) => {
    db.slots.filter((s) => s.type === "dars").forEach((slot) => {
      db.classes.forEach((cls) => {
        const cell = db.schedule?.[cls.id]?.[day.id]?.[slot.id];
        if (!cell) return;
        cell.entries.forEach((e) => {
          if (e.teacherId === teacher.id) {
            rows.push({
              day: day.name,
              slot: slot.name,
              time: `${slot.start}–${slot.end}`,
              cls: cls.name,
              subject: db.subjects.find((s) => s.id === e.subjectId)?.name || "?",
              room: db.rooms.find((r) => r.id === e.roomId)?.name || "?",
              comment: e.comment,
              hasTemp: cell.tempEdits?.length > 0,
            });
          }
        });
      });
    });
  });

  const grouped = db.days.map((d) => ({ day: d.name, items: rows.filter((r) => r.day === d.name) }));

  return (
    <div>
      <div className="flex items-center justify-between mb-6">
        <div>
          <h1 style={{ color: C.navy }} className="text-xl font-bold">Xush kelibsiz, {teacher.firstName} {teacher.lastName}</h1>
          <p style={{ color: C.textSoft }} className="text-sm">Shaxsiy dars jadvalingiz</p>
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={() => setTasksOpen(true)}
            title="Vazifalar"
            style={{ background: C.navy, color: "#fff" }}
            className={`w-11 h-11 rounded-lg flex items-center justify-center hover:opacity-90 relative ${unreadTaskCount > 0 ? "bell-shake" : ""}`}
          >
            <Bell size={20} strokeWidth={1.8} />
            {unreadTaskCount > 0 && (
              <span style={{ background: C.red }} className="absolute -top-1.5 -right-1.5 text-white text-[10px] font-semibold w-5 h-5 rounded-full flex items-center justify-center">{unreadTaskCount}</span>
            )}
          </button>
          <Btn kind="ghost" onClick={() => setPlanOpen(true)}><span className="inline-flex items-center gap-1.5"><BookOpen size={15} strokeWidth={2} /> Yillik reja</span></Btn>
          <Btn kind="dark" onClick={() => setAbsenceOpen(true)}><span className="inline-flex items-center gap-1.5"><CalendarOff size={15} strokeWidth={2} /> ISHDA BO‘LMAYMAN</span></Btn>
          <Btn kind="ghost" onClick={onLogout}>Chiqish</Btn>
        </div>
      </div>

      {tasksOpen && (
        <Modal title="Vazifalar" onClose={() => setTasksOpen(false)} wide>
          {myTasks.length === 0 && <p style={{ color: C.textSoft }} className="text-sm">Sizga hali vazifa berilmagan.</p>}
          <div className="space-y-2">
            {myTasks.map((t) => {
              const isOpen = openTaskId === t.id;
              return (
                <div key={t.id} style={{ borderColor: t.read ? C.line : C.tempBorder, background: t.read ? "#fff" : C.temp }} className="border rounded-md overflow-hidden">
                  <button
                    className="w-full text-left px-3 py-2.5 flex items-center justify-between gap-2"
                    onClick={() => {
                      setOpenTaskId(isOpen ? null : t.id);
                      if (!t.read) mutate((d) => { const task = d.tasks.find((x) => x.id === t.id); if (task) task.read = true; return d; });
                    }}
                  >
                    <span style={{ color: C.navy }} className="font-semibold text-sm">{!t.read && <span style={{ color: C.red }} className="mr-1">●</span>}{t.title}</span>
                    <span style={{ color: C.textSoft }} className="text-xs shrink-0">{isOpen ? "▲" : "▼"}</span>
                  </button>
                  {isOpen && (
                    <div className="px-3 pb-3">
                      <p style={{ color: C.textMain }} className="text-sm whitespace-pre-wrap">{t.description || "Tavsif kiritilmagan."}</p>
                      <p style={{ color: C.textSoft }} className="text-xs mt-2 mb-3">{fmtDate(t.createdAt)}</p>

                      {t.reply && editingReplyId !== t.id && (
                        <div style={{ borderColor: C.line, background: "#fff" }} className="border rounded-md px-2.5 py-2 mb-2">
                          <div className="flex items-start justify-between gap-2">
                            <div style={{ color: C.textSoft }} className="text-[11px] uppercase font-medium mb-0.5">Sizning javobingiz</div>
                            <button onClick={() => { setEditingReplyId(t.id); setReplyDraft((s) => ({ ...s, [t.id]: t.reply })); }} style={{ color: C.textSoft }} className="hover:text-black shrink-0">
                              <Pencil size={13} />
                            </button>
                          </div>
                          <div style={{ color: C.textMain }} className="text-sm whitespace-pre-wrap">{t.reply}</div>
                        </div>
                      )}

                      {(!t.reply || editingReplyId === t.id) && (
                        <>
                          <textarea
                            className={inputCls}
                            style={inputStyle}
                            rows={2}
                            placeholder="Adminga javob yozing..."
                            value={replyDraft[t.id] ?? ""}
                            onChange={(e) => setReplyDraft((s) => ({ ...s, [t.id]: e.target.value }))}
                          />
                          <div className="flex justify-end gap-2 mt-2">
                            {editingReplyId === t.id && (
                              <Btn small kind="ghost" onClick={() => setEditingReplyId(null)}>Bekor qilish</Btn>
                            )}
                            <Btn
                              small
                              onClick={() => {
                                const text = (replyDraft[t.id] ?? "").trim();
                                if (!text) return;
                                mutate((d) => {
                                  const task = d.tasks.find((x) => x.id === t.id);
                                  if (task) { task.reply = text; task.repliedAt = new Date().toISOString(); task.replySeen = false; }
                                  return d;
                                });
                                setReplyDraft((s) => { const c = { ...s }; delete c[t.id]; return c; });
                                setEditingReplyId(null);
                              }}
                            >
                              {t.reply ? "Javobni yangilash" : "Javobni yuborish"}
                            </Btn>
                          </div>
                        </>
                      )}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </Modal>
      )}

      {myRequests.length > 0 && (
        <div style={{ background: C.card, borderColor: C.line }} className="rounded-xl border p-4 mb-4">
          <h3 style={{ color: C.navy }} className="font-semibold mb-2 text-sm">Mening so‘rovlarim</h3>
          <div className="space-y-2">
            {myRequests.map((r) => (
              <div key={r.id} style={{ borderColor: C.line }} className="border rounded-md px-3 py-2 flex items-center justify-between text-sm">
                <div>
                  <span style={{ color: C.textMain }} className="font-medium">{r.startDate} — {r.endDate}</span>
                  <span style={{ color: C.textSoft }} className="ml-2 italic">{r.reason}</span>
                </div>
                <span style={{ color: r.read ? C.teal : C.goldDark }} className="text-xs font-medium shrink-0 ml-2">{r.read ? "✓ Ko‘rib chiqildi" : "⏳ Yuborildi"}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      {absenceOpen && (
        <AbsenceModal
          onClose={() => setAbsenceOpen(false)}
          onSubmit={({ startDate, endDate, reason }) => {
            mutate((d) => {
              d.absenceRequests = d.absenceRequests || [];
              d.absenceRequests.push({ id: uid(), teacherId: teacher.id, teacherName: `${teacher.firstName} ${teacher.lastName}`, startDate, endDate, reason, createdAt: new Date().toISOString(), read: false });
              return d;
            });
            setAbsenceOpen(false);
          }}
        />
      )}

      {rows.length === 0 && <p style={{ color: C.textSoft }} className="text-sm">Sizga hali dars biriktirilmagan.</p>}
      <div className="space-y-4">
        {grouped.filter((g) => g.items.length > 0).map((g) => (
          <div key={g.day} style={{ background: C.card, borderColor: C.line }} className="rounded-xl border p-4">
            <h3 style={{ color: C.navy }} className="font-semibold mb-2">{g.day}</h3>
            <div className="space-y-2">
              {g.items.map((r, i) => (
                <div key={i} style={{ borderColor: r.hasTemp ? C.tempBorder : C.line, background: r.hasTemp ? C.temp : "#fff" }} className="border rounded-md px-3 py-2 flex flex-wrap items-center gap-x-4 gap-y-1 text-sm">
                  <span style={{ color: C.textSoft }} className="w-28">{r.time}</span>
                  <span style={{ color: C.navy }} className="font-semibold">{r.subject}</span>
                  <span style={{ color: C.textMain }}>{r.cls}-sinf</span>
                  <span style={{ color: C.textSoft }}>{r.room}</span>
                  {r.comment && <span style={{ color: C.teal }} className="italic">💬 {r.comment}</span>}
                  {r.hasTemp && <span style={{ color: C.goldDark }} className="text-xs font-medium">⏱ vaqtinchalik o‘zgarish mavjud</span>}
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ---------- decorative education banner ---------- */
function EducationPattern() {
  const items = [
    { Icon: BookOpen, top: "12%", left: "6%", size: 46, rotate: -18, opacity: 0.16 },
    { Icon: Pencil, top: "62%", left: "10%", size: 34, rotate: 25, opacity: 0.14 },
    { Icon: GraduationCap, top: "20%", left: "88%", size: 54, rotate: 10, opacity: 0.18 },
    { Icon: Ruler, top: "70%", left: "82%", size: 36, rotate: -20, opacity: 0.14 },
    { Icon: Calculator, top: "8%", left: "42%", size: 30, rotate: 8, opacity: 0.12 },
    { Icon: School, top: "68%", left: "45%", size: 40, rotate: 0, opacity: 0.13 },
    { Icon: BookOpen, top: "35%", left: "20%", size: 26, rotate: 12, opacity: 0.1 },
    { Icon: Pencil, top: "18%", left: "70%", size: 24, rotate: -10, opacity: 0.1 },
  ];
  return (
    <div className="absolute inset-0 overflow-hidden pointer-events-none">
      {items.map((it, i) => {
        const Icon = it.Icon;
        return (
          <Icon
            key={i}
            size={it.size}
            strokeWidth={1.2}
            style={{ position: "absolute", top: it.top, left: it.left, transform: `translate(-50%,-50%) rotate(${it.rotate}deg)`, opacity: it.opacity, color: "#fff" }}
          />
        );
      })}
    </div>
  );
}

/* ---------- landing / auth ---------- */
function Landing({ onAdmin, onTeacher }) {
  return (
    <div>
      <div style={{ background: `linear-gradient(135deg, ${C.navy} 0%, ${C.navy2} 100%)` }} className="relative overflow-hidden rounded-2xl px-6 py-14 mb-10 text-center">
        <EducationPattern />
        <div className="relative z-10">
          <div style={{ background: "rgba(255,255,255,0.12)" }} className="inline-flex w-16 h-16 rounded-2xl items-center justify-center mb-4">
            <GraduationCap size={32} strokeWidth={1.6} className="text-white" />
          </div>
          <h1 className="text-white text-3xl font-bold mb-1">Interaktiv dars jadvali</h1>
          <p style={{ color: "rgba(255,255,255,0.75)" }}>Tizimga kirish usulini tanlang</p>
        </div>
      </div>
      <div className="flex flex-col sm:flex-row gap-4 justify-center">
        <button onClick={onAdmin} style={{ background: C.navy }} className="text-white rounded-xl px-8 py-6 w-64 hover:opacity-90 transition mx-auto sm:mx-0">
          <Settings size={34} strokeWidth={1.6} className="mx-auto mb-2" />
          <div className="font-semibold">Admin</div>
          <div className="text-xs opacity-70 mt-1">Jadval va bazani boshqarish</div>
        </button>
        <button onClick={onTeacher} style={{ background: C.gold }} className="text-white rounded-xl px-8 py-6 w-64 hover:opacity-90 transition mx-auto sm:mx-0">
          <GraduationCap size={34} strokeWidth={1.6} className="mx-auto mb-2" />
          <div className="font-semibold">O‘qituvchi</div>
          <div className="text-xs opacity-70 mt-1">Shaxsiy kabinetga kirish</div>
        </button>
      </div>
    </div>
  );
}

function AdminLogin({ db, onSuccess, onBack }) {
  const [pw, setPw] = useState("");
  const [err, setErr] = useState("");
  return (
    <div className="min-h-[70vh] flex items-center justify-center px-4">
      <div style={{ background: C.card, borderColor: C.line }} className="border rounded-xl p-6 w-full max-w-sm">
        <h2 style={{ color: C.navy }} className="font-semibold text-lg mb-4">Admin sifatida kirish</h2>
        <Field label="Parol">
          <input type="password" className={inputCls} style={inputStyle} value={pw} onChange={(e) => setPw(e.target.value)} onKeyDown={(e) => e.key === "Enter" && (pw === db.adminPassword ? onSuccess() : setErr("Parol noto‘g‘ri"))} autoFocus />
        </Field>
        {err && <p style={{ color: C.red }} className="text-xs mb-2">{err}</p>}
        <p style={{ color: C.textSoft }} className="text-xs mb-4">Standart parol: 1234</p>
        <div className="flex justify-end gap-2">
          <Btn kind="ghost" onClick={onBack}>Orqaga</Btn>
          <Btn onClick={() => (pw === db.adminPassword ? onSuccess() : setErr("Parol noto‘g‘ri"))}>Kirish</Btn>
        </div>
      </div>
    </div>
  );
}

function TeacherLogin({ db, onSuccess, onBack }) {
  const [username, setUsername] = useState("");
  const [pw, setPw] = useState("");
  const [err, setErr] = useState("");

  const submit = () => {
    const t = db.teachers.find((x) => x.username === username && x.password === pw);
    if (t) onSuccess(t);
    else setErr("Login yoki parol noto‘g‘ri");
  };

  return (
    <div className="min-h-[70vh] flex items-center justify-center px-4">
      <div style={{ background: C.card, borderColor: C.line }} className="border rounded-xl p-6 w-full max-w-sm">
        <h2 style={{ color: C.navy }} className="font-semibold text-lg mb-4">O‘qituvchi sifatida kirish</h2>
        <Field label="Login"><input className={inputCls} style={inputStyle} value={username} onChange={(e) => setUsername(e.target.value)} autoFocus /></Field>
        <Field label="Parol"><input type="password" className={inputCls} style={inputStyle} value={pw} onChange={(e) => setPw(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submit()} /></Field>
        {err && <p style={{ color: C.red }} className="text-xs mb-2">{err}</p>}
        <div className="flex justify-end gap-2 mt-2">
          <Btn kind="ghost" onClick={onBack}>Orqaga</Btn>
          <Btn onClick={submit}>Kirish</Btn>
        </div>
      </div>
    </div>
  );
}

/* ---------- root app ---------- */
export default function App() {
  const [db, setDb] = useState(null);
  const [loading, setLoading] = useState(true);
  const [view, setView] = useState("landing"); // landing | adminLogin | admin | schedule | teacherLogin | teacherCabinet
  const [teacher, setTeacher] = useState(null);
  const [saveErr, setSaveErr] = useState(false);

  useEffect(() => {
    (async () => {
      try {
        const res = await storageGet(STORAGE_KEY);
        setDb(res && res.value ? JSON.parse(res.value) : defaultDb());
      } catch {
        try {
          const fresh = defaultDb();
          await storageSet(STORAGE_KEY, JSON.stringify(fresh));
          setDb(fresh);
        } catch {
          setDb(defaultDb());
        }
      } finally {
        setLoading(false);
      }
    })();
  }, []);

  const mutate = useCallback((fn) => {
    setDb((prev) => {
      const next = fn(clone(prev));
      storageSet(STORAGE_KEY, JSON.stringify(next)).catch(() => setSaveErr(true));
      return next;
    });
  }, []);

  const dbRef = useRef(db);
  useEffect(() => { dbRef.current = db; }, [db]);

  useEffect(() => {
    const checkDailyAlert = () => {
      const currentDb = dbRef.current;
      if (!currentDb) return;
      const now = new Date();
      if (now.getHours() < 16) return;
      const todayStr = now.toISOString().slice(0, 10);
      if ((currentDb.dailyTopicAlerts || []).some((a) => a.date === todayStr)) return;
      const items = computeMissingTopicItems(currentDb);
      if (items.length === 0) return;
      mutate((d) => {
        d.dailyTopicAlerts = d.dailyTopicAlerts || [];
        if (d.dailyTopicAlerts.some((a) => a.date === todayStr)) return d;
        d.dailyTopicAlerts.push({ id: uid(), date: todayStr, items, createdAt: now.toISOString(), read: false });
        return d;
      });
    };
    checkDailyAlert();
    const interval = setInterval(checkDailyAlert, 60000);
    return () => clearInterval(interval);
  }, [mutate]);

  if (loading || !db) {
    return <div style={{ background: C.bg, minHeight: "100vh" }} className="flex items-center justify-center"><p style={{ color: C.textSoft }}>Yuklanmoqda…</p></div>;
  }

  return (
    <div style={{ background: C.bg, minHeight: "100vh", fontFamily: "system-ui, -apple-system, Segoe UI, sans-serif" }}>
      <style>{`
        @keyframes bellShake {
          0%, 55%, 100% { transform: rotate(0deg); }
          58% { transform: rotate(-16deg); }
          61% { transform: rotate(14deg); }
          64% { transform: rotate(-10deg); }
          67% { transform: rotate(8deg); }
          70% { transform: rotate(-4deg); }
          73%, 100% { transform: rotate(0deg); }
        }
        .bell-shake { animation: bellShake 2.2s ease-in-out infinite; transform-origin: 50% 20%; }
        @keyframes cellBlink {
          0%, 100% { background-color: #ffffff; }
          50% { background-color: #FBDAD5; }
        }
        .cell-blink { animation: cellBlink 1.4s ease-in-out infinite; }
        @keyframes planBlink {
          0%, 100% { background-color: #ffffff; }
          50% { background-color: #FCE9C9; }
        }
        .plan-blink { animation: planBlink 1.6s ease-in-out infinite; }
      `}</style>
      <div className="max-w-6xl mx-auto px-4 py-6">
        {saveErr && <div style={{ background: "#FDECEA", color: C.red }} className="text-xs rounded-md px-3 py-2 mb-4">Ma’lumotlarni saqlashda xatolik yuz berdi. Ulanishni tekshiring.</div>}
        <p style={{ color: C.textSoft }} className="text-xs mb-3 text-right">Bu ma’lumotlar barcha foydalanuvchilar bilan umumiy saqlanadi.</p>

        {view === "landing" && <Landing onAdmin={() => setView("adminLogin")} onTeacher={() => setView("teacherLogin")} />}
        {view === "adminLogin" && <AdminLogin db={db} onSuccess={() => setView("admin")} onBack={() => setView("landing")} />}
        {view === "teacherLogin" && <TeacherLogin db={db} onSuccess={(t) => { setTeacher(t); setView("teacherCabinet"); }} onBack={() => setView("landing")} />}
        {view === "teacherCabinet" && teacher && <TeacherCabinet db={db} teacher={teacher} mutate={mutate} onLogout={() => { setTeacher(null); setView("landing"); }} />}
        {view === "admin" && <AdminPanel db={db} mutate={mutate} onOpenSchedule={() => setView("schedule")} onLogout={() => setView("landing")} />}
        {view === "schedule" && <ScheduleView db={db} mutate={mutate} onClose={() => setView("admin")} />}
      </div>
    </div>
  );
}