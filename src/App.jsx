import React, { useState, useMemo, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { storageGet, storageSet } from "./storage";
import {
  Waves, Check, X, RotateCcw, Clock, Search, Plus, ChevronLeft, ChevronRight,
  Star, CreditCard, LayoutDashboard, Users, Calendar as CalendarIcon,
  UserPlus, Trash2, Pencil, ArrowLeft, AlertTriangle, Copy, Archive,
  ArchiveRestore, ClipboardList, ShieldCheck, ChevronDown, StickyNote, Upload, FileSpreadsheet
} from "lucide-react";

/* ============================== COSTANTI DI DOMINIO ============================== */

const INK = "#0E2A3A";
const INK_SOFT = "#4A6373";
const SURFACE = "#FFFFFF";
const BG = "#F2F6F7";
const BORDER = "#DCE6E8";
const PRIMARY = "#0E5C73";
const PRIMARY_DARK = "#0A3F4F";
const ACCENT = "#2AA9BF";
const PRESENTE = "#1E9E6B";
const ASSENTE = "#D6584A";
const RECUPERO = "#7B5CC2";
const RITARDO = "#D69A2D";
const FESTIVO_BG = "#EDEFF0";

const DAY_INDEX = { Lun: 1, Mar: 2, Mer: 3, Gio: 4, Ven: 5, Sab: 6 };
const WEEKDAY_LABEL = ["Dom", "Lun", "Mar", "Mer", "Gio", "Ven", "Sab"];

const SECTIONS = [
  { key: "lg", label: "Lunedì + Giovedì", days: ["Lun", "Gio"] },
  { key: "mv", label: "Martedì + Venerdì", days: ["Mar", "Ven"] },
  { key: "m", label: "Mercoledì", days: ["Mer"] },
  { key: "s", label: "Sabato", days: ["Sab"] },
];

const SEASON_MONTHS = [
  { y: 2026, m: 8, label: "Settembre" },
  { y: 2026, m: 9, label: "Ottobre" },
  { y: 2026, m: 10, label: "Novembre" },
  { y: 2026, m: 11, label: "Dicembre" },
  { y: 2027, m: 0, label: "Gennaio" },
  { y: 2027, m: 1, label: "Febbraio" },
  { y: 2027, m: 2, label: "Marzo" },
  { y: 2027, m: 3, label: "Aprile" },
  { y: 2027, m: 4, label: "Maggio" },
  { y: 2027, m: 5, label: "Giugno" },
];

const HOLIDAYS = new Set([
  "2026-11-01", "2026-12-08", "2026-12-25", "2026-12-26",
  "2027-01-01", "2027-01-06", "2027-03-29", "2027-04-25",
  "2027-05-01", "2027-06-02",
]);

// Data di riferimento del prototipo (in produzione si userebbe la data reale del dispositivo)
const TODAY = new Date(2026, 10, 10); // martedì 10 novembre 2026

const CYCLE_ADMIN = ["", "P", "A", "R", "L"];
const CYCLE_SEGRETERIA = ["", "P", "A", "L"];

/* ============================== UTILITY DATE ============================== */

function pad(n) { return n < 10 ? "0" + n : "" + n; }
function toKey(d) { return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; }
function fmtShort(d) { return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}`; }
function weekdayOf(d) { return WEEKDAY_LABEL[d.getDay()]; }
function isHolidayDate(d) { return HOLIDAYS.has(toKey(d)); }
function sameDay(a, b) { return toKey(a) === toKey(b); }

function getLessonDatesForMonth(year, month, dayLabels) {
  const dates = [];
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  for (let day = 1; day <= daysInMonth; day++) {
    const d = new Date(year, month, day);
    const wd = d.getDay();
    for (const label of dayLabels) {
      if (DAY_INDEX[label] === wd) { dates.push(d); break; }
    }
  }
  return dates;
}

function getSeasonLessonDates(dayLabels) {
  let all = [];
  for (const sm of SEASON_MONTHS) all = all.concat(getLessonDatesForMonth(sm.y, sm.m, dayLabels));
  return all;
}

function calcAge(dobStr) {
  const dob = new Date(dobStr);
  let age = TODAY.getFullYear() - dob.getFullYear();
  const m = TODAY.getMonth() - dob.getMonth();
  if (m < 0 || (m === 0 && TODAY.getDate() < dob.getDate())) age--;
  return age;
}

function daysUntil(dateStr) {
  const d = new Date(dateStr);
  return Math.round((d.getTime() - TODAY.getTime()) / 86400000);
}

/* ============================== IMPORT EXCEL ============================== */

const EXCEL_HEADER_MAP = {
  nome: ["nome"],
  cognome: ["cognome"],
  dataNascita: ["data di nascita", "data nascita", "datanascita", "nascita", "data"],
  livello: ["livello"],
  stato: ["stato"],
  certificato: ["certificato", "scadenza certificato", "certificato scadenza", "scadenza"],
  note: ["note", "nota"],
};

function normalizeHeader(h) { return String(h || "").trim().toLowerCase(); }

function mapExcelRow(rawRow) {
  const out = {};
  Object.keys(rawRow).forEach((key) => {
    const nk = normalizeHeader(key);
    for (const [field, variants] of Object.entries(EXCEL_HEADER_MAP)) {
      if (variants.includes(nk)) { out[field] = rawRow[key]; break; }
    }
  });
  return out;
}

function excelDateToKey(v) {
  if (!v) return "";
  if (v instanceof Date && !isNaN(v.getTime())) return toKey(v);
  const s = String(v).trim();
  const m1 = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (m1) {
    let [, d, mo, y] = m1;
    if (y.length === 2) y = "20" + y;
    return `${y}-${pad(parseInt(mo, 10))}-${pad(parseInt(d, 10))}`;
  }
  const m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m2) return s;
  return "";
}

function parseStudentsWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        const sheets = wb.SheetNames.map((name) => {
          const ws = wb.Sheets[name];
          const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
          const rows = json.map(mapExcelRow).filter((r) => (r.nome && String(r.nome).trim()) || (r.cognome && String(r.cognome).trim()));
          return { sheetName: name, rows };
        }).filter((s) => s.rows.length > 0);
        resolve(sheets);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("Errore nella lettura del file"));
    reader.readAsArrayBuffer(file);
  });
}

/* ============================== DATI SEED ============================== */

function seedStudents() {
  return [
    { id: "s1", nome: "Sofia", cognome: "Ricci", dataNascita: "2016-03-12", livello: "Principianti", stato: "attivo", certificato: "2027-03-01", note: "" },
    { id: "s2", nome: "Leonardo", cognome: "Greco", dataNascita: "2015-07-22", livello: "Principianti", stato: "attivo", certificato: "2026-11-25", note: "" },
    { id: "s3", nome: "Aurora", cognome: "Fontana", dataNascita: "2014-01-05", livello: "Base", stato: "attivo", certificato: "2027-02-10", note: "" },
    { id: "s4", nome: "Matteo", cognome: "Barbieri", dataNascita: "2016-09-30", livello: "Principianti", stato: "attivo", certificato: "2027-01-18", note: "" },
    { id: "s5", nome: "Giulia", cognome: "Moretti", dataNascita: "2012-05-14", livello: "Avanzato", stato: "attivo", certificato: "2027-04-20", note: "" },
    { id: "s6", nome: "Alessandro", cognome: "Villa", dataNascita: "2011-11-02", livello: "Avanzato", stato: "attivo", certificato: "2027-05-12", note: "" },
    { id: "s7", nome: "Chiara", cognome: "De Luca", dataNascita: "2012-08-19", livello: "Avanzato", stato: "sospeso", certificato: "2026-12-01", note: "Sospesa per infortunio" },
    { id: "s8", nome: "Tommaso", cognome: "Marino", dataNascita: "2021-02-08", livello: "Bambini 3-5", stato: "attivo", certificato: "2027-06-01", note: "" },
    { id: "s9", nome: "Beatrice", cognome: "Gallo", dataNascita: "2020-10-17", livello: "Bambini 3-5", stato: "attivo", certificato: "2026-11-15", note: "" },
    { id: "s10", nome: "Francesco", cognome: "Bruno", dataNascita: "2013-04-25", livello: "Base", stato: "attivo", certificato: "2027-01-30", note: "" },
    { id: "s11", nome: "Emma", cognome: "Colombo", dataNascita: "2013-12-11", livello: "Base", stato: "attivo", certificato: "2027-03-22", note: "" },
    { id: "s12", nome: "Riccardo", cognome: "Santoro", dataNascita: "1985-06-09", livello: "Master", stato: "attivo", certificato: "2027-02-14", note: "" },
    { id: "s13", nome: "Valentina", cognome: "Rizzo", dataNascita: "1990-09-27", livello: "Master", stato: "attivo", certificato: "2026-11-20", note: "" },
    { id: "s14", nome: "Davide", cognome: "Ferrara", dataNascita: "2014-02-16", livello: "Intermedio", stato: "attivo", certificato: "2027-05-05", note: "" },
  ];
}

function seedCourses() {
  return [
    { id: "c1", sectionKey: "lg", nome: "Squali Junior", giorni: ["Lun", "Gio"], ora: "17:00", oraFine: "18:00", piscina: "Grande", corsie: [1, 2], livello: "Principianti", istruttore: "Marco Rossi", capienza: 10, inizio: "2026-09-14", fine: "2027-06-12", studentIds: ["s1", "s2", "s3", "s4"], archived: false },
    { id: "c2", sectionKey: "lg", nome: "Delfini Avanzato", giorni: ["Lun", "Gio"], ora: "18:00", oraFine: "19:00", piscina: "Grande", corsie: [3, 4], livello: "Avanzato", istruttore: "Giulia Ferri", capienza: 8, inizio: "2026-09-14", fine: "2027-06-12", studentIds: ["s5", "s6", "s7"], archived: false },
    { id: "c3", sectionKey: "lg", nome: "Acquaticità Bimbi", giorni: ["Lun", "Gio"], ora: "16:00", oraFine: "16:45", piscina: "Piccola", corsie: [1], livello: "Bambini 3-5", istruttore: "Sara Conti", capienza: 6, inizio: "2026-09-14", fine: "2027-06-12", studentIds: ["s8", "s9"], archived: false },
    { id: "c4", sectionKey: "mv", nome: "Rane Base", giorni: ["Mar", "Ven"], ora: "17:00", oraFine: "18:00", piscina: "Grande", corsie: [2, 3], livello: "Base", istruttore: "Marco Rossi", capienza: 10, inizio: "2026-09-15", fine: "2027-06-11", studentIds: ["s1", "s10", "s11"], archived: false },
    { id: "c5", sectionKey: "mv", nome: "Master Nuoto Libero", giorni: ["Mar", "Ven"], ora: "20:00", oraFine: "21:00", piscina: "Grande", corsie: [5, 6], livello: "Master", istruttore: "Luca Bianchi", capienza: 12, inizio: "2026-09-15", fine: "2027-06-11", studentIds: ["s6", "s12", "s13"], archived: false },
    { id: "c6", sectionKey: "mv", nome: "Perfezionamento Dorso", giorni: ["Mar", "Ven"], ora: "19:00", oraFine: "20:00", piscina: "Grande", corsie: [4], livello: "Intermedio", istruttore: "Giulia Ferri", capienza: 8, inizio: "2026-09-15", fine: "2027-06-11", studentIds: ["s5", "s14"], archived: false },
    { id: "c7", sectionKey: "m", nome: "Intensivo Stile Libero", giorni: ["Mer"], ora: "18:00", oraFine: "19:00", piscina: "Grande", corsie: [1, 2, 3], livello: "Intermedio", istruttore: "Marco Rossi", capienza: 15, inizio: "2026-09-16", fine: "2027-06-10", studentIds: ["s2", "s10", "s14"], archived: false },
    { id: "c8", sectionKey: "m", nome: "Acquagym", giorni: ["Mer"], ora: "09:30", oraFine: "10:30", piscina: "Piccola", corsie: [1, 2], livello: "Adulti", istruttore: "Sara Conti", capienza: 20, inizio: "2026-09-16", fine: "2027-06-10", studentIds: ["s12", "s13"], archived: false },
    { id: "c9", sectionKey: "s", nome: "Scuola Nuoto Sabato", giorni: ["Sab"], ora: "09:00", oraFine: "10:00", piscina: "Grande", corsie: [1, 2], livello: "Principianti", istruttore: "Luca Bianchi", capienza: 10, inizio: "2026-09-19", fine: "2027-06-12", studentIds: ["s3", "s9", "s14"], archived: false },
    { id: "c10", sectionKey: "s", nome: "Baby Nuoto", giorni: ["Sab"], ora: "10:00", oraFine: "10:40", piscina: "Piccola", corsie: [1], livello: "Bambini 1-3", istruttore: "Sara Conti", capienza: 6, inizio: "2026-09-19", fine: "2027-06-12", studentIds: ["s8"], archived: false },
  ];
}

function seedAttendance(courses) {
  const att = {};
  courses.forEach((course, ci) => {
    const dates = getSeasonLessonDates(course.giorni)
      .filter((d) => !isHolidayDate(d) && d.getTime() <= TODAY.getTime());
    const courseAtt = {};
    course.studentIds.forEach((sid, si) => {
      const studentAtt = {};
      dates.forEach((d, di) => {
        const h = (ci * 31 + si * 17 + di * 7) % 11;
        let status = "P";
        if (h === 0) status = "A";
        else if (h === 1) status = "L";
        const cell = { status };
        if (status === "L") cell.minutes = 5;
        studentAtt[toKey(d)] = cell;
      });
      courseAtt[sid] = studentAtt;
    });
    att[course.id] = courseAtt;
  });
  // Un paio di recuperi già utilizzati, a scopo dimostrativo
  const c1Dates = Object.keys(att.c1.s1 || {});
  if (c1Dates.length > 2) att.c1.s1[c1Dates[2]] = { status: "R" };
  const c4Dates = Object.keys(att.c4.s1 || {});
  if (c4Dates.length > 3) att.c4.s1[c4Dates[3]] = { status: "R" };
  return att;
}

function seedPayments() {
  return [
    { id: "p1", studentId: "s1", importo: 120, periodo: "Ottobre 2026", data: "2026-10-02", scadenza: "2026-10-05", stato: "Pagato", metodo: "Bonifico", note: "" },
    { id: "p2", studentId: "s1", importo: 120, periodo: "Novembre 2026", data: "", scadenza: "2026-11-05", stato: "Scaduto", metodo: "", note: "" },
    { id: "p3", studentId: "s2", importo: 120, periodo: "Novembre 2026", data: "2026-11-03", scadenza: "2026-11-05", stato: "Pagato", metodo: "Contanti", note: "" },
    { id: "p4", studentId: "s5", importo: 150, periodo: "Novembre 2026", data: "", scadenza: "2026-11-12", stato: "In attesa", metodo: "", note: "" },
    { id: "p5", studentId: "s9", importo: 100, periodo: "Novembre 2026", data: "", scadenza: "2026-11-08", stato: "Scaduto", metodo: "", note: "Sollecitare famiglia" },
    { id: "p6", studentId: "s12", importo: 180, periodo: "Novembre 2026", data: "2026-11-01", scadenza: "2026-11-05", stato: "Pagato", metodo: "POS", note: "" },
    { id: "p7", studentId: "s13", importo: 180, periodo: "Novembre 2026", data: "", scadenza: "2026-11-15", stato: "In attesa", metodo: "", note: "" },
    { id: "p8", studentId: "s10", importo: 120, periodo: "Ottobre 2026", data: "2026-10-04", scadenza: "2026-10-05", stato: "Pagato", metodo: "Bonifico", note: "" },
  ];
}

function seedEvaluations() {
  return [
    { id: "e1", studentId: "s1", criterio: "Tecnica", periodo: "Ottobre 2026", stelle: 3 },
    { id: "e2", studentId: "s1", criterio: "Comportamento", periodo: "Ottobre 2026", stelle: 4 },
    { id: "e3", studentId: "s5", criterio: "Tecnica", periodo: "Ottobre 2026", stelle: 5 },
    { id: "e4", studentId: "s6", criterio: "Resistenza", periodo: "Ottobre 2026", stelle: 4 },
  ];
}

const CRITERI = ["Tecnica", "Resistenza", "Comportamento", "Respirazione", "Partenze"];

/* ============================== COMPONENTI DI SUPPORTO ============================== */

function StatoBadge({ status }) {
  const map = {
    P: { bg: "#E6F5EE", fg: PRESENTE, label: "Presente" },
    A: { bg: "#FBEAE8", fg: ASSENTE, label: "Assente" },
    R: { bg: "#EFE9F9", fg: RECUPERO, label: "Recupero" },
    L: { bg: "#FBF1E0", fg: RITARDO, label: "Ritardo" },
  };
  const s = map[status];
  if (!s) return null;
  return (
    <span className="text-[11px] font-medium px-1.5 py-0.5 rounded" style={{ backgroundColor: s.bg, color: s.fg }}>
      {s.label}
    </span>
  );
}

function CellIcon({ status, minutes }) {
  if (status === "P") return <Check size={16} color={PRESENTE} strokeWidth={3} />;
  if (status === "A") return <X size={16} color={ASSENTE} strokeWidth={3} />;
  if (status === "R") return <RotateCcw size={14} color={RECUPERO} strokeWidth={2.5} />;
  if (status === "L") return (
    <span className="flex items-center gap-0.5">
      <Clock size={13} color={RITARDO} strokeWidth={2.5} />
      <span className="text-[10px] font-semibold" style={{ color: RITARDO }}>{minutes}'</span>
    </span>
  );
  return <span className="text-slate-300 text-xs">–</span>;
}

/* ============================== APP ============================== */

export default function SwimSchoolApp() {
  const [role, setRole] = useState("admin");
  const [view, setView] = useState("registro");
  const [activeSection, setActiveSection] = useState("lg");
  const [openCourseId, setOpenCourseId] = useState(null);
  const [courseViewMode, setCourseViewMode] = useState("mese");
  const [selectedMonthIdx, setSelectedMonthIdx] = useState(2); // Novembre 2026

  const [loaded, setLoaded] = useState(false);
  const [loadError, setLoadError] = useState(false);
  const [students, setStudents] = useState(seedStudents);
  const [courses, setCourses] = useState(seedCourses);
  const [attendance, setAttendance] = useState(() => seedAttendance(seedCourses()));
  const [payments, setPayments] = useState(seedPayments);
  const [evaluations, setEvaluations] = useState(seedEvaluations);

  // Caricamento dati salvati (o inizializzazione con i dati di esempio al primo avvio)
  useEffect(() => {
    let cancelled = false;
    async function load() {
      try {
        const keys = ["students-v1", "courses-v1", "attendance-v1", "payments-v1", "evaluations-v1"];
        const results = await Promise.all(keys.map(async (k) => {
          try { return await storageGet(k); }
          catch { return null; }
        }));
        if (cancelled) return;
        const [sStudents, sCourses, sAttendance, sPayments, sEvaluations] = results;
        const initialCourses = sCourses || seedCourses();
        setCourses(initialCourses);
        setStudents(sStudents || seedStudents());
        setAttendance(sAttendance || seedAttendance(initialCourses));
        setPayments(sPayments || seedPayments());
        setEvaluations(sEvaluations || seedEvaluations());
      } catch (e) {
        if (!cancelled) setLoadError(true);
      } finally {
        if (!cancelled) setLoaded(true);
      }
    }
    load();
    return () => { cancelled = true; };
  }, []);

  // Salvataggio automatico ad ogni modifica (solo dopo il caricamento iniziale)
  useEffect(() => { if (loaded) storageSet("students-v1", students).catch(() => {}); }, [students, loaded]);
  useEffect(() => { if (loaded) storageSet("courses-v1", courses).catch(() => {}); }, [courses, loaded]);
  useEffect(() => { if (loaded) storageSet("attendance-v1", attendance).catch(() => {}); }, [attendance, loaded]);
  useEffect(() => { if (loaded) storageSet("payments-v1", payments).catch(() => {}); }, [payments, loaded]);
  useEffect(() => { if (loaded) storageSet("evaluations-v1", evaluations).catch(() => {}); }, [evaluations, loaded]);

  const [anagraficaId, setAnagraficaId] = useState(null);
  const [addingStudentOpen, setAddingStudentOpen] = useState(false);
  const [studentSearch, setStudentSearch] = useState("");
  const [newStudentForm, setNewStudentForm] = useState(null);
  const [courseForm, setCourseForm] = useState(null); // { mode:'new'|'edit', data:{...} }
  const [showArchived, setShowArchived] = useState(false);
  const [globalSearch, setGlobalSearch] = useState("");
  const fileInputRef = useRef(null);
  const [importSheets, setImportSheets] = useState(null);
  const [importError, setImportError] = useState("");

  const isAdmin = role === "admin";
  const studentById = (id) => students.find((s) => s.id === id);

  /* ---------- Calcoli recuperi (globali per allievo) ---------- */
  function studentAbsenceCount(sid) {
    let n = 0;
    courses.forEach((c) => {
      const ca = attendance[c.id];
      if (!ca || !ca[sid]) return;
      Object.values(ca[sid]).forEach((cell) => { if (cell.status === "A") n++; });
    });
    return n;
  }
  function studentRecuperoUsedCount(sid) {
    let n = 0;
    courses.forEach((c) => {
      const ca = attendance[c.id];
      if (!ca || !ca[sid]) return;
      Object.values(ca[sid]).forEach((cell) => { if (cell.status === "R") n++; });
    });
    return n;
  }
  function studentRecuperi(sid) {
    const assenze = studentAbsenceCount(sid);
    const maturati = Math.floor(assenze / 2);
    const usati = studentRecuperoUsedCount(sid);
    const disponibili = Math.max(0, maturati - usati);
    return { assenze, maturati, usati, disponibili };
  }

  /* ---------- Azioni corso ---------- */
  function handleAddExistingStudent(courseId, studentId) {
    setCourses((prev) => prev.map((c) => c.id === courseId && !c.studentIds.includes(studentId)
      ? { ...c, studentIds: [...c.studentIds, studentId] } : c));
  }
  function handleRemoveStudentFromCourse(courseId, studentId) {
    if (!window.confirm("Rimuovere questo allievo dal corso? L'anagrafica e lo storico presenze restano invariati.")) return;
    setCourses((prev) => prev.map((c) => c.id === courseId
      ? { ...c, studentIds: c.studentIds.filter((id) => id !== studentId) } : c));
  }
  function handleCreateAndAddStudent(courseId, form) {
    const id = "s" + (Math.max(0, ...students.map((s) => parseInt(s.id.slice(1), 10))) + 1);
    const newStudent = { id, nome: form.nome, cognome: form.cognome, dataNascita: form.dataNascita || "2015-01-01", livello: form.livello || "Base", stato: "attivo", certificato: form.certificato || "", note: "" };
    setStudents((prev) => [...prev, newStudent]);
    setCourses((prev) => prev.map((c) => c.id === courseId ? { ...c, studentIds: [...c.studentIds, id] } : c));
    setNewStudentForm(null);
  }
  function handleDuplicateCourse(course) {
    const id = "c" + (Math.max(0, ...courses.map((c) => parseInt(c.id.slice(1), 10))) + 1);
    setCourses((prev) => [...prev, { ...course, id, nome: course.nome + " (copia)", studentIds: [] }]);
  }
  function handleArchiveCourse(courseId, archived) {
    setCourses((prev) => prev.map((c) => c.id === courseId ? { ...c, archived } : c));
    if (archived && openCourseId === courseId) setOpenCourseId(null);
  }
  function handleSaveCourse(form) {
    if (form.mode === "new") {
      const id = "c" + (Math.max(0, ...courses.map((c) => parseInt(c.id.slice(1), 10))) + 1);
      setCourses((prev) => [...prev, { ...form.data, id, studentIds: [], archived: false }]);
    } else {
      setCourses((prev) => prev.map((c) => c.id === form.data.id ? { ...form.data } : c));
    }
    setCourseForm(null);
  }

  /* ---------- Import Excel ---------- */
  function handleFileSelected(e) {
    const file = e.target.files[0];
    e.target.value = "";
    if (!file) return;
    setImportError("");
    parseStudentsWorkbook(file)
      .then((sheets) => {
        if (sheets.length === 0) { setImportError("Nessun dato leggibile trovato nel file."); return; }
        const withMatch = sheets.map((s) => {
          const lower = s.sheetName.trim().toLowerCase();
          const match = courses.find((c) => !c.archived && c.nome.trim().toLowerCase() === lower)
            || courses.find((c) => !c.archived && c.nome.trim().toLowerCase().includes(lower));
          return { ...s, courseId: match ? match.id : "" };
        });
        setImportSheets(withMatch);
      })
      .catch(() => setImportError("Impossibile leggere il file. Verifica che sia un file Excel (.xlsx) valido."));
  }

  function updateImportCourse(sheetName, courseId) {
    setImportSheets((prev) => prev.map((s) => s.sheetName === sheetName ? { ...s, courseId } : s));
  }

  function confirmImport() {
    let newStudents = [...students];
    let newCourses = [...courses];
    let nextIdNum = Math.max(0, ...newStudents.map((s) => parseInt(s.id.slice(1), 10) || 0)) + 1;

    importSheets.forEach((sheet) => {
      if (!sheet.courseId) return;
      sheet.rows.forEach((row) => {
        const nome = String(row.nome || "").trim();
        const cognome = String(row.cognome || "").trim();
        if (!nome && !cognome) return;
        let existing = newStudents.find((s) => s.nome.trim().toLowerCase() === nome.toLowerCase() && s.cognome.trim().toLowerCase() === cognome.toLowerCase());
        let studentId;
        if (existing) {
          studentId = existing.id;
        } else {
          studentId = "s" + nextIdNum++;
          const statoRaw = String(row.stato || "").trim().toLowerCase();
          newStudents.push({
            id: studentId,
            nome, cognome,
            dataNascita: excelDateToKey(row.dataNascita) || "2015-01-01",
            livello: String(row.livello || "").trim() || "Base",
            stato: ["attivo", "sospeso", "ritirato"].includes(statoRaw) ? statoRaw : "attivo",
            certificato: excelDateToKey(row.certificato) || "",
            note: String(row.note || "").trim(),
          });
        }
        const cid = sheet.courseId;
        newCourses = newCourses.map((c) => c.id === cid && !c.studentIds.includes(studentId)
          ? { ...c, studentIds: [...c.studentIds, studentId] } : c);
      });
    });

    setStudents(newStudents);
    setCourses(newCourses);
    setImportSheets(null);
  }

  /* ---------- Presenze ---------- */
  function cycleFor(role_) { return role_ === "admin" ? CYCLE_ADMIN : CYCLE_SEGRETERIA; }

  function handleCellClick(courseId, studentId, dateKey) {
    const cycle = cycleFor(role);
    setAttendance((prev) => {
      const cur = prev[courseId]?.[studentId]?.[dateKey]?.status || "";
      let idx = cycle.indexOf(cur);
      if (idx === -1) idx = 0;
      const next = cycle[(idx + 1) % cycle.length];
      const prevNote = prev[courseId]?.[studentId]?.[dateKey]?.note;
      let minutes;
      if (next === "L") {
        const m = window.prompt("Minuti di ritardo:", "5");
        minutes = m ? (parseInt(m, 10) || 5) : 5;
      }
      const courseAtt = { ...(prev[courseId] || {}) };
      const studentAtt = { ...(courseAtt[studentId] || {}) };
      if (next === "") {
        const { [dateKey]: _drop, ...rest } = studentAtt;
        if (prevNote) rest[dateKey] = { status: "", note: prevNote };
        courseAtt[studentId] = rest;
      } else {
        courseAtt[studentId] = { ...studentAtt, [dateKey]: { status: next, ...(next === "L" ? { minutes } : {}), ...(prevNote ? { note: prevNote } : {}) } };
      }
      return { ...prev, [courseId]: courseAtt };
    });
  }

  function handleNoteClick(e, courseId, studentId, dateKey) {
    e.stopPropagation();
    const existing = attendance[courseId]?.[studentId]?.[dateKey]?.note || "";
    const note = window.prompt("Nota per questa lezione:", existing);
    if (note === null) return;
    setAttendance((prev) => {
      const courseAtt = { ...(prev[courseId] || {}) };
      const studentAtt = { ...(courseAtt[studentId] || {}) };
      const cell = { ...(studentAtt[dateKey] || { status: "" }) };
      cell.note = note;
      courseAtt[studentId] = { ...studentAtt, [dateKey]: cell };
      return { ...prev, [courseId]: courseAtt };
    });
  }

  function markAllPresent(courseId, displayedDates) {
    const todayKey = toKey(TODAY);
    let target = displayedDates.find((d) => toKey(d) === todayKey);
    if (!target) {
      const past = displayedDates.filter((d) => d.getTime() <= TODAY.getTime());
      target = past[past.length - 1];
    }
    if (!target) { alert("Nessuna lezione disponibile in questa vista."); return; }
    const dateKey = toKey(target);
    const course = courses.find((c) => c.id === courseId);
    setAttendance((prev) => {
      const courseAtt = { ...(prev[courseId] || {}) };
      course.studentIds.forEach((sid) => {
        const cell = { ...(courseAtt[sid]?.[dateKey] || {}) };
        cell.status = "P";
        courseAtt[sid] = { ...(courseAtt[sid] || {}), [dateKey]: cell };
      });
      return { ...prev, [courseId]: courseAtt };
    });
  }

  /* ---------- Derivati ---------- */
  const sectionCourses = useMemo(
    () => courses.filter((c) => c.sectionKey === activeSection && !!c.archived === showArchived),
    [courses, activeSection, showArchived]
  );
  const openCourse = openCourseId ? courses.find((c) => c.id === openCourseId) : null;

  const globalResults = useMemo(() => {
    if (!globalSearch.trim()) return { students: [], courses: [] };
    const q = globalSearch.trim().toLowerCase();
    const st = students.filter((s) => `${s.nome} ${s.cognome} ${s.livello}`.toLowerCase().includes(q)).slice(0, 6);
    const co = courses.filter((c) => !c.archived && `${c.nome} ${c.istruttore} ${c.livello}`.toLowerCase().includes(q)).slice(0, 6);
    return { students: st, courses: co };
  }, [globalSearch, students, courses]);

  const todayWeekday = weekdayOf(TODAY);
  const coursesToday = courses.filter((c) => !c.archived && c.giorni.includes(todayWeekday));
  const certExpiring = students.filter((s) => s.certificato && daysUntil(s.certificato) <= 30);
  const paymentsOverdue = payments.filter((p) => p.stato !== "Pagato" && daysUntil(p.scadenza) < 0);
  const paymentsPending = payments.filter((p) => p.stato === "In attesa");
  const almostFull = courses.filter((c) => !c.archived && c.studentIds.length / c.capienza >= 0.8);

  function openStudentCourse(courseId) {
    const c = courses.find((x) => x.id === courseId);
    if (!c) return;
    setActiveSection(c.sectionKey);
    setView("registro");
    setOpenCourseId(courseId);
    setGlobalSearch("");
  }

  /* ============================== RENDER ============================== */

  if (!loaded) {
    return (
      <div className="min-h-screen w-full flex items-center justify-center" style={{ backgroundColor: BG, color: INK_SOFT, fontFamily: "'Manrope', system-ui, sans-serif" }}>
        <div className="flex flex-col items-center gap-2">
          <Waves size={28} color={PRIMARY} />
          <span className="text-[13px] font-semibold">Caricamento registro…</span>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen w-full" style={{ backgroundColor: BG, color: INK, fontFamily: "'Manrope', system-ui, sans-serif" }}>
      {loadError && (
        <div className="text-center text-[12px] font-semibold py-1.5" style={{ backgroundColor: "#FBF1E0", color: RITARDO }}>
          Impossibile recuperare i dati salvati: mostro i dati di esempio. Le modifiche da qui in poi verranno comunque salvate.
        </div>
      )}
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap');
        .grid-body { font-family: system-ui, -apple-system, sans-serif; }
        .overflow-auto { -webkit-overflow-scrolling: touch; }
        html, body { overscroll-behavior-x: none; }
        ::-webkit-scrollbar { height: 8px; width: 8px; }
        ::-webkit-scrollbar-thumb { background: #C3D3D6; border-radius: 8px; }
        ::-webkit-scrollbar-track { background: transparent; }
      `}</style>

      {/* ===== HEADER ===== */}
      <header className="sticky top-0 z-30 border-b" style={{ backgroundColor: PRIMARY_DARK, borderColor: BORDER }}>
        <div className="flex items-center gap-3 px-3 sm:px-5 py-2.5">
          <Waves size={22} color="#DDF4F8" />
          <span className="font-bold text-white text-[15px] sm:text-base tracking-tight">Meeting Evolution · Scuola Nuoto</span>

          <nav className="hidden md:flex items-center gap-1 ml-4">
            <NavBtn icon={<ClipboardList size={15} />} label="Registro" active={view === "registro"} onClick={() => { setView("registro"); setOpenCourseId(null); }} />
            <NavBtn icon={<LayoutDashboard size={15} />} label="Dashboard" active={view === "dashboard"} onClick={() => setView("dashboard")} />
            {isAdmin && <NavBtn icon={<Star size={15} />} label="Valutazioni" active={view === "valutazioni"} onClick={() => setView("valutazioni")} />}
            {isAdmin && <NavBtn icon={<CreditCard size={15} />} label="Pagamenti" active={view === "pagamenti"} onClick={() => setView("pagamenti")} />}
          </nav>

          <div className="ml-auto flex items-center gap-2">
            <div className="relative">
              <Search size={14} className="absolute left-2.5 top-2.5" color="#9DC3CC" />
              <input
                value={globalSearch}
                onChange={(e) => setGlobalSearch(e.target.value)}
                placeholder="Cerca allievo, corso, istruttore…"
                className="pl-8 pr-2 py-1.5 rounded-lg text-[13px] w-40 sm:w-64 outline-none"
                style={{ backgroundColor: "#134A5C", color: "white" }}
              />
              {globalSearch.trim() && (
                <div className="absolute right-0 mt-1 w-72 rounded-lg shadow-xl border overflow-hidden" style={{ backgroundColor: SURFACE, borderColor: BORDER }}>
                  {globalResults.students.length === 0 && globalResults.courses.length === 0 && (
                    <div className="px-3 py-2 text-xs" style={{ color: INK_SOFT }}>Nessun risultato</div>
                  )}
                  {globalResults.students.map((s) => (
                    <button key={s.id} onClick={() => { setAnagraficaId(s.id); setGlobalSearch(""); }}
                      className="w-full text-left px-3 py-2 text-[13px] hover:bg-slate-50 flex items-center justify-between">
                      <span>{s.nome} {s.cognome}</span>
                      <span className="text-[11px]" style={{ color: INK_SOFT }}>{s.livello}</span>
                    </button>
                  ))}
                  {globalResults.courses.map((c) => (
                    <button key={c.id} onClick={() => openStudentCourse(c.id)}
                      className="w-full text-left px-3 py-2 text-[13px] hover:bg-slate-50 flex items-center justify-between">
                      <span>{c.nome}</span>
                      <span className="text-[11px]" style={{ color: INK_SOFT }}>{c.istruttore}</span>
                    </button>
                  ))}
                </div>
              )}
            </div>

            {isAdmin && (
              <>
                <button onClick={() => fileInputRef.current?.click()}
                  className="px-2.5 py-1.5 rounded-lg text-[12px] font-semibold flex items-center gap-1 whitespace-nowrap"
                  style={{ backgroundColor: "#134A5C", color: "#EAFBFE" }}>
                  <Upload size={13} /> <span className="hidden sm:inline">Importa Excel</span>
                </button>
                <input ref={fileInputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleFileSelected} />
              </>
            )}

            <div className="flex items-center rounded-lg overflow-hidden border" style={{ borderColor: "#134A5C" }}>
              <button onClick={() => setRole("admin")} className="px-2.5 py-1.5 text-[12px] font-semibold flex items-center gap-1"
                style={{ backgroundColor: isAdmin ? ACCENT : "transparent", color: isAdmin ? PRIMARY_DARK : "#BEE0E6" }}>
                <ShieldCheck size={13} /> Admin
              </button>
              <button onClick={() => setRole("segreteria")} className="px-2.5 py-1.5 text-[12px] font-semibold"
                style={{ backgroundColor: !isAdmin ? ACCENT : "transparent", color: !isAdmin ? PRIMARY_DARK : "#BEE0E6" }}>
                Segreteria
              </button>
            </div>
          </div>
        </div>

        {/* nav mobile */}
        <div className="flex md:hidden items-center gap-1 px-3 pb-2 overflow-x-auto">
          <NavBtn compact icon={<ClipboardList size={14} />} label="Registro" active={view === "registro"} onClick={() => { setView("registro"); setOpenCourseId(null); }} />
          <NavBtn compact icon={<LayoutDashboard size={14} />} label="Dashboard" active={view === "dashboard"} onClick={() => setView("dashboard")} />
          {isAdmin && <NavBtn compact icon={<Star size={14} />} label="Valutazioni" active={view === "valutazioni"} onClick={() => setView("valutazioni")} />}
          {isAdmin && <NavBtn compact icon={<CreditCard size={14} />} label="Pagamenti" active={view === "pagamenti"} onClick={() => setView("pagamenti")} />}
        </div>
      </header>

      <main className="px-3 sm:px-5 py-4 max-w-[1400px] mx-auto">
        {view === "registro" && !openCourse && (
          <RegistroView
            activeSection={activeSection} setActiveSection={setActiveSection}
            sectionCourses={sectionCourses} isAdmin={isAdmin}
            showArchived={showArchived} setShowArchived={setShowArchived}
            onOpenCourse={setOpenCourseId}
            onNewCourse={() => setCourseForm({ mode: "new", data: { sectionKey: activeSection, nome: "", giorni: SECTIONS.find(s => s.key === activeSection).days, ora: "17:00", oraFine: "18:00", piscina: "Grande", corsie: [1], livello: "", istruttore: "", capienza: 10, inizio: "2026-09-14", fine: "2027-06-12" } })}
            onEditCourse={(c) => setCourseForm({ mode: "edit", data: { ...c } })}
            onDuplicateCourse={handleDuplicateCourse}
            onArchiveCourse={handleArchiveCourse}
          />
        )}

        {view === "registro" && openCourse && (
          <CourseView
            course={openCourse} isAdmin={isAdmin} role={role}
            students={students} attendance={attendance}
            courseViewMode={courseViewMode} setCourseViewMode={setCourseViewMode}
            selectedMonthIdx={selectedMonthIdx} setSelectedMonthIdx={setSelectedMonthIdx}
            onBack={() => setOpenCourseId(null)}
            onCellClick={handleCellClick} onNoteClick={handleNoteClick}
            onMarkAllPresent={markAllPresent}
            onOpenAnagrafica={setAnagraficaId}
            onRemoveStudent={handleRemoveStudentFromCourse}
            addingStudentOpen={addingStudentOpen} setAddingStudentOpen={setAddingStudentOpen}
            studentSearch={studentSearch} setStudentSearch={setStudentSearch}
            newStudentForm={newStudentForm} setNewStudentForm={setNewStudentForm}
            allStudents={students}
            onAddExisting={handleAddExistingStudent}
            onCreateAndAdd={handleCreateAndAddStudent}
            studentRecuperi={studentRecuperi}
          />
        )}

        {view === "dashboard" && (
          <DashboardView
            coursesToday={coursesToday} certExpiring={certExpiring}
            paymentsOverdue={paymentsOverdue} paymentsPending={paymentsPending}
            almostFull={almostFull} studentById={studentById}
            onOpenCourse={openStudentCourse} onOpenStudent={setAnagraficaId}
          />
        )}

        {view === "valutazioni" && isAdmin && (
          <ValutazioniView students={students} evaluations={evaluations} setEvaluations={setEvaluations} />
        )}

        {view === "pagamenti" && isAdmin && (
          <PagamentiView students={students} payments={payments} setPayments={setPayments} />
        )}
      </main>

      {anagraficaId && (
        <AnagraficaModal
          student={studentById(anagraficaId)} isAdmin={isAdmin}
          courses={courses} attendance={attendance} payments={payments} evaluations={evaluations}
          studentRecuperi={studentRecuperi}
          onClose={() => setAnagraficaId(null)}
          onSave={(updated) => setStudents((prev) => prev.map((s) => s.id === updated.id ? updated : s))}
        />
      )}

      {courseForm && (
        <CourseFormModal
          form={courseForm}
          onCancel={() => setCourseForm(null)}
          onSave={handleSaveCourse}
        />
      )}

      {importError && (
        <div className="fixed bottom-4 left-1/2 -translate-x-1/2 z-50 px-4 py-2.5 rounded-xl text-[13px] font-semibold text-white shadow-lg" style={{ backgroundColor: ASSENTE }}>
          {importError}
          <button onClick={() => setImportError("")} className="ml-3 underline">Chiudi</button>
        </div>
      )}

      {importSheets && (
        <ImportExcelModal
          sheets={importSheets}
          courses={courses.filter((c) => !c.archived)}
          onUpdateCourse={updateImportCourse}
          onConfirm={confirmImport}
          onCancel={() => setImportSheets(null)}
        />
      )}
    </div>
  );
}

/* ============================== SOTTOCOMPONENTI ============================== */

function NavBtn({ icon, label, active, onClick, compact }) {
  return (
    <button onClick={onClick}
      className={`flex items-center gap-1.5 rounded-lg font-semibold whitespace-nowrap ${compact ? "px-2.5 py-1.5 text-[12px]" : "px-3 py-1.5 text-[13px]"}`}
      style={{ backgroundColor: active ? "#134A5C" : "transparent", color: active ? "#EAFBFE" : "#9DC3CC" }}>
      {icon} {label}
    </button>
  );
}

function RegistroView({ activeSection, setActiveSection, sectionCourses, isAdmin, showArchived, setShowArchived, onOpenCourse, onNewCourse, onEditCourse, onDuplicateCourse, onArchiveCourse }) {
  return (
    <div>
      <div className="flex flex-wrap items-center gap-2 mb-4">
        {SECTIONS.map((s) => (
          <button key={s.key} onClick={() => setActiveSection(s.key)}
            className="px-3.5 py-2 rounded-xl text-[13px] font-bold border transition"
            style={{
              backgroundColor: activeSection === s.key ? PRIMARY : SURFACE,
              color: activeSection === s.key ? "white" : INK,
              borderColor: activeSection === s.key ? PRIMARY : BORDER,
            }}>
            {s.label}
          </button>
        ))}
        <div className="ml-auto flex items-center gap-2">
          {isAdmin && (
            <button onClick={() => setShowArchived(!showArchived)} className="text-[12px] font-medium px-2.5 py-1.5 rounded-lg border" style={{ borderColor: BORDER, color: INK_SOFT }}>
              {showArchived ? "Mostra corsi attivi" : "Mostra corsi archiviati"}
            </button>
          )}
          {isAdmin && !showArchived && (
            <button onClick={onNewCourse} className="text-[13px] font-bold px-3 py-2 rounded-xl text-white flex items-center gap-1.5" style={{ backgroundColor: PRIMARY }}>
              <Plus size={15} /> Nuovo corso
            </button>
          )}
        </div>
      </div>

      {sectionCourses.length === 0 && (
        <div className="text-center py-16 rounded-2xl border-2 border-dashed" style={{ borderColor: BORDER, color: INK_SOFT }}>
          {showArchived ? "Nessun corso archiviato in questa sezione." : "Nessun corso in questa sezione. Crea il primo corso per iniziare."}
        </div>
      )}

      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
        {sectionCourses.map((c) => (
          <div key={c.id} className="rounded-2xl border p-4 relative group" style={{ backgroundColor: SURFACE, borderColor: BORDER }}>
            <button onClick={() => onOpenCourse(c.id)} className="text-left w-full">
              <div className="flex items-start justify-between gap-2">
                <h3 className="font-bold text-[15px] leading-snug">{c.nome}</h3>
                <span className="text-[11px] font-semibold px-2 py-0.5 rounded-full whitespace-nowrap" style={{ backgroundColor: c.studentIds.length >= c.capienza ? "#FBEAE8" : "#EAF4F0", color: c.studentIds.length >= c.capienza ? ASSENTE : PRESENTE }}>
                  {c.studentIds.length}/{c.capienza}
                </span>
              </div>
              <div className="text-[12px] mt-1.5 space-y-0.5" style={{ color: INK_SOFT }}>
                <div>{c.giorni.join(" · ")} &nbsp;·&nbsp; {c.ora}–{c.oraFine}</div>
                <div>Vasca {c.piscina} · Corsie {c.corsie.join(", ")}</div>
                <div>{c.livello} · Istr. {c.istruttore}</div>
              </div>
            </button>
            {isAdmin && (
              <div className="flex items-center gap-1 mt-3 pt-2 border-t" style={{ borderColor: BORDER }}>
                <IconBtn title="Modifica" onClick={() => onEditCourse(c)}><Pencil size={13} /></IconBtn>
                <IconBtn title="Duplica" onClick={() => onDuplicateCourse(c)}><Copy size={13} /></IconBtn>
                {!c.archived ? (
                  <IconBtn title="Archivia" onClick={() => onArchiveCourse(c.id, true)}><Archive size={13} /></IconBtn>
                ) : (
                  <IconBtn title="Ripristina" onClick={() => onArchiveCourse(c.id, false)}><ArchiveRestore size={13} /></IconBtn>
                )}
              </div>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}

function IconBtn({ children, onClick, title }) {
  return (
    <button onClick={onClick} title={title} className="p-1.5 rounded-lg hover:bg-slate-100" style={{ color: INK_SOFT }}>
      {children}
    </button>
  );
}

function CourseView({
  course, isAdmin, role, students, attendance, courseViewMode, setCourseViewMode,
  selectedMonthIdx, setSelectedMonthIdx, onBack, onCellClick, onNoteClick, onMarkAllPresent,
  onOpenAnagrafica, onRemoveStudent, addingStudentOpen, setAddingStudentOpen,
  studentSearch, setStudentSearch, newStudentForm, setNewStudentForm, allStudents,
  onAddExisting, onCreateAndAdd, studentRecuperi,
}) {
  const courseStart = new Date(course.inizio);
  const courseEnd = new Date(course.fine);

  const displayedDates = useMemo(() => {
    let dates;
    if (courseViewMode === "mese") {
      const sm = SEASON_MONTHS[selectedMonthIdx];
      dates = getLessonDatesForMonth(sm.y, sm.m, course.giorni);
    } else {
      dates = getSeasonLessonDates(course.giorni);
    }
    return dates.filter((d) => d.getTime() >= courseStart.getTime() && d.getTime() <= courseEnd.getTime());
  }, [courseViewMode, selectedMonthIdx, course]);

  const enrolled = course.studentIds.map((id) => students.find((s) => s.id === id)).filter(Boolean);

  const searchResults = allStudents.filter((s) =>
    !course.studentIds.includes(s.id) &&
    `${s.nome} ${s.cognome}`.toLowerCase().includes(studentSearch.toLowerCase()) &&
    studentSearch.trim().length > 0
  ).slice(0, 8);

  return (
    <div>
      <div className="flex items-center gap-2 mb-3">
        <button onClick={onBack} className="p-2 rounded-lg hover:bg-slate-200/60 flex items-center gap-1 text-[13px] font-semibold" style={{ color: INK_SOFT }}>
          <ArrowLeft size={16} /> Torna alle sezioni
        </button>
      </div>

      <div className="rounded-2xl border p-4 mb-3" style={{ backgroundColor: SURFACE, borderColor: BORDER }}>
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div>
            <h2 className="font-extrabold text-lg">{course.nome}</h2>
            <div className="text-[13px] mt-1" style={{ color: INK_SOFT }}>
              {course.giorni.join(" · ")} · {course.ora}–{course.oraFine} · Vasca {course.piscina} (corsie {course.corsie.join(", ")}) · {course.livello} · Istr. {course.istruttore} · {enrolled.length}/{course.capienza} iscritti
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button onClick={() => onMarkAllPresent(course.id, displayedDates)}
              className="text-[13px] font-bold px-3 py-2 rounded-xl text-white flex items-center gap-1.5" style={{ backgroundColor: PRESENTE }}>
              <Check size={15} /> Segna tutti presenti
            </button>
            <button onClick={() => setAddingStudentOpen(!addingStudentOpen)}
              className="text-[13px] font-bold px-3 py-2 rounded-xl text-white flex items-center gap-1.5" style={{ backgroundColor: PRIMARY }}>
              <UserPlus size={15} /> Aggiungi allievo
            </button>
          </div>
        </div>

        {addingStudentOpen && (
          <div className="mt-3 p-3 rounded-xl border" style={{ borderColor: BORDER, backgroundColor: BG }}>
            <div className="relative mb-2">
              <Search size={14} className="absolute left-2.5 top-2.5" color={INK_SOFT} />
              <input autoFocus value={studentSearch} onChange={(e) => setStudentSearch(e.target.value)}
                placeholder="Cerca per nome o cognome…" className="w-full pl-8 pr-2 py-2 rounded-lg text-[13px] border outline-none" style={{ borderColor: BORDER }} />
            </div>
            {studentSearch.trim() && (
              <div className="max-h-40 overflow-auto rounded-lg border mb-2" style={{ borderColor: BORDER, backgroundColor: SURFACE }}>
                {searchResults.length === 0 && <div className="px-3 py-2 text-[12px]" style={{ color: INK_SOFT }}>Nessun allievo trovato</div>}
                {searchResults.map((s) => (
                  <button key={s.id} onClick={() => { onAddExisting(course.id, s.id); setStudentSearch(""); }}
                    className="w-full text-left px-3 py-2 text-[13px] hover:bg-slate-50 flex items-center justify-between">
                    <span>{s.nome} {s.cognome}</span>
                    <span className="text-[11px]" style={{ color: INK_SOFT }}>{s.livello}</span>
                  </button>
                ))}
              </div>
            )}
            {!newStudentForm ? (
              <button onClick={() => setNewStudentForm({ nome: "", cognome: "", dataNascita: "", livello: "" })}
                className="text-[12px] font-semibold flex items-center gap-1" style={{ color: PRIMARY }}>
                <Plus size={13} /> Nuovo allievo
              </button>
            ) : (
              <div className="grid grid-cols-2 gap-2 p-2 rounded-lg border" style={{ borderColor: BORDER, backgroundColor: SURFACE }}>
                <input placeholder="Nome" value={newStudentForm.nome} onChange={(e) => setNewStudentForm({ ...newStudentForm, nome: e.target.value })} className="px-2 py-1.5 rounded border text-[13px]" style={{ borderColor: BORDER }} />
                <input placeholder="Cognome" value={newStudentForm.cognome} onChange={(e) => setNewStudentForm({ ...newStudentForm, cognome: e.target.value })} className="px-2 py-1.5 rounded border text-[13px]" style={{ borderColor: BORDER }} />
                <input type="date" placeholder="Data di nascita" value={newStudentForm.dataNascita} onChange={(e) => setNewStudentForm({ ...newStudentForm, dataNascita: e.target.value })} className="px-2 py-1.5 rounded border text-[13px]" style={{ borderColor: BORDER }} />
                <input placeholder="Livello" value={newStudentForm.livello} onChange={(e) => setNewStudentForm({ ...newStudentForm, livello: e.target.value })} className="px-2 py-1.5 rounded border text-[13px]" style={{ borderColor: BORDER }} />
                <div className="col-span-2 flex gap-2 mt-1">
                  <button onClick={() => onCreateAndAdd(course.id, newStudentForm)} disabled={!newStudentForm.nome || !newStudentForm.cognome}
                    className="text-[12px] font-bold px-3 py-1.5 rounded-lg text-white disabled:opacity-40" style={{ backgroundColor: PRIMARY }}>Crea e aggiungi al corso</button>
                  <button onClick={() => setNewStudentForm(null)} className="text-[12px] font-semibold px-3 py-1.5 rounded-lg" style={{ color: INK_SOFT }}>Annulla</button>
                </div>
              </div>
            )}
          </div>
        )}
      </div>

      <div className="flex flex-wrap items-center gap-2 mb-2">
        <div className="flex rounded-lg overflow-hidden border" style={{ borderColor: BORDER }}>
          <button onClick={() => setCourseViewMode("mese")} className="px-3 py-1.5 text-[12px] font-bold" style={{ backgroundColor: courseViewMode === "mese" ? PRIMARY : SURFACE, color: courseViewMode === "mese" ? "white" : INK }}>Vista mensile</button>
          <button onClick={() => setCourseViewMode("stagione")} className="px-3 py-1.5 text-[12px] font-bold" style={{ backgroundColor: courseViewMode === "stagione" ? PRIMARY : SURFACE, color: courseViewMode === "stagione" ? "white" : INK }}>Vista stagione</button>
        </div>
        {courseViewMode === "mese" && (
          <div className="flex items-center gap-1 flex-wrap">
            {SEASON_MONTHS.map((sm, i) => (
              <button key={i} onClick={() => setSelectedMonthIdx(i)}
                className="px-2.5 py-1 rounded-lg text-[12px] font-semibold border"
                style={{ backgroundColor: selectedMonthIdx === i ? ACCENT : SURFACE, color: selectedMonthIdx === i ? PRIMARY_DARK : INK_SOFT, borderColor: selectedMonthIdx === i ? ACCENT : BORDER }}>
                {sm.label.slice(0, 3)}
              </button>
            ))}
          </div>
        )}
      </div>

      <div className="rounded-2xl border overflow-hidden" style={{ borderColor: BORDER, backgroundColor: SURFACE }}>
        <div className="overflow-auto max-h-[65vh] grid-body" style={{ WebkitOverflowScrolling: "touch", overscrollBehavior: "contain" }}>
          <table className="border-collapse text-[13px]" style={{ width: "max-content", minWidth: "100%" }}>
            <thead>
              <tr>
                <th className="sticky left-0 top-0 z-20 text-left px-3 py-2 border-b border-r font-bold" style={{ backgroundColor: SURFACE, borderColor: BORDER, minWidth: 160 }}>Allievo</th>
                {displayedDates.map((d) => {
                  const hol = isHolidayDate(d);
                  const today = sameDay(d, TODAY);
                  return (
                    <th key={toKey(d)} className="sticky top-0 z-10 px-1 py-2 border-b text-center font-semibold"
                      style={{ backgroundColor: hol ? FESTIVO_BG : (today ? "#DFF3F6" : SURFACE), borderColor: BORDER, minWidth: 56, color: hol ? "#9AA5A8" : INK }}>
                      <div>{fmtShort(d)}</div>
                      <div className="text-[10px] font-normal" style={{ color: INK_SOFT }}>{hol ? "Festivo" : weekdayOf(d)}</div>
                    </th>
                  );
                })}
              </tr>
            </thead>
            <tbody>
              {enrolled.length === 0 && (
                <tr><td className="px-3 py-6 text-center" colSpan={displayedDates.length + 1} style={{ color: INK_SOFT }}>Nessun allievo iscritto. Usa "Aggiungi allievo" per iniziare.</td></tr>
              )}
              {enrolled.map((st) => {
                const rec = studentRecuperi(st.id);
                return (
                  <tr key={st.id} className="group">
                    <td className="sticky left-0 z-10 px-3 py-1.5 border-r border-b" style={{ backgroundColor: SURFACE, borderColor: BORDER }}>
                      <div className="flex items-center justify-between gap-1">
                        <button onClick={() => onOpenAnagrafica(st.id)} className="text-left font-semibold hover:underline truncate" style={{ maxWidth: 120 }}>
                          {st.nome} {st.cognome}
                        </button>
                        {isAdmin && (
                          <button onClick={() => onRemoveStudent(course.id, st.id)} title="Rimuovi dal corso" className="opacity-0 group-hover:opacity-100 p-0.5">
                            <Trash2 size={12} color={ASSENTE} />
                          </button>
                        )}
                      </div>
                      {rec.disponibili > 0 && (
                        <span className="text-[10px] font-bold px-1.5 py-0.5 rounded" style={{ backgroundColor: "#EFE9F9", color: RECUPERO }}>R disp. {rec.disponibili}</span>
                      )}
                    </td>
                    {displayedDates.map((d) => {
                      const hol = isHolidayDate(d);
                      const dateKey = toKey(d);
                      const cell = attendance[course.id]?.[st.id]?.[dateKey];
                      const status = cell?.status || "";
                      const today = sameDay(d, TODAY);
                      return (
                        <td key={dateKey} className="border-b text-center relative"
                          style={{ borderColor: BORDER, backgroundColor: hol ? FESTIVO_BG : (today ? "#F2FBFC" : "transparent") }}>
                          {hol ? (
                            <div className="py-2 text-slate-300 text-xs">—</div>
                          ) : (
                            <button onClick={() => onCellClick(course.id, st.id, dateKey)}
                              className="w-full h-full py-2.5 flex items-center justify-center hover:bg-slate-50 relative">
                              <CellIcon status={status} minutes={cell?.minutes} />
                              {cell?.note && <StickyNote size={9} color={ACCENT} className="absolute top-0.5 right-0.5" />}
                            </button>
                          )}
                          {!hol && (
                            <button onClick={(e) => onNoteClick(e, course.id, st.id, dateKey)} title="Nota"
                              className="absolute bottom-0 right-0 opacity-0 hover:opacity-100 group-hover:opacity-40 hover:!opacity-100 p-0.5">
                              <StickyNote size={9} color={INK_SOFT} />
                            </button>
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
      </div>
      <div className="flex flex-wrap gap-3 mt-2 text-[11px]" style={{ color: INK_SOFT }}>
        <span className="flex items-center gap-1"><Check size={12} color={PRESENTE} /> Presente</span>
        <span className="flex items-center gap-1"><X size={12} color={ASSENTE} /> Assente</span>
        <span className="flex items-center gap-1"><RotateCcw size={12} color={RECUPERO} /> Recupero {!isAdmin && "(solo Admin)"}</span>
        <span className="flex items-center gap-1"><Clock size={12} color={RITARDO} /> Ritardo</span>
        <span>Tocca una cella per cambiare stato · tocca l'icona nota in basso a destra per aggiungere una nota</span>
      </div>
    </div>
  );
}

function DashboardView({ coursesToday, certExpiring, paymentsOverdue, paymentsPending, almostFull, studentById, onOpenCourse, onOpenStudent }) {
  return (
    <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
      <Card title="Corsi di oggi" icon={<CalendarIcon size={16} />}>
        {coursesToday.length === 0 && <Empty>Nessun corso in programma oggi.</Empty>}
        {coursesToday.map((c) => (
          <RowBtn key={c.id} onClick={() => onOpenCourse(c.id)}>
            <span className="font-semibold">{c.nome}</span>
            <span className="text-[12px]" style={{ color: INK_SOFT }}>{c.ora}–{c.oraFine} · {c.studentIds.length}/{c.capienza}</span>
          </RowBtn>
        ))}
      </Card>

      <Card title="Certificati in scadenza (30 gg)" icon={<AlertTriangle size={16} color={RITARDO} />}>
        {certExpiring.length === 0 && <Empty>Nessun certificato in scadenza.</Empty>}
        {certExpiring.map((s) => {
          const d = daysUntil(s.certificato);
          return (
            <RowBtn key={s.id} onClick={() => onOpenStudent(s.id)}>
              <span className="font-semibold">{s.nome} {s.cognome}</span>
              <span className="text-[12px] font-semibold" style={{ color: d < 0 ? ASSENTE : RITARDO }}>{d < 0 ? `Scaduto da ${-d} gg` : `Scade tra ${d} gg`}</span>
            </RowBtn>
          );
        })}
      </Card>

      <Card title="Pagamenti scaduti" icon={<CreditCard size={16} color={ASSENTE} />}>
        {paymentsOverdue.length === 0 && <Empty>Nessun pagamento scaduto.</Empty>}
        {paymentsOverdue.map((p) => {
          const s = studentById(p.studentId);
          return (
            <RowBtn key={p.id} onClick={() => onOpenStudent(p.studentId)}>
              <span className="font-semibold">{s ? `${s.nome} ${s.cognome}` : "—"}</span>
              <span className="text-[12px]" style={{ color: INK_SOFT }}>{p.periodo} · €{p.importo}</span>
            </RowBtn>
          );
        })}
        {paymentsPending.length > 0 && (
          <div className="mt-2 pt-2 border-t text-[12px]" style={{ borderColor: BORDER, color: INK_SOFT }}>
            + {paymentsPending.length} pagamenti in attesa
          </div>
        )}
      </Card>

      <Card title="Corsi pieni o quasi pieni" icon={<Users size={16} />}>
        {almostFull.length === 0 && <Empty>Nessun corso vicino alla capienza massima.</Empty>}
        {almostFull.map((c) => (
          <RowBtn key={c.id} onClick={() => onOpenCourse(c.id)}>
            <span className="font-semibold">{c.nome}</span>
            <span className="text-[12px] font-semibold" style={{ color: c.studentIds.length >= c.capienza ? ASSENTE : RITARDO }}>{c.studentIds.length}/{c.capienza}</span>
          </RowBtn>
        ))}
      </Card>
    </div>
  );
}

function Card({ title, icon, children }) {
  return (
    <div className="rounded-2xl border p-4" style={{ backgroundColor: SURFACE, borderColor: BORDER }}>
      <div className="flex items-center gap-2 font-bold text-[14px] mb-2">{icon} {title}</div>
      <div className="space-y-1">{children}</div>
    </div>
  );
}
function Empty({ children }) { return <div className="text-[13px] py-2" style={{ color: INK_SOFT }}>{children}</div>; }
function RowBtn({ children, onClick }) {
  return (
    <button onClick={onClick} className="w-full flex items-center justify-between px-2.5 py-2 rounded-lg hover:bg-slate-50 text-left text-[13px]">
      {children}
    </button>
  );
}

function ValutazioniView({ students, evaluations, setEvaluations }) {
  const [studentId, setStudentId] = useState(students[0]?.id || "");
  const [criterio, setCriterio] = useState(CRITERI[0]);
  const [periodo, setPeriodo] = useState("Novembre 2026");
  const [stelle, setStelle] = useState(3);

  const studentEvals = evaluations.filter((e) => e.studentId === studentId);

  function addEval() {
    setEvaluations((prev) => [...prev, { id: "e" + (prev.length + 1) + Date.now(), studentId, criterio, periodo, stelle }]);
  }

  return (
    <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
      <div className="rounded-2xl border p-3" style={{ backgroundColor: SURFACE, borderColor: BORDER }}>
        <div className="font-bold text-[13px] mb-2" style={{ color: INK_SOFT }}>ALLIEVI</div>
        <div className="space-y-0.5 max-h-[60vh] overflow-auto">
          {students.map((s) => (
            <button key={s.id} onClick={() => setStudentId(s.id)} className="w-full text-left px-2.5 py-1.5 rounded-lg text-[13px]"
              style={{ backgroundColor: studentId === s.id ? "#EAF6F8" : "transparent", fontWeight: studentId === s.id ? 700 : 500 }}>
              {s.nome} {s.cognome}
            </button>
          ))}
        </div>
      </div>

      <div className="lg:col-span-2 rounded-2xl border p-4" style={{ backgroundColor: SURFACE, borderColor: BORDER }}>
        <div className="font-bold mb-3">Nuova valutazione</div>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 mb-4">
          <select value={criterio} onChange={(e) => setCriterio(e.target.value)} className="px-2 py-2 rounded-lg border text-[13px]" style={{ borderColor: BORDER }}>
            {CRITERI.map((c) => <option key={c}>{c}</option>)}
          </select>
          <input value={periodo} onChange={(e) => setPeriodo(e.target.value)} placeholder="Periodo" className="px-2 py-2 rounded-lg border text-[13px]" style={{ borderColor: BORDER }} />
          <div className="flex items-center gap-1 px-2 py-2 rounded-lg border" style={{ borderColor: BORDER }}>
            {[1, 2, 3, 4, 5].map((n) => (
              <button key={n} onClick={() => setStelle(n)}>
                <Star size={16} fill={n <= stelle ? "#E0A230" : "none"} color="#E0A230" />
              </button>
            ))}
          </div>
          <button onClick={addEval} className="text-[13px] font-bold px-3 py-2 rounded-lg text-white" style={{ backgroundColor: PRIMARY }}>Aggiungi</button>
        </div>

        <div className="font-bold text-[13px] mb-1" style={{ color: INK_SOFT }}>STORICO</div>
        <div className="space-y-1">
          {studentEvals.length === 0 && <Empty>Nessuna valutazione registrata.</Empty>}
          {studentEvals.map((e) => (
            <div key={e.id} className="flex items-center justify-between px-3 py-2 rounded-lg" style={{ backgroundColor: BG }}>
              <div>
                <div className="text-[13px] font-semibold">{e.criterio}</div>
                <div className="text-[11px]" style={{ color: INK_SOFT }}>{e.periodo}</div>
              </div>
              <div className="flex">
                {[1, 2, 3, 4, 5].map((n) => <Star key={n} size={14} fill={n <= e.stelle ? "#E0A230" : "none"} color="#E0A230" />)}
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function PagamentiView({ students, payments, setPayments }) {
  const [filter, setFilter] = useState("Tutti");
  const [form, setForm] = useState(null);

  const studentName = (id) => { const s = students.find((x) => x.id === id); return s ? `${s.nome} ${s.cognome}` : "—"; };
  const filtered = payments.filter((p) => filter === "Tutti" || p.stato === filter);

  function toggleStato(id, stato) {
    setPayments((prev) => prev.map((p) => p.id === id ? { ...p, stato, data: stato === "Pagato" ? toKey(TODAY) : p.data } : p));
  }
  function addPayment() {
    if (!form?.studentId) return;
    setPayments((prev) => [...prev, { id: "p" + (prev.length + 1) + Date.now(), studentId: form.studentId, importo: parseFloat(form.importo) || 0, periodo: form.periodo, data: "", scadenza: form.scadenza, stato: "In attesa", metodo: "", note: form.note || "" }]);
    setForm(null);
  }

  return (
    <div>
      <div className="flex flex-wrap items-center gap-2 mb-3">
        {["Tutti", "Pagato", "In attesa", "Scaduto"].map((f) => (
          <button key={f} onClick={() => setFilter(f)} className="px-3 py-1.5 rounded-lg text-[12px] font-bold border"
            style={{ backgroundColor: filter === f ? PRIMARY : SURFACE, color: filter === f ? "white" : INK, borderColor: filter === f ? PRIMARY : BORDER }}>
            {f}
          </button>
        ))}
        <button onClick={() => setForm({ studentId: students[0]?.id, importo: "", periodo: "", scadenza: "" })}
          className="ml-auto text-[13px] font-bold px-3 py-2 rounded-xl text-white flex items-center gap-1.5" style={{ backgroundColor: PRIMARY }}>
          <Plus size={15} /> Nuovo pagamento
        </button>
      </div>

      {form && (
        <div className="grid grid-cols-2 sm:grid-cols-5 gap-2 mb-3 p-3 rounded-xl border" style={{ borderColor: BORDER, backgroundColor: SURFACE }}>
          <select value={form.studentId} onChange={(e) => setForm({ ...form, studentId: e.target.value })} className="px-2 py-2 rounded-lg border text-[13px]" style={{ borderColor: BORDER }}>
            {students.map((s) => <option key={s.id} value={s.id}>{s.nome} {s.cognome}</option>)}
          </select>
          <input placeholder="Importo €" value={form.importo} onChange={(e) => setForm({ ...form, importo: e.target.value })} className="px-2 py-2 rounded-lg border text-[13px]" style={{ borderColor: BORDER }} />
          <input placeholder="Periodo" value={form.periodo} onChange={(e) => setForm({ ...form, periodo: e.target.value })} className="px-2 py-2 rounded-lg border text-[13px]" style={{ borderColor: BORDER }} />
          <input type="date" value={form.scadenza} onChange={(e) => setForm({ ...form, scadenza: e.target.value })} className="px-2 py-2 rounded-lg border text-[13px]" style={{ borderColor: BORDER }} />
          <button onClick={addPayment} className="text-[13px] font-bold px-3 py-2 rounded-lg text-white" style={{ backgroundColor: PRIMARY }}>Salva</button>
        </div>
      )}

      <div className="rounded-2xl border overflow-hidden" style={{ borderColor: BORDER, backgroundColor: SURFACE }}>
        <table className="w-full text-[13px]">
          <thead>
            <tr style={{ backgroundColor: BG }}>
              {["Allievo", "Periodo", "Importo", "Scadenza", "Stato", "Metodo", "Note", ""].map((h) => (
                <th key={h} className="text-left px-3 py-2 font-bold text-[12px]" style={{ color: INK_SOFT }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map((p) => (
              <tr key={p.id} className="border-t" style={{ borderColor: BORDER }}>
                <td className="px-3 py-2 font-semibold">{studentName(p.studentId)}</td>
                <td className="px-3 py-2">{p.periodo}</td>
                <td className="px-3 py-2">€{p.importo}</td>
                <td className="px-3 py-2">{p.scadenza}</td>
                <td className="px-3 py-2">
                  <span className="text-[11px] font-bold px-2 py-0.5 rounded-full" style={{ backgroundColor: p.stato === "Pagato" ? "#E6F5EE" : p.stato === "Scaduto" ? "#FBEAE8" : "#FBF1E0", color: p.stato === "Pagato" ? PRESENTE : p.stato === "Scaduto" ? ASSENTE : RITARDO }}>{p.stato}</span>
                </td>
                <td className="px-3 py-2">{p.metodo || "—"}</td>
                <td className="px-3 py-2" style={{ color: INK_SOFT }}>{p.note || "—"}</td>
                <td className="px-3 py-2">
                  {p.stato !== "Pagato" && (
                    <button onClick={() => toggleStato(p.id, "Pagato")} className="text-[11px] font-bold px-2 py-1 rounded-lg" style={{ backgroundColor: "#E6F5EE", color: PRESENTE }}>Segna pagato</button>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function AnagraficaModal({ student, isAdmin, courses, attendance, payments, evaluations, studentRecuperi, onClose, onSave }) {
  const [edit, setEdit] = useState(false);
  const [form, setForm] = useState(student);
  if (!student) return null;

  const studentCourses = courses.filter((c) => c.studentIds.includes(student.id));
  const rec = studentRecuperi(student.id);
  const studentPayments = payments.filter((p) => p.studentId === student.id);
  const studentEvals = evaluations.filter((e) => e.studentId === student.id);
  const certDays = student.certificato ? daysUntil(student.certificato) : null;

  function perCourseStats(courseId) {
    const ca = attendance[courseId]?.[student.id] || {};
    let p = 0, a = 0, r = 0, l = 0;
    Object.values(ca).forEach((c) => { if (c.status === "P") p++; else if (c.status === "A") a++; else if (c.status === "R") r++; else if (c.status === "L") l++; });
    return { p, a, r, l };
  }

  return (
    <div className="fixed inset-0 z-40 flex items-start sm:items-center justify-center p-0 sm:p-4" style={{ backgroundColor: "rgba(10,25,32,0.55)" }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} className="w-full sm:max-w-2xl sm:rounded-2xl bg-white h-full sm:h-auto sm:max-h-[90vh] overflow-auto" style={{ backgroundColor: SURFACE }}>
        <div className="sticky top-0 flex items-center justify-between px-5 py-3.5 border-b" style={{ borderColor: BORDER, backgroundColor: SURFACE }}>
          <div>
            <h2 className="font-extrabold text-lg">{student.nome} {student.cognome}</h2>
            <div className="text-[12px]" style={{ color: INK_SOFT }}>{calcAge(student.dataNascita)} anni · {student.livello}</div>
          </div>
          <div className="flex items-center gap-2">
            {isAdmin && !edit && <button onClick={() => { setForm(student); setEdit(true); }} className="p-2 rounded-lg hover:bg-slate-100"><Pencil size={15} /></button>}
            <button onClick={onClose} className="p-2 rounded-lg hover:bg-slate-100"><X size={18} /></button>
          </div>
        </div>

        <div className="p-5 space-y-4">
          {edit ? (
            <div className="grid grid-cols-2 gap-2">
              <Field label="Nome"><input value={form.nome} onChange={(e) => setForm({ ...form, nome: e.target.value })} className="px-2 py-1.5 rounded border w-full text-[13px]" style={{ borderColor: BORDER }} /></Field>
              <Field label="Cognome"><input value={form.cognome} onChange={(e) => setForm({ ...form, cognome: e.target.value })} className="px-2 py-1.5 rounded border w-full text-[13px]" style={{ borderColor: BORDER }} /></Field>
              <Field label="Data di nascita"><input type="date" value={form.dataNascita} onChange={(e) => setForm({ ...form, dataNascita: e.target.value })} className="px-2 py-1.5 rounded border w-full text-[13px]" style={{ borderColor: BORDER }} /></Field>
              <Field label="Livello"><input value={form.livello} onChange={(e) => setForm({ ...form, livello: e.target.value })} className="px-2 py-1.5 rounded border w-full text-[13px]" style={{ borderColor: BORDER }} /></Field>
              <Field label="Stato">
                <select value={form.stato} onChange={(e) => setForm({ ...form, stato: e.target.value })} className="px-2 py-1.5 rounded border w-full text-[13px]" style={{ borderColor: BORDER }}>
                  <option value="attivo">Attivo</option><option value="sospeso">Sospeso</option><option value="ritirato">Ritirato</option>
                </select>
              </Field>
              <Field label="Scadenza certificato"><input type="date" value={form.certificato} onChange={(e) => setForm({ ...form, certificato: e.target.value })} className="px-2 py-1.5 rounded border w-full text-[13px]" style={{ borderColor: BORDER }} /></Field>
              <div className="col-span-2">
                <Field label="Note"><textarea value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} className="px-2 py-1.5 rounded border w-full text-[13px]" style={{ borderColor: BORDER }} rows={2} /></Field>
              </div>
              <div className="col-span-2 flex gap-2">
                <button onClick={() => { onSave(form); setEdit(false); }} className="text-[13px] font-bold px-3 py-2 rounded-lg text-white" style={{ backgroundColor: PRIMARY }}>Salva</button>
                <button onClick={() => setEdit(false)} className="text-[13px] font-semibold px-3 py-2" style={{ color: INK_SOFT }}>Annulla</button>
              </div>
            </div>
          ) : (
            <>
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                <Stat label="Stato" value={student.stato} />
                <Stat label="Certificato" value={student.certificato || "—"} warn={certDays !== null && certDays <= 30} />
                <Stat label="Recuperi disponibili" value={rec.disponibili} />
                <Stat label="Assenze totali" value={rec.assenze} />
              </div>
              {student.note && <div className="text-[13px] p-2.5 rounded-lg" style={{ backgroundColor: BG, color: INK_SOFT }}>{student.note}</div>}
            </>
          )}

          <div>
            <SectionLabel>Corsi frequentati</SectionLabel>
            <div className="space-y-1.5">
              {studentCourses.map((c) => {
                const st = perCourseStats(c.id);
                return (
                  <div key={c.id} className="flex items-center justify-between px-3 py-2 rounded-lg" style={{ backgroundColor: BG }}>
                    <span className="text-[13px] font-semibold">{c.nome}</span>
                    <span className="text-[11px] flex gap-2" style={{ color: INK_SOFT }}>
                      <span style={{ color: PRESENTE }}>P {st.p}</span>
                      <span style={{ color: ASSENTE }}>A {st.a}</span>
                      <span style={{ color: RECUPERO }}>R {st.r}</span>
                      <span style={{ color: RITARDO }}>L {st.l}</span>
                    </span>
                  </div>
                );
              })}
              {studentCourses.length === 0 && <Empty>Nessun corso.</Empty>}
            </div>
          </div>

          <div>
            <SectionLabel>Recuperi</SectionLabel>
            <div className="text-[13px]" style={{ color: INK_SOFT }}>Maturati {rec.maturati} · Usati {rec.usati} · Disponibili <b style={{ color: RECUPERO }}>{rec.disponibili}</b></div>
          </div>

          {isAdmin && (
            <div>
              <SectionLabel>Valutazioni</SectionLabel>
              {studentEvals.length === 0 && <Empty>Nessuna valutazione.</Empty>}
              {studentEvals.map((e) => (
                <div key={e.id} className="flex items-center justify-between text-[13px] px-2 py-1">
                  <span>{e.criterio} · {e.periodo}</span>
                  <span className="flex">{[1, 2, 3, 4, 5].map((n) => <Star key={n} size={12} fill={n <= e.stelle ? "#E0A230" : "none"} color="#E0A230" />)}</span>
                </div>
              ))}
            </div>
          )}

          {isAdmin && (
            <div>
              <SectionLabel>Pagamenti</SectionLabel>
              {studentPayments.length === 0 && <Empty>Nessun pagamento registrato.</Empty>}
              {studentPayments.map((p) => (
                <div key={p.id} className="flex items-center justify-between text-[13px] px-2 py-1">
                  <span>{p.periodo} · €{p.importo}</span>
                  <StatoBadge status={p.stato === "Pagato" ? "P" : p.stato === "Scaduto" ? "A" : "L"} />
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function Field({ label, children }) {
  return <label className="block"><div className="text-[11px] font-semibold mb-0.5" style={{ color: INK_SOFT }}>{label}</div>{children}</label>;
}
function SectionLabel({ children }) {
  return <div className="font-bold text-[13px] mb-1.5" style={{ color: INK_SOFT }}>{children.toString().toUpperCase()}</div>;
}
function Stat({ label, value, warn }) {
  return (
    <div className="px-3 py-2 rounded-lg" style={{ backgroundColor: warn ? "#FBF1E0" : BG }}>
      <div className="text-[10px] font-semibold" style={{ color: INK_SOFT }}>{label}</div>
      <div className="text-[14px] font-bold capitalize" style={{ color: warn ? RITARDO : INK }}>{value}</div>
    </div>
  );
}

function CourseFormModal({ form, onCancel, onSave }) {
  const [data, setData] = useState(form.data);
  const section = SECTIONS.find((s) => s.key === data.sectionKey) || SECTIONS[0];

  function toggleGiorno(day) {
    setData((d) => d.giorni.includes(day) ? { ...d, giorni: d.giorni.filter((g) => g !== day) } : { ...d, giorni: [...d.giorni, day] });
  }
  function toggleCorsia(n) {
    setData((d) => d.corsie.includes(n) ? { ...d, corsie: d.corsie.filter((c) => c !== n) } : { ...d, corsie: [...d.corsie, n].sort((a, b) => a - b) });
  }
  const laneOptions = data.piscina === "Grande" ? [1, 2, 3, 4, 5, 6] : [1, 2];

  return (
    <div className="fixed inset-0 z-50 flex items-start sm:items-center justify-center" style={{ backgroundColor: "rgba(10,25,32,0.55)" }} onClick={onCancel}>
      <div onClick={(e) => e.stopPropagation()} className="w-full sm:max-w-lg sm:rounded-2xl bg-white h-full sm:h-auto sm:max-h-[90vh] overflow-auto" style={{ backgroundColor: SURFACE }}>
        <div className="sticky top-0 px-5 py-3.5 border-b font-extrabold text-lg" style={{ borderColor: BORDER, backgroundColor: SURFACE }}>
          {form.mode === "new" ? "Nuovo corso" : "Modifica corso"}
        </div>
        <div className="p-5 space-y-3">
          <Field label="Nome corso"><input value={data.nome} onChange={(e) => setData({ ...data, nome: e.target.value })} className="px-2 py-1.5 rounded border w-full text-[13px]" style={{ borderColor: BORDER }} /></Field>

          <Field label="Giorni">
            <div className="flex gap-1.5 flex-wrap">
              {section.days.map((d) => (
                <button key={d} onClick={() => toggleGiorno(d)} className="px-2.5 py-1 rounded-lg text-[12px] font-bold border"
                  style={{ backgroundColor: data.giorni.includes(d) ? PRIMARY : SURFACE, color: data.giorni.includes(d) ? "white" : INK, borderColor: BORDER }}>{d}</button>
              ))}
            </div>
          </Field>

          <div className="grid grid-cols-2 gap-2">
            <Field label="Ora inizio"><input type="time" value={data.ora} onChange={(e) => setData({ ...data, ora: e.target.value })} className="px-2 py-1.5 rounded border w-full text-[13px]" style={{ borderColor: BORDER }} /></Field>
            <Field label="Ora fine"><input type="time" value={data.oraFine} onChange={(e) => setData({ ...data, oraFine: e.target.value })} className="px-2 py-1.5 rounded border w-full text-[13px]" style={{ borderColor: BORDER }} /></Field>
          </div>

          <Field label="Piscina">
            <div className="flex gap-1.5">
              {["Grande", "Piccola"].map((p) => (
                <button key={p} onClick={() => setData({ ...data, piscina: p, corsie: [1] })} className="px-2.5 py-1 rounded-lg text-[12px] font-bold border"
                  style={{ backgroundColor: data.piscina === p ? PRIMARY : SURFACE, color: data.piscina === p ? "white" : INK, borderColor: BORDER }}>{p}</button>
              ))}
            </div>
          </Field>

          <Field label="Corsie">
            <div className="flex gap-1.5 flex-wrap">
              {laneOptions.map((n) => (
                <button key={n} onClick={() => toggleCorsia(n)} className="w-8 h-8 rounded-lg text-[12px] font-bold border"
                  style={{ backgroundColor: data.corsie.includes(n) ? ACCENT : SURFACE, color: data.corsie.includes(n) ? PRIMARY_DARK : INK, borderColor: BORDER }}>{n}</button>
              ))}
            </div>
          </Field>

          <div className="grid grid-cols-2 gap-2">
            <Field label="Livello"><input value={data.livello} onChange={(e) => setData({ ...data, livello: e.target.value })} className="px-2 py-1.5 rounded border w-full text-[13px]" style={{ borderColor: BORDER }} /></Field>
            <Field label="Istruttore"><input value={data.istruttore} onChange={(e) => setData({ ...data, istruttore: e.target.value })} className="px-2 py-1.5 rounded border w-full text-[13px]" style={{ borderColor: BORDER }} /></Field>
          </div>

          <div className="grid grid-cols-3 gap-2">
            <Field label="Capienza"><input type="number" value={data.capienza} onChange={(e) => setData({ ...data, capienza: parseInt(e.target.value, 10) || 0 })} className="px-2 py-1.5 rounded border w-full text-[13px]" style={{ borderColor: BORDER }} /></Field>
            <Field label="Inizio"><input type="date" value={data.inizio} onChange={(e) => setData({ ...data, inizio: e.target.value })} className="px-2 py-1.5 rounded border w-full text-[13px]" style={{ borderColor: BORDER }} /></Field>
            <Field label="Fine"><input type="date" value={data.fine} onChange={(e) => setData({ ...data, fine: e.target.value })} className="px-2 py-1.5 rounded border w-full text-[13px]" style={{ borderColor: BORDER }} /></Field>
          </div>

          <div className="flex gap-2 pt-2">
            <button onClick={() => onSave({ mode: form.mode, data })} disabled={!data.nome || data.giorni.length === 0}
              className="text-[13px] font-bold px-4 py-2 rounded-lg text-white disabled:opacity-40" style={{ backgroundColor: PRIMARY }}>Salva corso</button>
            <button onClick={onCancel} className="text-[13px] font-semibold px-3 py-2" style={{ color: INK_SOFT }}>Annulla</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function ImportExcelModal({ sheets, courses, onUpdateCourse, onConfirm, onCancel }) {
  const totalRows = sheets.reduce((n, s) => n + s.rows.length, 0);
  const assignedRows = sheets.filter((s) => s.courseId).reduce((n, s) => n + s.rows.length, 0);
  const canConfirm = sheets.some((s) => s.courseId);

  return (
    <div className="fixed inset-0 z-50 flex items-start sm:items-center justify-center" style={{ backgroundColor: "rgba(10,25,32,0.55)" }} onClick={onCancel}>
      <div onClick={(e) => e.stopPropagation()} className="w-full sm:max-w-xl sm:rounded-2xl bg-white h-full sm:h-auto sm:max-h-[90vh] overflow-auto" style={{ backgroundColor: SURFACE }}>
        <div className="sticky top-0 px-5 py-3.5 border-b flex items-center gap-2" style={{ borderColor: BORDER, backgroundColor: SURFACE }}>
          <FileSpreadsheet size={18} color={PRIMARY} />
          <div>
            <div className="font-extrabold text-lg leading-tight">Importa allievi da Excel</div>
            <div className="text-[12px]" style={{ color: INK_SOFT }}>{sheets.length} fogli trovati · {totalRows} allievi totali</div>
          </div>
        </div>

        <div className="p-5 space-y-3">
          <div className="text-[12px] p-2.5 rounded-lg" style={{ backgroundColor: BG, color: INK_SOFT }}>
            Per ogni foglio, scegli a quale corso associare gli allievi. Se il nome del foglio corrisponde a un corso esistente, l'abbinamento è già proposto: verificalo prima di confermare.
          </div>

          {sheets.map((s) => (
            <div key={s.sheetName} className="rounded-xl border p-3" style={{ borderColor: BORDER }}>
              <div className="flex items-center justify-between gap-2 mb-2">
                <div className="font-bold text-[13px]">{s.sheetName}</div>
                <span className="text-[11px]" style={{ color: INK_SOFT }}>{s.rows.length} righe</span>
              </div>
              <select value={s.courseId} onChange={(e) => onUpdateCourse(s.sheetName, e.target.value)}
                className="w-full px-2 py-1.5 rounded-lg border text-[13px] mb-2" style={{ borderColor: BORDER }}>
                <option value="">— Nessuno (salta questo foglio) —</option>
                {courses.map((c) => <option key={c.id} value={c.id}>{c.nome}</option>)}
              </select>
              <div className="text-[11px] space-y-0.5" style={{ color: INK_SOFT }}>
                {s.rows.slice(0, 3).map((r, i) => (
                  <div key={i}>• {String(r.nome || "").trim()} {String(r.cognome || "").trim()}</div>
                ))}
                {s.rows.length > 3 && <div>+ altri {s.rows.length - 3}</div>}
              </div>
            </div>
          ))}

          <div className="flex items-center gap-2 pt-2">
            <button onClick={onConfirm} disabled={!canConfirm}
              className="text-[13px] font-bold px-4 py-2 rounded-lg text-white disabled:opacity-40" style={{ backgroundColor: PRIMARY }}>
              Importa {assignedRows} allievi
            </button>
            <button onClick={onCancel} className="text-[13px] font-semibold px-3 py-2" style={{ color: INK_SOFT }}>Annulla</button>
          </div>
        </div>
      </div>
    </div>
  );
}
