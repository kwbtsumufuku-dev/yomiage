// main-esm.jsx
import React2 from "react";
import { createRoot } from "react-dom/client";

// standalone-app.jsx
import React, { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import {
  Upload,
  Play,
  Pause,
  SkipBack,
  SkipForward,
  Plus,
  Trash2,
  ChevronUp,
  ChevronDown,
  BookOpen,
  Volume2,
  FileText,
  RotateCcw,
  AlertCircle,
  CheckCircle2,
  X,
  Maximize2,
  Minimize2,
  ArrowLeft,
  Printer,
  History,
  ChevronRight,
  Building2,
  Settings,
  ToggleLeft,
  ToggleRight
} from "lucide-react";
var DEFAULT_DOC_TYPES = [
  { id: "important", label: "\u91CD\u8981\u4E8B\u9805\u8AAC\u660E\u66F8", deletable: false },
  { id: "consent", label: "\u500B\u4EBA\u60C5\u5831\u4F7F\u7528\u540C\u610F\u66F8", deletable: false },
  { id: "contract", label: "\u5951\u7D04\u66F8", deletable: false }
];
var HEADING_PATTERN = /^[0-9０-９]{1,2}[　\s.．]/;
var MAX_VERSIONS_PER_DOC = 5;
var DEFAULT_DICT_ENTRIES = {
  "\u6240\u5728\u5730": "\u3057\u3087\u3056\u3044\u3061",
  "\u66F8": "\u3057\u3087"
};
var DIGIT_READING = {
  "0": "\u30BC\u30ED",
  "1": "\u3044\u3061",
  "2": "\u306B",
  "3": "\u3055\u3093",
  "4": "\u3088\u3093",
  "5": "\u3054",
  "6": "\u308D\u304F",
  "7": "\u306A\u306A",
  "8": "\u306F\u3061",
  "9": "\u304D\u3085\u3046",
  "\uFF10": "\u30BC\u30ED",
  "\uFF11": "\u3044\u3061",
  "\uFF12": "\u306B",
  "\uFF13": "\u3055\u3093",
  "\uFF14": "\u3088\u3093",
  "\uFF15": "\u3054",
  "\uFF16": "\u308D\u304F",
  "\uFF17": "\u306A\u306A",
  "\uFF18": "\u306F\u3061",
  "\uFF19": "\u304D\u3085\u3046"
};
var NUMBER_HYPHEN_PATTERN = /[0-9０-９]+(?:[-‐－ー–—][0-9０-９]+)+/g;
var NUMBER_LABEL_PATTERN = new RegExp(
  "(\u756A\u53F7|\u30B3\u30FC\u30C9|\uFF29\uFF24|ID|\u90F5\u4FBF|\uFF34\uFF25\uFF2C|TEL|\u96FB\u8A71|\uFF26\uFF21\uFF38|FAX)([^0-9\uFF10-\uFF19]{0,6})([0-9\uFF10-\uFF19]{3,}(?:[-\u2010\uFF0D\u30FC\u2013\u2014][0-9\uFF10-\uFF19]+)*)",
  "g"
);
function toDigitByDigitReading(numStr) {
  return Array.from(numStr).map((ch) => DIGIT_READING[ch] || ch).join("");
}
function applyLabeledNumberReading(text) {
  return text.replace(NUMBER_LABEL_PATTERN, (match, label, sep, numPart) => {
    const groups = numPart.split(/[-‐－ー–—]/);
    return label + sep + groups.map(toDigitByDigitReading).join("\u306E");
  });
}
function applyPhoneAndAddressNumberReading(text) {
  return text.replace(NUMBER_HYPHEN_PATTERN, (match) => {
    const groups = match.split(/[-‐－ー–—]/);
    return groups.map(toDigitByDigitReading).join("\u306E");
  });
}
function applyNumberReadingRules(text) {
  return applyPhoneAndAddressNumberReading(applyLabeledNumberReading(text));
}
function applyDevRules(text, rules) {
  let out = text;
  (rules || []).forEach((r) => {
    if (!r || r.enabled === false || !r.pattern) return;
    try {
      const flags = "g" + (r.caseInsensitive ? "i" : "");
      const re = new RegExp(r.pattern, flags);
      out = out.replace(re, r.replacement ?? "");
    } catch (e) {
    }
  });
  return out;
}
var idCounter = 1;
var nextId = () => `s${idCounter++}`;
var genId = (prefix) => `${prefix}_${Date.now().toString(36)}_${Math.floor(Math.random() * 1e4)}`;
function storageKey(suffix) {
  return `juuyoujikou:${suffix}`;
}
var LOCAL_STORAGE_PREFIX = "juuyoujikouYomiage:";
async function loadJSON(key, fallback) {
  try {
    const raw = window.localStorage.getItem(LOCAL_STORAGE_PREFIX + key);
    if (raw === null) return fallback;
    return JSON.parse(raw);
  } catch (e) {
    return fallback;
  }
}
async function saveJSON(key, value) {
  try {
    window.localStorage.setItem(LOCAL_STORAGE_PREFIX + key, JSON.stringify(value));
    return true;
  } catch (e) {
    console.error("\u4FDD\u5B58\u306B\u5931\u6557\u3057\u307E\u3057\u305F", e);
    return false;
  }
}
function formatDate(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return "";
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}/${pad(d.getMonth() + 1)}/${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
async function parseDocx(arrayBuffer) {
  const result = await mammoth.convertToHtml({ arrayBuffer });
  const html = result.value;
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, "text/html");
  const chunks = [];
  const pushText = (text, isHeading) => {
    const trimmed = text.replace(/\s+/g, " ").trim();
    if (!trimmed) return;
    chunks.push({
      id: nextId(),
      isHeading: !!isHeading,
      text: trimmed
    });
  };
  Array.from(doc.body.children).forEach((el) => {
    const tag = el.tagName.toLowerCase();
    if (/^h[1-6]$/.test(tag)) {
      pushText(el.textContent, true);
      return;
    }
    if (tag === "p") {
      const text = el.textContent || "";
      const isHeading = HEADING_PATTERN.test(text.trim());
      pushText(text, isHeading);
      return;
    }
    if (tag === "table") {
      const rows = Array.from(el.querySelectorAll("tr"));
      rows.forEach((tr) => {
        const cells = Array.from(tr.querySelectorAll("td,th")).map((td) => (td.textContent || "").replace(/\s+/g, " ").trim()).filter(Boolean);
        if (cells.length === 0) return;
        if (cells.length === 1) {
          pushText(cells[0], false);
        } else if (cells.length === 2) {
          pushText(`${cells[0]}\uFF1A${cells[1]}`, false);
        } else {
          pushText(cells.join(" / "), false);
        }
      });
      return;
    }
    if (tag === "ul" || tag === "ol") {
      Array.from(el.querySelectorAll("li")).forEach((li) => {
        pushText("\u30FB" + (li.textContent || ""), false);
      });
    }
  });
  return chunks;
}
async function parseXlsx(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });
  const chunks = [];
  let headerRow = null;
  rows.forEach((row) => {
    const cells = row.map((c) => c == null ? "" : String(c).trim()).filter(Boolean);
    if (cells.length === 0) return;
    if (!headerRow && cells.length > 1) {
      headerRow = row.map((c) => c == null ? "" : String(c).trim());
      return;
    }
    if (cells.length === 1) {
      chunks.push({ id: nextId(), isHeading: HEADING_PATTERN.test(cells[0]), text: cells[0] });
      return;
    }
    if (headerRow) {
      const parts = [];
      row.forEach((val, i) => {
        const v = val == null ? "" : String(val).trim();
        const label = headerRow[i] ? String(headerRow[i]).trim() : "";
        if (v && label) parts.push(`${label}\uFF1A${v}`);
        else if (v) parts.push(v);
      });
      if (parts.length) chunks.push({ id: nextId(), isHeading: false, text: parts.join(" / ") });
    } else {
      chunks.push({ id: nextId(), isHeading: false, text: cells.join(" / ") });
    }
  });
  return chunks;
}
function applyDict(text, dict) {
  const terms = Object.keys(dict).filter((t) => t && dict[t] !== void 0 && dict[t] !== null).sort((a, b) => b.length - a.length);
  let out = text;
  terms.forEach((term) => {
    const escaped = term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    out = out.split(new RegExp(escaped, "g")).join(dict[term]);
  });
  return out;
}
function wrapWords(text) {
  const chars = Array.from(text);
  const spans = [];
  let buf = "";
  chars.forEach((ch) => {
    buf += ch;
    if ("\u3001\u3002".includes(ch)) {
      spans.push(buf);
      buf = "";
    }
  });
  if (buf) spans.push(buf);
  return spans;
}
function JuuyoujikouYomiageApp() {
  const [loading, setLoading] = useState(true);
  const [toast, setToast] = useState(null);
  const [services, setServices] = useState([]);
  const [docTypeDefs, setDocTypeDefs] = useState(DEFAULT_DOC_TYPES);
  const [docs, setDocs] = useState({});
  const [dict, setDict] = useState({});
  const [devRules, setDevRules] = useState([]);
  const [screen, setScreen] = useState("menu");
  const [selectedServiceId, setSelectedServiceId] = useState(null);
  const [activeDocType, setActiveDocType] = useState(null);
  const [activeVersionMap, setActiveVersionMap] = useState({});
  const [newServiceName, setNewServiceName] = useState("");
  const [showAddDocType, setShowAddDocType] = useState(false);
  const [newDocTypeName, setNewDocTypeName] = useState("");
  const [dictTermInput, setDictTermInput] = useState("");
  const [dictReadingInput, setDictReadingInput] = useState("");
  const [dictSkipInput, setDictSkipInput] = useState(false);
  const [queue, setQueue] = useState([]);
  const [current, setCurrent] = useState(0);
  const [isSpeaking, setIsSpeaking] = useState(false);
  const [rate, setRate] = useState(1);
  const [autoAdvance, setAutoAdvance] = useState(true);
  const [jaVoice, setJaVoice] = useState(null);
  const [wordSpans, setWordSpans] = useState([]);
  const [activeWordIdx, setActiveWordIdx] = useState(-1);
  const utteranceRef = useRef(null);
  const [printTarget, setPrintTarget] = useState(null);
  const [confirmDialog, setConfirmDialog] = useState(null);
  const [editingServiceId, setEditingServiceId] = useState(null);
  const [editingServiceName, setEditingServiceName] = useState("");
  const showToast = (msg, type = "info") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3200);
  };
  useEffect(() => {
    let cancelled = false;
    (async () => {
      setLoading(true);
      const svc = await loadJSON(storageKey("services"), []);
      const dtDefs = await loadJSON(storageKey("docTypeDefs"), DEFAULT_DOC_TYPES);
      const docsData = await loadJSON(storageKey("docs"), {});
      let dictData = await loadJSON(storageKey("dict"), {});
      const devRulesData = await loadJSON(storageKey("devRules"), []);
      const seeded = await loadJSON(storageKey("dictSeeded"), false);
      if (!seeded) {
        dictData = { ...DEFAULT_DICT_ENTRIES, ...dictData || {} };
        await saveJSON(storageKey("dict"), dictData);
        await saveJSON(storageKey("dictSeeded"), true);
      }
      if (!cancelled) {
        setServices(svc);
        setDocTypeDefs(dtDefs && dtDefs.length ? dtDefs : DEFAULT_DOC_TYPES);
        setDocs(docsData || {});
        setDict(dictData || {});
        setDevRules(devRulesData || []);
        setLoading(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);
  useEffect(() => {
    if (!("speechSynthesis" in window)) return;
    const pick = () => {
      const voices = window.speechSynthesis.getVoices();
      const v = voices.find((v2) => v2.lang && v2.lang.toLowerCase().startsWith("ja"));
      setJaVoice(v || null);
    };
    window.speechSynthesis.onvoiceschanged = pick;
    pick();
  }, []);
  useEffect(() => {
    if (screen !== "menu" && selectedServiceId && !services.some((s) => s.id === selectedServiceId) && !loading) {
      setScreen("menu");
      setSelectedServiceId(null);
    }
  }, [screen, selectedServiceId, services, loading]);
  const persistServices = async (next) => {
    const prev = services;
    setServices(next);
    const ok = await saveJSON(storageKey("services"), next);
    if (!ok) {
      setServices(prev);
      showToast("\u4FDD\u5B58\u306B\u5931\u6557\u3057\u307E\u3057\u305F\u3002\u4FDD\u5B58\u5BB9\u91CF\u306E\u4E0A\u9650\u306B\u9054\u3057\u3066\u3044\u308B\u53EF\u80FD\u6027\u304C\u3042\u308A\u307E\u3059\u3002\u4E0D\u8981\u306A\u66F8\u985E\u30D0\u30FC\u30B8\u30E7\u30F3\u3092\u524A\u9664\u3057\u3066\u304B\u3089\u3001\u3082\u3046\u4E00\u5EA6\u304A\u8A66\u3057\u304F\u3060\u3055\u3044\u3002", "error");
    }
    return ok;
  };
  const persistDocTypeDefs = async (next) => {
    const prev = docTypeDefs;
    setDocTypeDefs(next);
    const ok = await saveJSON(storageKey("docTypeDefs"), next);
    if (!ok) {
      setDocTypeDefs(prev);
      showToast("\u4FDD\u5B58\u306B\u5931\u6557\u3057\u307E\u3057\u305F\u3002\u4FDD\u5B58\u5BB9\u91CF\u306E\u4E0A\u9650\u306B\u9054\u3057\u3066\u3044\u308B\u53EF\u80FD\u6027\u304C\u3042\u308A\u307E\u3059\u3002\u4E0D\u8981\u306A\u66F8\u985E\u30D0\u30FC\u30B8\u30E7\u30F3\u3092\u524A\u9664\u3057\u3066\u304B\u3089\u3001\u3082\u3046\u4E00\u5EA6\u304A\u8A66\u3057\u304F\u3060\u3055\u3044\u3002", "error");
    }
    return ok;
  };
  const persistDocs = async (next) => {
    const prev = docs;
    setDocs(next);
    const ok = await saveJSON(storageKey("docs"), next);
    if (!ok) {
      setDocs(prev);
      showToast("\u4FDD\u5B58\u306B\u5931\u6557\u3057\u307E\u3057\u305F\u3002\u4FDD\u5B58\u5BB9\u91CF\u306E\u4E0A\u9650\u306B\u9054\u3057\u3066\u3044\u308B\u53EF\u80FD\u6027\u304C\u3042\u308A\u307E\u3059\u3002\u4E0D\u8981\u306A\u66F8\u985E\u30D0\u30FC\u30B8\u30E7\u30F3\u3092\u524A\u9664\u3057\u3066\u304B\u3089\u3001\u3082\u3046\u4E00\u5EA6\u304A\u8A66\u3057\u304F\u3060\u3055\u3044\u3002", "error");
    }
    return ok;
  };
  const persistDict = async (next) => {
    const prev = dict;
    setDict(next);
    const ok = await saveJSON(storageKey("dict"), next);
    if (!ok) {
      setDict(prev);
      showToast("\u4FDD\u5B58\u306B\u5931\u6557\u3057\u307E\u3057\u305F\u3002\u4FDD\u5B58\u5BB9\u91CF\u306E\u4E0A\u9650\u306B\u9054\u3057\u3066\u3044\u308B\u53EF\u80FD\u6027\u304C\u3042\u308A\u307E\u3059\u3002\u4E0D\u8981\u306A\u66F8\u985E\u30D0\u30FC\u30B8\u30E7\u30F3\u3092\u524A\u9664\u3057\u3066\u304B\u3089\u3001\u3082\u3046\u4E00\u5EA6\u304A\u8A66\u3057\u304F\u3060\u3055\u3044\u3002", "error");
    }
    return ok;
  };
  const persistDevRules = async (next) => {
    const prev = devRules;
    setDevRules(next);
    const ok = await saveJSON(storageKey("devRules"), next);
    if (!ok) {
      setDevRules(prev);
      showToast("\u4FDD\u5B58\u306B\u5931\u6557\u3057\u307E\u3057\u305F\u3002\u4FDD\u5B58\u5BB9\u91CF\u306E\u4E0A\u9650\u306B\u9054\u3057\u3066\u3044\u308B\u53EF\u80FD\u6027\u304C\u3042\u308A\u307E\u3059\u3002", "error");
    }
    return ok;
  };
  const addDevRule = (rule) => {
    if (!rule.pattern || !rule.pattern.trim()) return false;
    try {
      new RegExp(rule.pattern, "g" + (rule.caseInsensitive ? "i" : ""));
    } catch (e) {
      showToast("\u6B63\u898F\u8868\u73FE\u304C\u6B63\u3057\u304F\u3042\u308A\u307E\u305B\u3093\u3002\u30D1\u30BF\u30FC\u30F3\u3092\u898B\u76F4\u3057\u3066\u304F\u3060\u3055\u3044\u3002", "error");
      return false;
    }
    const next = {
      id: genId("devrule"),
      label: (rule.label || "").trim(),
      pattern: rule.pattern,
      replacement: rule.replacement ?? "",
      caseInsensitive: !!rule.caseInsensitive,
      enabled: true
    };
    persistDevRules([...devRules, next]);
    return true;
  };
  const updateDevRule = (id, patch) => {
    persistDevRules(devRules.map((r) => r.id === id ? { ...r, ...patch } : r));
  };
  const deleteDevRule = (id) => {
    persistDevRules(devRules.filter((r) => r.id !== id));
  };
  const addService = () => {
    const name = newServiceName.trim();
    if (!name) return;
    const svc = { id: genId("svc"), name };
    persistServices([...services, svc]);
    setNewServiceName("");
    showToast(`\u300C${name}\u300D\u3092\u8FFD\u52A0\u3057\u307E\u3057\u305F`, "success");
  };
  const startRenameService = (id) => {
    const svc = services.find((s) => s.id === id);
    if (!svc) return;
    setEditingServiceId(id);
    setEditingServiceName(svc.name);
  };
  const cancelRenameService = () => {
    setEditingServiceId(null);
    setEditingServiceName("");
  };
  const saveRenameService = () => {
    const trimmed = editingServiceName.trim();
    if (!trimmed) {
      cancelRenameService();
      return;
    }
    persistServices(services.map((s) => s.id === editingServiceId ? { ...s, name: trimmed } : s));
    cancelRenameService();
  };
  const deleteService = (id) => {
    const svc = services.find((s) => s.id === id);
    if (!svc) return;
    setConfirmDialog({
      message: `\u300C${svc.name}\u300D\u3092\u524A\u9664\u3057\u307E\u3059\u304B\uFF1F
\u4FDD\u5B58\u3055\u308C\u3066\u3044\u308B\u66F8\u985E\u3082\u3059\u3079\u3066\u524A\u9664\u3055\u308C\u307E\u3059\u3002`,
      onConfirm: () => {
        persistServices(services.filter((s) => s.id !== id));
        const nextDocs = { ...docs };
        delete nextDocs[id];
        persistDocs(nextDocs);
        showToast(`\u300C${svc.name}\u300D\u3092\u524A\u9664\u3057\u307E\u3057\u305F`, "success");
        setConfirmDialog(null);
      }
    });
  };
  const addDocType = () => {
    const label = newDocTypeName.trim();
    if (!label) return;
    const dt = { id: genId("doctype"), label, deletable: true };
    persistDocTypeDefs([...docTypeDefs, dt]);
    setNewDocTypeName("");
    setShowAddDocType(false);
    setActiveDocType(dt.id);
    showToast(`\u300C${label}\u300D\u3092\u66F8\u985E\u306E\u7A2E\u985E\u306B\u8FFD\u52A0\u3057\u307E\u3057\u305F`, "success");
  };
  const deleteDocType = (id) => {
    const dt = docTypeDefs.find((d) => d.id === id);
    if (!dt || !dt.deletable) return;
    setConfirmDialog({
      message: `\u300C${dt.label}\u300D\u3092\u524A\u9664\u3057\u307E\u3059\u304B\uFF1F
\u3059\u3079\u3066\u306E\u30B5\u30FC\u30D3\u30B9\u3067\u4FDD\u5B58\u3055\u308C\u3066\u3044\u308B\u8A72\u5F53\u66F8\u985E\u3082\u524A\u9664\u3055\u308C\u307E\u3059\u3002`,
      onConfirm: () => {
        const remaining = docTypeDefs.filter((d) => d.id !== id);
        persistDocTypeDefs(remaining);
        const nextDocs = {};
        Object.entries(docs).forEach(([svcId, byType]) => {
          const rest = { ...byType };
          delete rest[id];
          nextDocs[svcId] = rest;
        });
        persistDocs(nextDocs);
        if (activeDocType === id) setActiveDocType(remaining[0]?.id || "dict");
        showToast(`\u300C${dt.label}\u300D\u3092\u524A\u9664\u3057\u307E\u3057\u305F`, "success");
        setConfirmDialog(null);
      }
    });
  };
  const getVersions = (serviceId, docTypeId) => {
    const arr = docs[serviceId] && docs[serviceId][docTypeId] || [];
    return [...arr].sort((a, b) => (b.uploadedAt || "").localeCompare(a.uploadedAt || ""));
  };
  const getActiveVersionId = (serviceId, docTypeId) => {
    const versions = getVersions(serviceId, docTypeId);
    if (!versions.length) return null;
    const key = `${serviceId}:${docTypeId}`;
    const chosen = activeVersionMap[key];
    if (chosen && versions.some((v) => v.id === chosen)) return chosen;
    return versions[0].id;
  };
  const getActiveVersion = (serviceId, docTypeId) => {
    const versions = getVersions(serviceId, docTypeId);
    const vid = getActiveVersionId(serviceId, docTypeId);
    return versions.find((v) => v.id === vid) || null;
  };
  const selectVersion = (serviceId, docTypeId, versionId) => {
    const key = `${serviceId}:${docTypeId}`;
    setActiveVersionMap((prev) => ({ ...prev, [key]: versionId }));
  };
  const updateVersions = async (serviceId, docTypeId, updater) => {
    const currentList = docs[serviceId] && docs[serviceId][docTypeId] || [];
    const nextVersions = updater(currentList);
    const nextDocs = {
      ...docs,
      [serviceId]: {
        ...docs[serviceId] || {},
        [docTypeId]: nextVersions
      }
    };
    return persistDocs(nextDocs);
  };
  const handleFile = async (serviceId, docTypeId, file) => {
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "pdf") {
      showToast("PDF\u306F\u975E\u5BFE\u5FDC\u3067\u3059\u3002Word\uFF08.docx\uFF09\u307E\u305F\u306FExcel\uFF08.xlsx\uFF09\u3092\u9078\u3093\u3067\u304F\u3060\u3055\u3044\u3002", "error");
      return;
    }
    if (!["docx", "xlsx"].includes(ext)) {
      showToast("\u5BFE\u5FDC\u3057\u3066\u3044\u306A\u3044\u30D5\u30A1\u30A4\u30EB\u5F62\u5F0F\u3067\u3059\uFF08.docx / .xlsx \u306E\u307F\uFF09", "error");
      return;
    }
    try {
      const buf = await file.arrayBuffer();
      let sections;
      if (ext === "docx") sections = await parseDocx(buf);
      else sections = await parseXlsx(buf);
      if (sections.length === 0) {
        showToast("\u6587\u7AE0\u3092\u62BD\u51FA\u3067\u304D\u307E\u305B\u3093\u3067\u3057\u305F\u3002\u30D5\u30A1\u30A4\u30EB\u306E\u4E2D\u8EAB\u3092\u3054\u78BA\u8A8D\u304F\u3060\u3055\u3044\u3002", "error");
        return;
      }
      const version = {
        id: genId("ver"),
        fileName: file.name,
        uploadedAt: (/* @__PURE__ */ new Date()).toISOString(),
        sections
      };
      const existingCount = (docs[serviceId] && docs[serviceId][docTypeId] || []).length;
      const willTrim = existingCount + 1 > MAX_VERSIONS_PER_DOC;
      const ok = await updateVersions(serviceId, docTypeId, (prev) => {
        const merged = [...prev, version];
        const sorted = [...merged].sort((a, b) => (b.uploadedAt || "").localeCompare(a.uploadedAt || ""));
        return sorted.slice(0, MAX_VERSIONS_PER_DOC);
      });
      if (ok) {
        selectVersion(serviceId, docTypeId, version.id);
        showToast(
          willTrim ? `${sections.length}\u4EF6\u306E\u9805\u76EE\u3092\u8AAD\u307F\u8FBC\u307F\u307E\u3057\u305F\uFF08\u65B0\u3057\u3044\u30D0\u30FC\u30B8\u30E7\u30F3\u3068\u3057\u3066\u4FDD\u5B58\uFF09\u3002\u4FDD\u5B58\u3067\u304D\u308B\u306E\u306F\u6700\u65B0+\u904E\u53BB4\u4EF6\u307E\u3067\u306E\u305F\u3081\u3001\u4E00\u756A\u53E4\u3044\u30D0\u30FC\u30B8\u30E7\u30F3\u306F\u81EA\u52D5\u7684\u306B\u524A\u9664\u3055\u308C\u307E\u3057\u305F\u3002` : `${sections.length}\u4EF6\u306E\u9805\u76EE\u3092\u8AAD\u307F\u8FBC\u307F\u307E\u3057\u305F\uFF08\u65B0\u3057\u3044\u30D0\u30FC\u30B8\u30E7\u30F3\u3068\u3057\u3066\u4FDD\u5B58\uFF09\u3002\u5185\u5BB9\u3092\u78BA\u8A8D\u30FB\u7DE8\u96C6\u3057\u3066\u304F\u3060\u3055\u3044\u3002`,
          "success"
        );
      }
    } catch (e) {
      console.error(e);
      showToast("\u8AAD\u307F\u8FBC\u307F\u4E2D\u306B\u30A8\u30E9\u30FC\u304C\u767A\u751F\u3057\u307E\u3057\u305F\u3002", "error");
    }
  };
  const createBlankVersion = (serviceId, docTypeId) => {
    const version = { id: genId("ver"), fileName: "\uFF08\u624B\u5165\u529B\u3067\u4F5C\u6210\uFF09", uploadedAt: (/* @__PURE__ */ new Date()).toISOString(), sections: [] };
    updateVersions(serviceId, docTypeId, (prev) => {
      const merged = [...prev, version];
      const sorted = [...merged].sort((a, b) => (b.uploadedAt || "").localeCompare(a.uploadedAt || ""));
      return sorted.slice(0, MAX_VERSIONS_PER_DOC);
    });
    selectVersion(serviceId, docTypeId, version.id);
  };
  const updateSectionInVersion = (serviceId, docTypeId, versionId, sectionId, text) => {
    updateVersions(serviceId, docTypeId, (prev) => prev.map((v) => v.id === versionId ? { ...v, sections: v.sections.map((s) => s.id === sectionId ? { ...s, text } : s) } : v));
  };
  const toggleHeadingInVersion = (serviceId, docTypeId, versionId, sectionId) => {
    updateVersions(serviceId, docTypeId, (prev) => prev.map((v) => v.id === versionId ? { ...v, sections: v.sections.map((s) => s.id === sectionId ? { ...s, isHeading: !s.isHeading } : s) } : v));
  };
  const deleteSectionInVersion = (serviceId, docTypeId, versionId, sectionId) => {
    updateVersions(serviceId, docTypeId, (prev) => prev.map((v) => v.id === versionId ? { ...v, sections: v.sections.filter((s) => s.id !== sectionId) } : v));
  };
  const addSectionInVersion = (serviceId, docTypeId, versionId) => {
    updateVersions(serviceId, docTypeId, (prev) => prev.map((v) => v.id === versionId ? { ...v, sections: [...v.sections, { id: nextId(), isHeading: false, text: "" }] } : v));
  };
  const moveSectionInVersion = (serviceId, docTypeId, versionId, sectionId, dir) => {
    updateVersions(serviceId, docTypeId, (prev) => prev.map((v) => {
      if (v.id !== versionId) return v;
      const sections = [...v.sections];
      const idx = sections.findIndex((s) => s.id === sectionId);
      const swapWith = idx + dir;
      if (idx < 0 || swapWith < 0 || swapWith >= sections.length) return v;
      [sections[idx], sections[swapWith]] = [sections[swapWith], sections[idx]];
      return { ...v, sections };
    }));
  };
  const deleteVersion = (serviceId, docTypeId, versionId) => {
    setConfirmDialog({
      message: "\u3053\u306E\u30D0\u30FC\u30B8\u30E7\u30F3\u3092\u524A\u9664\u3057\u307E\u3059\u304B\uFF1F\n\u5143\u306B\u623B\u305B\u307E\u305B\u3093\u3002",
      onConfirm: () => {
        updateVersions(serviceId, docTypeId, (prev) => prev.filter((v) => v.id !== versionId));
        const key = `${serviceId}:${docTypeId}`;
        setActiveVersionMap((prev) => {
          const next = { ...prev };
          if (next[key] === versionId) delete next[key];
          return next;
        });
        setConfirmDialog(null);
      }
    });
  };
  const addDictEntry = () => {
    const term = dictTermInput.trim();
    const reading = dictSkipInput ? "" : dictReadingInput.trim();
    if (!term) return;
    if (!dictSkipInput && !reading) return;
    const next = { ...dict, [term]: reading };
    persistDict(next);
    setDictTermInput("");
    setDictReadingInput("");
    setDictSkipInput(false);
    showToast(
      dictSkipInput ? `\u300C${term}\u300D\u3092\u8AAD\u307F\u4E0A\u3052\u304B\u3089\u9664\u5916\u3057\u307E\u3057\u305F` : `\u300C${term}\u300D\u2192\u300C${reading}\u300D\u3092\u767B\u9332\u3057\u307E\u3057\u305F`,
      "success"
    );
  };
  const removeDictEntry = (term) => {
    const next = { ...dict };
    delete next[term];
    persistDict(next);
  };
  const buildQueue = (serviceId, serviceName) => {
    const q = [];
    docTypeDefs.forEach(({ id, label }) => {
      const version = getActiveVersion(serviceId, id);
      if (!version) return;
      const secs = version.sections.filter((s) => s.text.trim());
      secs.forEach((s) => q.push({ docLabel: `${serviceName}\uFF5C${label}`, ...s }));
    });
    return q;
  };
  const goToPlayback = () => {
    if (!selectedServiceId) return;
    const service2 = services.find((s) => s.id === selectedServiceId);
    const q = buildQueue(selectedServiceId, service2 ? service2.name : "");
    if (q.length === 0) {
      showToast("\u8AAD\u307F\u4E0A\u3052\u308B\u9805\u76EE\u304C\u3042\u308A\u307E\u305B\u3093\u3002\u5148\u306B\u66F8\u985E\u3092\u8AAD\u307F\u8FBC\u3093\u3067\u304F\u3060\u3055\u3044\u3002", "error");
      return;
    }
    setQueue(q);
    setCurrent(0);
    setScreen("play");
  };
  const stopSpeaking = () => {
    if ("speechSynthesis" in window) window.speechSynthesis.cancel();
    setIsSpeaking(false);
    setActiveWordIdx(-1);
  };
  const speakItem = (index) => {
    if (!("speechSynthesis" in window) || !queue[index]) return;
    window.speechSynthesis.cancel();
    const item = queue[index];
    const displaySpans = wrapWords(item.text);
    setWordSpans(displaySpans);
    setActiveWordIdx(-1);
    const spokenText = applyDict(applyDevRules(applyNumberReadingRules(item.text), devRules), dict);
    const utter = new SpeechSynthesisUtterance(spokenText);
    utter.lang = "ja-JP";
    if (jaVoice) utter.voice = jaVoice;
    utter.rate = rate;
    utter.onboundary = (e) => {
      if (typeof e.charIndex === "number" && spokenText.length > 0) {
        const fraction = e.charIndex / spokenText.length;
        const idx = Math.min(displaySpans.length - 1, Math.floor(fraction * displaySpans.length));
        setActiveWordIdx(idx);
      }
    };
    utter.onstart = () => setIsSpeaking(true);
    utter.onend = () => {
      setIsSpeaking(false);
      setActiveWordIdx(-1);
      if (autoAdvance && index < queue.length - 1) {
        setCurrent(index + 1);
        setTimeout(() => speakItem(index + 1), 450);
      }
    };
    utteranceRef.current = utter;
    window.speechSynthesis.speak(utter);
  };
  const togglePlay = () => {
    if (!("speechSynthesis" in window)) return;
    if (isSpeaking) {
      stopSpeaking();
    } else {
      speakItem(current);
    }
  };
  const goPrev = () => {
    stopSpeaking();
    setCurrent((c) => Math.max(0, c - 1));
  };
  const goNext = () => {
    stopSpeaking();
    setCurrent((c) => Math.min(queue.length - 1, c + 1));
  };
  const restart = () => {
    stopSpeaking();
    setCurrent(0);
  };
  useEffect(() => {
    setWordSpans(queue[current] ? wrapWords(queue[current].text) : []);
    setActiveWordIdx(-1);
  }, [current, queue]);
  useEffect(() => {
    if (screen !== "play") {
      if ("speechSynthesis" in window) window.speechSynthesis.cancel();
      setIsSpeaking(false);
      setActiveWordIdx(-1);
    }
  }, [screen]);
  const openPrint = (docTypeId, versionId) => {
    setPrintTarget({ serviceId: selectedServiceId, docTypeId, versionId });
    setScreen("print");
  };
  if (loading) {
    return /* @__PURE__ */ React.createElement("div", { className: "min-h-screen bg-stone-50 flex items-center justify-center" }, /* @__PURE__ */ React.createElement("div", { className: "text-stone-500" }, "\u8AAD\u307F\u8FBC\u307F\u4E2D..."));
  }
  if (screen === "play") {
    return /* @__PURE__ */ React.createElement(
      PlaybackPanel,
      {
        queue,
        current,
        isSpeaking,
        wordSpans,
        activeWordIdx,
        rate,
        setRate,
        autoAdvance,
        setAutoAdvance,
        onTogglePlay: togglePlay,
        onPrev: goPrev,
        onNext: goNext,
        onRestart: restart,
        onBack: () => {
          stopSpeaking();
          setScreen("workspace");
        }
      }
    );
  }
  if (screen === "print" && printTarget) {
    const version = ((docs[printTarget.serviceId] || {})[printTarget.docTypeId] || []).find((v) => v.id === printTarget.versionId);
    return /* @__PURE__ */ React.createElement(
      PrintView,
      {
        service: services.find((s) => s.id === printTarget.serviceId),
        docType: docTypeDefs.find((d) => d.id === printTarget.docTypeId),
        version,
        onClose: () => {
          setPrintTarget(null);
          setScreen("workspace");
        }
      }
    );
  }
  if (screen === "devSettings") {
    return /* @__PURE__ */ React.createElement(
      DevSettingsScreen,
      {
        devRules,
        dict,
        onAdd: addDevRule,
        onUpdate: updateDevRule,
        onDelete: deleteDevRule,
        onBack: () => setScreen("menu")
      }
    );
  }
  if (screen === "menu" || !selectedServiceId) {
    return /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement(
      MenuScreen,
      {
        services,
        newServiceName,
        setNewServiceName,
        onAdd: addService,
        editingServiceId,
        editingServiceName,
        setEditingServiceName,
        onStartRename: startRenameService,
        onSaveRename: saveRenameService,
        onCancelRename: cancelRenameService,
        onDelete: deleteService,
        onSelect: (id) => {
          setSelectedServiceId(id);
          setActiveDocType(docTypeDefs[0] ? docTypeDefs[0].id : "dict");
          setScreen("workspace");
        },
        toast,
        onOpenDevSettings: () => setScreen("devSettings")
      }
    ), confirmDialog && /* @__PURE__ */ React.createElement(
      ConfirmDialog,
      {
        message: confirmDialog.message,
        onConfirm: confirmDialog.onConfirm,
        onCancel: () => setConfirmDialog(null)
      }
    ));
  }
  const service = services.find((s) => s.id === selectedServiceId);
  return /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement(
    WorkspaceScreen,
    {
      service,
      docTypeDefs,
      activeDocType,
      setActiveDocType,
      showAddDocType,
      setShowAddDocType,
      newDocTypeName,
      setNewDocTypeName,
      onAddDocType: addDocType,
      onDeleteDocType: deleteDocType,
      toast,
      getVersions: (docTypeId) => getVersions(selectedServiceId, docTypeId),
      getActiveVersionId: (docTypeId) => getActiveVersionId(selectedServiceId, docTypeId),
      selectVersion: (docTypeId, versionId) => selectVersion(selectedServiceId, docTypeId, versionId),
      onFile: (docTypeId, file) => handleFile(selectedServiceId, docTypeId, file),
      onCreateBlank: (docTypeId) => createBlankVersion(selectedServiceId, docTypeId),
      onUpdate: (docTypeId, versionId, sectionId, text) => updateSectionInVersion(selectedServiceId, docTypeId, versionId, sectionId, text),
      onToggleHeading: (docTypeId, versionId, sectionId) => toggleHeadingInVersion(selectedServiceId, docTypeId, versionId, sectionId),
      onDelete: (docTypeId, versionId, sectionId) => deleteSectionInVersion(selectedServiceId, docTypeId, versionId, sectionId),
      onAdd: (docTypeId, versionId) => addSectionInVersion(selectedServiceId, docTypeId, versionId),
      onMove: (docTypeId, versionId, sectionId, dir) => moveSectionInVersion(selectedServiceId, docTypeId, versionId, sectionId, dir),
      onDeleteVersion: (docTypeId, versionId) => deleteVersion(selectedServiceId, docTypeId, versionId),
      onPrint: (docTypeId, versionId) => openPrint(docTypeId, versionId),
      dict,
      dictTermInput,
      dictReadingInput,
      dictSkipInput,
      setDictTermInput,
      setDictReadingInput,
      setDictSkipInput,
      onDictAdd: addDictEntry,
      onDictRemove: removeDictEntry,
      onPlay: goToPlayback,
      onBackToMenu: () => setScreen("menu")
    }
  ), confirmDialog && /* @__PURE__ */ React.createElement(
    ConfirmDialog,
    {
      message: confirmDialog.message,
      onConfirm: confirmDialog.onConfirm,
      onCancel: () => setConfirmDialog(null)
    }
  ));
}
function ConfirmDialog({ message, onConfirm, onCancel }) {
  return /* @__PURE__ */ React.createElement("div", { className: "fixed inset-0 bg-black/40 flex items-center justify-center z-50 px-4" }, /* @__PURE__ */ React.createElement("div", { className: "bg-white rounded-2xl p-5 max-w-sm w-full shadow-xl" }, /* @__PURE__ */ React.createElement("p", { className: "text-sm text-stone-700 mb-5 whitespace-pre-wrap leading-relaxed" }, message), /* @__PURE__ */ React.createElement("div", { className: "flex gap-2 justify-end" }, /* @__PURE__ */ React.createElement("button", { onClick: onCancel, className: "px-4 py-2 rounded-lg text-sm font-semibold text-stone-500 border border-stone-300" }, "\u30AD\u30E3\u30F3\u30BB\u30EB"), /* @__PURE__ */ React.createElement("button", { onClick: onConfirm, className: "px-4 py-2 rounded-lg text-sm font-semibold bg-red-600 text-white" }, "\u524A\u9664\u3059\u308B"))));
}
function MenuScreen({
  services,
  newServiceName,
  setNewServiceName,
  onAdd,
  editingServiceId,
  editingServiceName,
  setEditingServiceName,
  onStartRename,
  onSaveRename,
  onCancelRename,
  onDelete,
  onSelect,
  toast,
  onOpenDevSettings
}) {
  return /* @__PURE__ */ React.createElement("div", { className: "min-h-screen bg-stone-50 text-stone-800" }, /* @__PURE__ */ React.createElement("div", { className: "max-w-2xl mx-auto px-4 py-8" }, /* @__PURE__ */ React.createElement("div", { className: "mb-6 text-center" }, /* @__PURE__ */ React.createElement("h1", { className: "text-xl font-bold text-teal-800" }, "\u91CD\u8981\u4E8B\u9805\u8AAC\u660E \u8AAD\u307F\u4E0A\u3052\u30A2\u30D7\u30EA"), /* @__PURE__ */ React.createElement("p", { className: "text-xs text-stone-500 mt-0.5" }, "\u30E2\u30C7\u30EB\u30B1\u30FC\u30B9\uFF0F\u307E\u305A\u30B5\u30FC\u30D3\u30B9\u3092\u9078\u3093\u3067\u304F\u3060\u3055\u3044")), toast && /* @__PURE__ */ React.createElement("div", { className: `mb-4 flex items-center gap-2 rounded-xl px-4 py-3 text-sm ${toast.type === "error" ? "bg-red-50 text-red-700 border border-red-200" : toast.type === "success" ? "bg-teal-50 text-teal-700 border border-teal-200" : "bg-stone-100 text-stone-600 border border-stone-200"}` }, toast.type === "error" ? /* @__PURE__ */ React.createElement(AlertCircle, { size: 16 }) : /* @__PURE__ */ React.createElement(CheckCircle2, { size: 16 }), toast.msg), services.length === 0 ? /* @__PURE__ */ React.createElement("div", { className: "text-center text-stone-400 text-sm py-10 border-2 border-dashed border-stone-200 rounded-2xl mb-6" }, "\u307E\u3060\u30B5\u30FC\u30D3\u30B9\u304C\u767B\u9332\u3055\u308C\u3066\u3044\u307E\u305B\u3093\u3002\u4E0B\u304B\u3089\u8FFD\u52A0\u3057\u3066\u304F\u3060\u3055\u3044\u3002") : /* @__PURE__ */ React.createElement("div", { className: "space-y-2 mb-6" }, services.map((s) => /* @__PURE__ */ React.createElement("div", { key: s.id, className: "flex items-center gap-1 bg-white border border-stone-200 rounded-2xl px-3 py-2" }, editingServiceId === s.id ? /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement(
    "input",
    {
      autoFocus: true,
      value: editingServiceName,
      onChange: (e) => setEditingServiceName(e.target.value),
      onKeyDown: (e) => {
        if (e.key === "Enter") onSaveRename();
        if (e.key === "Escape") onCancelRename();
      },
      className: "flex-1 border border-teal-300 rounded-lg px-2 py-1.5 text-sm"
    }
  ), /* @__PURE__ */ React.createElement("button", { onClick: onSaveRename, className: "text-teal-700 text-xs font-semibold px-2 shrink-0" }, "\u4FDD\u5B58"), /* @__PURE__ */ React.createElement("button", { onClick: onCancelRename, className: "text-stone-400 text-xs px-2 shrink-0" }, "\u30AD\u30E3\u30F3\u30BB\u30EB")) : /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("button", { onClick: () => onSelect(s.id), className: "flex-1 flex items-center gap-3 text-left px-1 py-1 min-w-0" }, /* @__PURE__ */ React.createElement("div", { className: "w-10 h-10 rounded-xl bg-teal-50 flex items-center justify-center text-teal-700 shrink-0" }, /* @__PURE__ */ React.createElement(Building2, { size: 18 })), /* @__PURE__ */ React.createElement("span", { className: "font-semibold text-stone-700 truncate" }, s.name)), /* @__PURE__ */ React.createElement("button", { onClick: () => onSelect(s.id), className: "text-teal-700 shrink-0 px-1" }, /* @__PURE__ */ React.createElement(ChevronRight, { size: 20 })), /* @__PURE__ */ React.createElement("button", { onClick: () => onStartRename(s.id), className: "text-stone-400 hover:text-teal-700 text-xs px-2 shrink-0" }, "\u7DE8\u96C6"), /* @__PURE__ */ React.createElement("button", { onClick: () => onDelete(s.id), className: "text-stone-400 hover:text-red-500 px-1 shrink-0" }, /* @__PURE__ */ React.createElement(Trash2, { size: 16 })))))), /* @__PURE__ */ React.createElement("div", { className: "bg-white rounded-2xl border border-stone-200 p-4 flex gap-2 flex-wrap" }, /* @__PURE__ */ React.createElement(
    "input",
    {
      value: newServiceName,
      onChange: (e) => setNewServiceName(e.target.value),
      onKeyDown: (e) => {
        if (e.key === "Enter") onAdd();
      },
      placeholder: "\u4F8B\uFF1A\u8A2A\u554F\u4ECB\u8B77\u3001\u901A\u6240\u4ECB\u8B77 \u306A\u3069",
      className: "flex-1 min-w-[200px] border border-stone-300 rounded-lg px-3 py-2 text-sm"
    }
  ), /* @__PURE__ */ React.createElement("button", { onClick: onAdd, className: "bg-teal-700 text-white px-4 py-2 rounded-lg text-sm font-semibold flex items-center gap-1" }, /* @__PURE__ */ React.createElement(Plus, { size: 16 }), " \u30B5\u30FC\u30D3\u30B9\u3092\u8FFD\u52A0")), /* @__PURE__ */ React.createElement("div", { className: "mt-8 text-center" }, /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: onOpenDevSettings,
      className: "text-stone-300 hover:text-stone-500 text-xs inline-flex items-center gap-1",
      title: "\u958B\u767A\u8005\u8A2D\u5B9A"
    },
    /* @__PURE__ */ React.createElement(Settings, { size: 12 }),
    " \u958B\u767A\u8005\u8A2D\u5B9A"
  ))));
}
function DevSettingsScreen({ devRules, dict, onAdd, onUpdate, onDelete, onBack }) {
  const [label, setLabel] = useState("");
  const [pattern, setPattern] = useState("");
  const [replacement, setReplacement] = useState("");
  const [caseInsensitive, setCaseInsensitive] = useState(false);
  const [testText, setTestText] = useState("\u96FB\u8A71\u756A\u53F7\uFF1A078-2841\u3000\u4E8B\u696D\u6240\u756A\u53F7\uFF1A\uFF12\uFF18\uFF17\uFF10\uFF19\uFF10\uFF14\uFF13\uFF17\uFF18");
  const handleAdd = () => {
    const ok = onAdd({ label, pattern, replacement, caseInsensitive });
    if (ok) {
      setLabel("");
      setPattern("");
      setReplacement("");
      setCaseInsensitive(false);
    }
  };
  const previewText = applyDict(applyDevRules(applyNumberReadingRules(testText), devRules), dict);
  return /* @__PURE__ */ React.createElement("div", { className: "min-h-screen bg-stone-50 text-stone-800" }, /* @__PURE__ */ React.createElement("div", { className: "max-w-2xl mx-auto px-4 py-6" }, /* @__PURE__ */ React.createElement("button", { onClick: onBack, className: "flex items-center gap-1 text-xs text-stone-500 mb-3" }, /* @__PURE__ */ React.createElement(ArrowLeft, { size: 14 }), " \u30E1\u30CB\u30E5\u30FC\u306B\u623B\u308B"), /* @__PURE__ */ React.createElement("div", { className: "mb-5" }, /* @__PURE__ */ React.createElement("h1", { className: "text-lg font-bold text-stone-700 flex items-center gap-2" }, /* @__PURE__ */ React.createElement(Settings, { size: 18 }), " \u958B\u767A\u8005\u8A2D\u5B9A\uFF1A\u56FA\u5B9A\u30EB\u30FC\u30EB"), /* @__PURE__ */ React.createElement("p", { className: "text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 mt-2" }, "\u3053\u3053\u306F\u30B9\u30BF\u30C3\u30D5\u5411\u3051\u306E\u300C\u3088\u307F\u304B\u305F\u8F9E\u66F8\u300D\u3068\u306F\u5225\u306E\u8A2D\u5B9A\u3067\u3059\u3002\u6B63\u898F\u8868\u73FE\u3067\u8AAD\u307F\u4E0A\u3052\u30C6\u30AD\u30B9\u30C8\u3092\u7F6E\u63DB\u3057\u307E\u3059\uFF08\u9069\u7528\u9806\u306F\u300C\u6570\u5B57\u306E\u81EA\u52D5\u5909\u63DB \u2192 \u958B\u767A\u8005\u8A2D\u5B9A\u306E\u56FA\u5B9A\u30EB\u30FC\u30EB \u2192 \u3088\u307F\u304B\u305F\u8F9E\u66F8\u300D\u3067\u3001\u3088\u307F\u304B\u305F\u8F9E\u66F8\u304C\u6700\u5F8C\uFF1D\u6700\u512A\u5148\u3067\u3059\uFF09\u3002\u66F8\u304D\u65B9\u3092\u8AA4\u308B\u3068\u8AAD\u307F\u4E0A\u3052\u5168\u4F53\u306B\u5F71\u97FF\u3059\u308B\u305F\u3081\u3001\u958B\u767A\u8005\u306E\u307F\u304C\u6271\u3063\u3066\u304F\u3060\u3055\u3044\u3002")), /* @__PURE__ */ React.createElement("div", { className: "bg-white rounded-2xl border border-stone-200 p-5 mb-4" }, /* @__PURE__ */ React.createElement("h2", { className: "font-bold text-stone-700 mb-3 text-sm" }, "\u30EB\u30FC\u30EB\u3092\u8FFD\u52A0"), /* @__PURE__ */ React.createElement("div", { className: "space-y-2" }, /* @__PURE__ */ React.createElement(
    "input",
    {
      value: label,
      onChange: (e) => setLabel(e.target.value),
      placeholder: "\u30E1\u30E2\uFF08\u4EFB\u610F\u3002\u4F8B\uFF1A\u300C\u4ECB\u8B77\u652F\u63F4\u5C02\u9580\u54E1\u300D\u306E\u8AAD\u307F\u4FEE\u6B63\uFF09",
      className: "w-full border border-stone-300 rounded-lg px-3 py-2 text-sm"
    }
  ), /* @__PURE__ */ React.createElement(
    "input",
    {
      value: pattern,
      onChange: (e) => setPattern(e.target.value),
      placeholder: "\u6B63\u898F\u8868\u73FE\u30D1\u30BF\u30FC\u30F3\uFF08\u4F8B\uFF1A\u4ECB\u8B77\u652F\u63F4\u5C02\u9580\u54E1\uFF09",
      className: "w-full border border-stone-300 rounded-lg px-3 py-2 text-sm font-mono"
    }
  ), /* @__PURE__ */ React.createElement(
    "input",
    {
      value: replacement,
      onChange: (e) => setReplacement(e.target.value),
      placeholder: "\u7F6E\u304D\u63DB\u3048\u5F8C\u306E\u8AAD\u307F\uFF08\u4F8B\uFF1A\u304B\u3044\u3054\u3057\u3048\u3093\u305B\u3093\u3082\u3093\u3044\u3093\u3002$1 \u3067\u6B63\u898F\u8868\u73FE\u306E\u30B0\u30EB\u30FC\u30D7\u3092\u53C2\u7167\u53EF\uFF09",
      className: "w-full border border-stone-300 rounded-lg px-3 py-2 text-sm font-mono"
    }
  ), /* @__PURE__ */ React.createElement("div", { className: "flex items-center justify-between" }, /* @__PURE__ */ React.createElement("label", { className: "flex items-center gap-2 text-xs text-stone-500 font-semibold" }, /* @__PURE__ */ React.createElement("input", { type: "checkbox", checked: caseInsensitive, onChange: (e) => setCaseInsensitive(e.target.checked) }), "\u5927\u6587\u5B57\u30FB\u5C0F\u6587\u5B57\u3092\u533A\u5225\u3057\u306A\u3044"), /* @__PURE__ */ React.createElement("button", { onClick: handleAdd, className: "bg-teal-700 text-white px-4 py-2 rounded-lg text-sm font-semibold flex items-center gap-1" }, /* @__PURE__ */ React.createElement(Plus, { size: 14 }), " \u30EB\u30FC\u30EB\u3092\u8FFD\u52A0")))), /* @__PURE__ */ React.createElement("div", { className: "bg-white rounded-2xl border border-stone-200 p-5 mb-4" }, /* @__PURE__ */ React.createElement("h2", { className: "font-bold text-stone-700 mb-1 text-sm" }, "\u52D5\u4F5C\u78BA\u8A8D"), /* @__PURE__ */ React.createElement("p", { className: "text-xs text-stone-500 mb-2" }, "\u4E0B\u306B\u30C6\u30B9\u30C8\u7528\u306E\u6587\u7AE0\u3092\u5165\u529B\u3059\u308B\u3068\u3001\u6570\u5B57\u30EB\u30FC\u30EB\u2192\u56FA\u5B9A\u30EB\u30FC\u30EB\u2192\u3088\u307F\u304B\u305F\u8F9E\u66F8\u306E\u9806\u3067\u9069\u7528\u3057\u305F\u7D50\u679C\u3092\u78BA\u8A8D\u3067\u304D\u307E\u3059\u3002"), /* @__PURE__ */ React.createElement(
    "textarea",
    {
      value: testText,
      onChange: (e) => setTestText(e.target.value),
      rows: 2,
      className: "w-full text-sm rounded-lg border border-stone-200 p-2 mb-2"
    }
  ), /* @__PURE__ */ React.createElement("div", { className: "text-xs text-stone-400 mb-1" }, "\u5909\u63DB\u7D50\u679C\uFF08\u5B9F\u969B\u306B\u8AAD\u307F\u4E0A\u3052\u3089\u308C\u308B\u30C6\u30AD\u30B9\u30C8\uFF09"), /* @__PURE__ */ React.createElement("div", { className: "text-sm bg-stone-50 rounded-lg px-3 py-2 whitespace-pre-wrap break-words" }, previewText)), /* @__PURE__ */ React.createElement("div", { className: "bg-white rounded-2xl border border-stone-200 p-5" }, /* @__PURE__ */ React.createElement("h2", { className: "font-bold text-stone-700 mb-3 text-sm" }, "\u767B\u9332\u6E08\u307F\u306E\u30EB\u30FC\u30EB\uFF08", devRules.length, "\u4EF6\uFF09"), devRules.length === 0 ? /* @__PURE__ */ React.createElement("div", { className: "text-center text-stone-400 text-sm py-8 border-2 border-dashed border-stone-200 rounded-xl" }, "\u307E\u3060\u767B\u9332\u304C\u3042\u308A\u307E\u305B\u3093") : /* @__PURE__ */ React.createElement("div", { className: "space-y-2" }, devRules.map((r) => /* @__PURE__ */ React.createElement("div", { key: r.id, className: "bg-stone-50 rounded-lg px-3 py-2" }, /* @__PURE__ */ React.createElement("div", { className: "flex items-start gap-2" }, /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: () => onUpdate(r.id, { enabled: !r.enabled }),
      title: r.enabled ? "\u7121\u52B9\u306B\u3059\u308B" : "\u6709\u52B9\u306B\u3059\u308B",
      className: r.enabled ? "text-teal-700 shrink-0" : "text-stone-300 shrink-0"
    },
    r.enabled ? /* @__PURE__ */ React.createElement(ToggleRight, { size: 22 }) : /* @__PURE__ */ React.createElement(ToggleLeft, { size: 22 })
  ), /* @__PURE__ */ React.createElement("div", { className: "flex-1 min-w-0 text-xs" }, r.label && /* @__PURE__ */ React.createElement("div", { className: "font-semibold text-stone-700 mb-0.5" }, r.label), /* @__PURE__ */ React.createElement("div", { className: "font-mono text-stone-600 break-all" }, "/", r.pattern, "/", r.caseInsensitive ? "gi" : "g", " \u2192 \u300C", r.replacement || "\uFF08\u524A\u9664\uFF09", "\u300D")), /* @__PURE__ */ React.createElement("button", { onClick: () => onDelete(r.id), className: "text-stone-400 hover:text-red-500 shrink-0" }, /* @__PURE__ */ React.createElement(Trash2, { size: 15 })))))))));
}
function WorkspaceScreen({
  service,
  docTypeDefs,
  activeDocType,
  setActiveDocType,
  showAddDocType,
  setShowAddDocType,
  newDocTypeName,
  setNewDocTypeName,
  onAddDocType,
  onDeleteDocType,
  toast,
  getVersions,
  getActiveVersionId,
  selectVersion,
  onFile,
  onCreateBlank,
  onUpdate,
  onToggleHeading,
  onDelete,
  onAdd,
  onMove,
  onDeleteVersion,
  onPrint,
  dict,
  dictTermInput,
  dictReadingInput,
  dictSkipInput,
  setDictTermInput,
  setDictReadingInput,
  setDictSkipInput,
  onDictAdd,
  onDictRemove,
  onPlay,
  onBackToMenu
}) {
  return /* @__PURE__ */ React.createElement("div", { className: "min-h-screen bg-stone-50 text-stone-800" }, /* @__PURE__ */ React.createElement("div", { className: "max-w-3xl mx-auto px-4 py-6" }, /* @__PURE__ */ React.createElement("div", { className: "mb-4" }, /* @__PURE__ */ React.createElement("button", { onClick: onBackToMenu, className: "flex items-center gap-1 text-xs text-stone-500 mb-1" }, /* @__PURE__ */ React.createElement(ArrowLeft, { size: 14 }), " \u30B5\u30FC\u30D3\u30B9\u9078\u629E\u306B\u623B\u308B"), /* @__PURE__ */ React.createElement("h1", { className: "text-lg font-bold text-teal-800" }, service ? service.name : "")), toast && /* @__PURE__ */ React.createElement("div", { className: `mb-4 flex items-center gap-2 rounded-xl px-4 py-3 text-sm ${toast.type === "error" ? "bg-red-50 text-red-700 border border-red-200" : toast.type === "success" ? "bg-teal-50 text-teal-700 border border-teal-200" : "bg-stone-100 text-stone-600 border border-stone-200"}` }, toast.type === "error" ? /* @__PURE__ */ React.createElement(AlertCircle, { size: 16 }) : /* @__PURE__ */ React.createElement(CheckCircle2, { size: 16 }), toast.msg), /* @__PURE__ */ React.createElement("div", { className: "flex gap-2 mb-5 flex-wrap items-center" }, docTypeDefs.map((d) => /* @__PURE__ */ React.createElement("div", { key: d.id, className: "relative" }, /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: () => setActiveDocType(d.id),
      className: `px-4 py-2 rounded-full text-sm font-semibold flex items-center gap-1.5 ${activeDocType === d.id ? "bg-teal-700 text-white" : "bg-white text-stone-600 border border-stone-300"}`
    },
    /* @__PURE__ */ React.createElement(FileText, { size: 14 }),
    " ",
    d.label
  ), d.deletable && activeDocType === d.id && /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: () => onDeleteDocType(d.id),
      title: "\u3053\u306E\u66F8\u985E\u306E\u7A2E\u985E\u3092\u524A\u9664",
      className: "absolute -top-1.5 -right-1.5 w-5 h-5 bg-stone-400 text-white rounded-full flex items-center justify-center"
    },
    /* @__PURE__ */ React.createElement(X, { size: 12 })
  ))), showAddDocType ? /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-1 bg-white border border-stone-300 rounded-full pl-3 pr-1 py-1" }, /* @__PURE__ */ React.createElement(
    "input",
    {
      autoFocus: true,
      value: newDocTypeName,
      onChange: (e) => setNewDocTypeName(e.target.value),
      onKeyDown: (e) => {
        if (e.key === "Enter") onAddDocType();
        if (e.key === "Escape") setShowAddDocType(false);
      },
      placeholder: "\u66F8\u985E\u540D",
      className: "text-sm outline-none w-28"
    }
  ), /* @__PURE__ */ React.createElement("button", { onClick: onAddDocType, className: "bg-teal-700 text-white rounded-full w-6 h-6 flex items-center justify-center shrink-0" }, /* @__PURE__ */ React.createElement(Plus, { size: 13 }))) : /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: () => setShowAddDocType(true),
      className: "px-3 py-2 rounded-full text-sm font-semibold text-teal-700 border border-dashed border-teal-300 flex items-center gap-1"
    },
    /* @__PURE__ */ React.createElement(Plus, { size: 14 }),
    " \u66F8\u985E\u3092\u8FFD\u52A0"
  ), /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: () => setActiveDocType("dict"),
      className: `px-4 py-2 rounded-lg text-sm font-semibold flex items-center gap-1.5 border-2 ${activeDocType === "dict" ? "bg-violet-600 text-white border-violet-600" : "bg-violet-50 text-violet-700 border-violet-200"}`
    },
    /* @__PURE__ */ React.createElement(BookOpen, { size: 14 }),
    " \u3088\u307F\u304B\u305F\u8F9E\u66F8"
  ), /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: onPlay,
      className: "ml-auto px-4 py-2 rounded-full text-sm font-semibold bg-amber-500 text-white flex items-center gap-1.5"
    },
    /* @__PURE__ */ React.createElement(Volume2, { size: 14 }),
    " \u8AAD\u307F\u4E0A\u3052\u958B\u59CB"
  )), activeDocType !== "dict" && docTypeDefs.some((d) => d.id === activeDocType) && /* @__PURE__ */ React.createElement(
    DocPanel,
    {
      key: activeDocType,
      label: docTypeDefs.find((d) => d.id === activeDocType).label,
      versions: getVersions(activeDocType),
      activeVersionId: getActiveVersionId(activeDocType),
      onSelectVersion: (vid) => selectVersion(activeDocType, vid),
      onFile: (f) => onFile(activeDocType, f),
      onCreateBlank: () => onCreateBlank(activeDocType),
      onUpdate: (vid, id, text) => onUpdate(activeDocType, vid, id, text),
      onToggleHeading: (vid, id) => onToggleHeading(activeDocType, vid, id),
      onDelete: (vid, id) => onDelete(activeDocType, vid, id),
      onAdd: (vid) => onAdd(activeDocType, vid),
      onMove: (vid, id, dir) => onMove(activeDocType, vid, id, dir),
      onDeleteVersion: (vid) => onDeleteVersion(activeDocType, vid),
      onPrint: (vid) => onPrint(activeDocType, vid)
    }
  ), activeDocType === "dict" && /* @__PURE__ */ React.createElement(
    DictPanel,
    {
      dict,
      termInput: dictTermInput,
      readingInput: dictReadingInput,
      skipInput: dictSkipInput,
      setTermInput: setDictTermInput,
      setReadingInput: setDictReadingInput,
      setSkipInput: setDictSkipInput,
      onAdd: onDictAdd,
      onRemove: onDictRemove
    }
  )));
}
function DocPanel({
  label,
  versions,
  activeVersionId,
  onSelectVersion,
  onFile,
  onCreateBlank,
  onUpdate,
  onToggleHeading,
  onDelete,
  onAdd,
  onMove,
  onDeleteVersion,
  onPrint
}) {
  const inputRef = useRef(null);
  const activeVersion = versions.find((v) => v.id === activeVersionId) || null;
  return /* @__PURE__ */ React.createElement("div", { className: "bg-white rounded-2xl border border-stone-200 p-5" }, /* @__PURE__ */ React.createElement("div", { className: "flex items-center justify-between mb-3 flex-wrap gap-2" }, /* @__PURE__ */ React.createElement("h2", { className: "font-bold text-stone-700" }, label), /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-2" }, /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: () => inputRef.current && inputRef.current.click(),
      className: "flex items-center gap-1.5 text-sm bg-teal-700 text-white px-3 py-1.5 rounded-lg"
    },
    /* @__PURE__ */ React.createElement(Upload, { size: 14 }),
    " \u30D5\u30A1\u30A4\u30EB\u3092\u9078\u629E"
  ), /* @__PURE__ */ React.createElement(
    "input",
    {
      ref: inputRef,
      type: "file",
      accept: ".docx,.xlsx",
      className: "hidden",
      onChange: (e) => {
        if (e.target.files[0]) onFile(e.target.files[0]);
        e.target.value = "";
      }
    }
  ))), /* @__PURE__ */ React.createElement("p", { className: "text-xs text-stone-500 mb-4" }, "Word\uFF08.docx\uFF09\u307E\u305F\u306FExcel\uFF08.xlsx\uFF09\u3092\u8AAD\u307F\u8FBC\u307F\u307E\u3059\u3002PDF\u306B\u306F\u5BFE\u5FDC\u3057\u3066\u3044\u307E\u305B\u3093\u3002\u8AAD\u307F\u8FBC\u3080\u305F\u3073\u306B\u65B0\u3057\u3044\u30D0\u30FC\u30B8\u30E7\u30F3\u3068\u3057\u3066\u4FDD\u5B58\u3055\u308C\u3001\u904E\u53BB\u306E\u30D5\u30A1\u30A4\u30EB\u3082\u547C\u3073\u51FA\u3057\u3066\u78BA\u8A8D\u30FB\u5370\u5237\u3067\u304D\u307E\u3059\uFF08\u4FDD\u5B58\u3067\u304D\u308B\u306E\u306F\u6700\u65B0\uFF0B\u904E\u53BB4\u4EF6\u306E\u5408\u8A085\u4EF6\u307E\u3067\u3002\u305D\u308C\u3092\u8D85\u3048\u308B\u3068\u4E00\u756A\u53E4\u3044\u3082\u306E\u304B\u3089\u81EA\u52D5\u7684\u306B\u524A\u9664\u3055\u308C\u307E\u3059\uFF09\u3002\u8AAD\u307F\u8FBC\u3093\u3060\u5185\u5BB9\u306F\u4E0B\u66F8\u304D\u3067\u3059\u3002\u5B9F\u969B\u306B\u8AAD\u307F\u4E0A\u3052\u308B\u6587\u7AE0\u306B\u306A\u308B\u3088\u3046\u78BA\u8A8D\u30FB\u7DE8\u96C6\u3057\u3066\u304F\u3060\u3055\u3044\u3002"), versions.length > 0 && /* @__PURE__ */ React.createElement("div", { className: "mb-4" }, /* @__PURE__ */ React.createElement("div", { className: "text-xs font-semibold text-stone-500 mb-1.5 flex items-center gap-1" }, /* @__PURE__ */ React.createElement(History, { size: 13 }), " \u4FDD\u5B58\u3055\u308C\u3066\u3044\u308B\u30D0\u30FC\u30B8\u30E7\u30F3\uFF08", versions.length, " / \u6700\u5927", MAX_VERSIONS_PER_DOC, "\u4EF6\uFF09"), /* @__PURE__ */ React.createElement("div", { className: "flex flex-col gap-1.5" }, versions.map((v, i) => /* @__PURE__ */ React.createElement(
    "div",
    {
      key: v.id,
      className: `flex items-center gap-2 rounded-lg px-3 py-2 text-xs ${v.id === activeVersionId ? "bg-teal-50 border border-teal-300" : "bg-stone-50 border border-transparent"}`
    },
    /* @__PURE__ */ React.createElement("button", { onClick: () => onSelectVersion(v.id), className: "flex-1 min-w-0 text-left flex items-center gap-2" }, i === 0 && /* @__PURE__ */ React.createElement("span", { className: "shrink-0 bg-teal-700 text-white text-[10px] font-bold px-1.5 py-0.5 rounded" }, "\u6700\u65B0"), /* @__PURE__ */ React.createElement("span", { className: "truncate font-medium text-stone-700" }, v.fileName || "\uFF08\u30D5\u30A1\u30A4\u30EB\u540D\u306A\u3057\uFF09"), /* @__PURE__ */ React.createElement("span", { className: "shrink-0 text-stone-400" }, formatDate(v.uploadedAt))),
    /* @__PURE__ */ React.createElement("button", { onClick: () => onPrint(v.id), title: "\u5370\u5237", className: "text-stone-400 hover:text-teal-700 shrink-0" }, /* @__PURE__ */ React.createElement(Printer, { size: 14 })),
    /* @__PURE__ */ React.createElement("button", { onClick: () => onDeleteVersion(v.id), title: "\u524A\u9664", className: "text-stone-400 hover:text-red-500 shrink-0" }, /* @__PURE__ */ React.createElement(Trash2, { size: 14 }))
  )))), !activeVersion ? /* @__PURE__ */ React.createElement("div", { className: "text-center text-stone-400 text-sm py-8 border-2 border-dashed border-stone-200 rounded-xl mb-3" }, "\u307E\u3060\u30D5\u30A1\u30A4\u30EB\u304C\u8AAD\u307F\u8FBC\u307E\u308C\u3066\u3044\u307E\u305B\u3093", /* @__PURE__ */ React.createElement("div", { className: "mt-3" }, /* @__PURE__ */ React.createElement("button", { onClick: onCreateBlank, className: "text-teal-700 text-xs font-semibold border border-teal-300 rounded-lg px-3 py-1.5" }, "\u30D5\u30A1\u30A4\u30EB\u3092\u4F7F\u308F\u305A\u624B\u5165\u529B\u3067\u4F5C\u6210\u3059\u308B"))) : /* @__PURE__ */ React.createElement("div", { className: "space-y-2" }, activeVersion.sections.map((s, i) => /* @__PURE__ */ React.createElement("div", { key: s.id, className: `flex gap-2 items-start rounded-xl p-2 ${s.isHeading ? "bg-amber-50" : "bg-stone-50"}` }, /* @__PURE__ */ React.createElement("div", { className: "flex flex-col gap-1 pt-1" }, /* @__PURE__ */ React.createElement("button", { onClick: () => onMove(activeVersion.id, s.id, -1), disabled: i === 0, className: "text-stone-400 disabled:opacity-30" }, /* @__PURE__ */ React.createElement(ChevronUp, { size: 16 })), /* @__PURE__ */ React.createElement("button", { onClick: () => onMove(activeVersion.id, s.id, 1), disabled: i === activeVersion.sections.length - 1, className: "text-stone-400 disabled:opacity-30" }, /* @__PURE__ */ React.createElement(ChevronDown, { size: 16 }))), /* @__PURE__ */ React.createElement("div", { className: "flex-1" }, /* @__PURE__ */ React.createElement("label", { className: "flex items-center gap-1.5 text-xs text-stone-500 mb-1" }, /* @__PURE__ */ React.createElement("input", { type: "checkbox", checked: s.isHeading, onChange: () => onToggleHeading(activeVersion.id, s.id) }), "\u898B\u51FA\u3057\u3068\u3057\u3066\u6271\u3046"), /* @__PURE__ */ React.createElement(
    "textarea",
    {
      value: s.text,
      onChange: (e) => onUpdate(activeVersion.id, s.id, e.target.value),
      rows: s.isHeading ? 1 : 2,
      className: `w-full text-sm rounded-lg border border-stone-200 p-2 ${s.isHeading ? "font-bold text-teal-800" : ""}`
    }
  )), /* @__PURE__ */ React.createElement("button", { onClick: () => onDelete(activeVersion.id, s.id), className: "text-stone-400 hover:text-red-500 pt-1" }, /* @__PURE__ */ React.createElement(Trash2, { size: 16 }))))), /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: () => activeVersion && onAdd(activeVersion.id),
      disabled: !activeVersion,
      className: "mt-3 flex items-center gap-1.5 text-sm text-teal-700 font-semibold disabled:opacity-30"
    },
    /* @__PURE__ */ React.createElement(Plus, { size: 16 }),
    " \u9805\u76EE\u3092\u8FFD\u52A0"
  ));
}
function DictPanel({ dict, termInput, readingInput, skipInput, setTermInput, setReadingInput, setSkipInput, onAdd, onRemove }) {
  const entries = Object.entries(dict);
  return /* @__PURE__ */ React.createElement("div", { className: "bg-white rounded-2xl border border-stone-200 p-5" }, /* @__PURE__ */ React.createElement("h2", { className: "font-bold text-stone-700 mb-1" }, "\u3088\u307F\u304B\u305F\u8F9E\u66F8"), /* @__PURE__ */ React.createElement("p", { className: "text-xs text-stone-500 mb-4" }, "\u8AA4\u3063\u3066\u8AAD\u307E\u308C\u305F\u5358\u8A9E\uFF08\u4EBA\u540D\u30FB\u4E8B\u696D\u6240\u540D\u30FB\u5C02\u9580\u7528\u8A9E\u306A\u3069\uFF09\u306E\u6B63\u3057\u3044\u8AAD\u307F\u65B9\u3092\u3072\u3089\u304C\u306A\u3067\u767B\u9332\u3057\u307E\u3059\u3002\u30AB\u30C3\u30B3\u66F8\u304D\u306E\u6CE8\u91C8\u306A\u3069\u300C\u8AAD\u307F\u4E0A\u3052\u81EA\u4F53\u304C\u4E0D\u8981\u306A\u90E8\u5206\u300D\u306F\u3001\u4E0B\u306E\u300C\u8AAD\u307F\u4E0A\u3052\u304B\u3089\u9664\u5916\u3059\u308B\u300D\u3092\u4F7F\u3046\u3068\u3001\u305D\u306E\u90E8\u5206\u3060\u3051\u8AAD\u307E\u308C\u306A\u304F\u306A\u308A\u307E\u3059\u3002\u4E00\u5EA6\u767B\u9332\u3059\u308C\u3070\u3001\u4EE5\u5F8C\u3059\u3079\u3066\u306E\u30B5\u30FC\u30D3\u30B9\u30FB\u3059\u3079\u3066\u306E\u8AAD\u307F\u4E0A\u3052\u306B\u81EA\u52D5\u3067\u53CD\u6620\u3055\u308C\u3001\u4ED6\u306E\u3069\u306E\u81EA\u52D5\u5909\u63DB\u3088\u308A\u3082\u512A\u5148\u3055\u308C\u307E\u3059\u3002"), /* @__PURE__ */ React.createElement("p", { className: "text-xs text-teal-700 bg-teal-50 rounded-lg px-3 py-2 mb-4" }, "\u{1F4A1}\u300C\u6240\u5728\u5730\u300D\u300C\u66F8\u300D\u306A\u3069\u3088\u304F\u3042\u308B\u8AA4\u8AAD\u306F\u6700\u521D\u304B\u3089\u767B\u9332\u6E08\u307F\u3067\u3059\uFF08\u4E0D\u8981\u306A\u3089\u524A\u9664\u3067\u304D\u307E\u3059\uFF09\u3002\u307E\u305F\u3001\u96FB\u8A71\u756A\u53F7\u30FB\u4F4F\u6240\u306E\u756A\u5730\u30FB\u4E8B\u696D\u6240\u756A\u53F7\u306A\u3069\u300C\u756A\u53F7\u300D\u300C\u30B3\u30FC\u30C9\u300D\u300C\u90F5\u4FBF\u300D\u7B49\u306E\u30E9\u30D9\u30EB\u304C\u4ED8\u3044\u305F\u6570\u5B57\u306F\u3001\u8F9E\u66F8\u306B\u767B\u9332\u3057\u306A\u304F\u3066\u3082\u81EA\u52D5\u7684\u306B\u300C\u6570\u5B57\u306F1\u6587\u5B57\u305A\u3064\u30FB\u30CF\u30A4\u30D5\u30F3\u306F\u300E\u306E\u300F\u300D\u3067\u8AAD\u307F\u4E0A\u3052\u307E\u3059\uFF08\u4F8B\uFF1A20-2841 \u2192 \u306B\u30BC\u30ED\u306E\u306B\u306F\u3061\u3088\u3093\u3044\u3061\uFF0F\u4E8B\u696D\u6240\u756A\u53F7\uFF12\uFF18\uFF17\uFF10\uFF19\uFF10\uFF14\uFF13\uFF17\uFF18 \u2192 \u306B\u306F\u3061\u306A\u306A\u30BC\u30ED\u2026\u30681\u6587\u5B57\u305A\u3064\uFF09\u3002"), /* @__PURE__ */ React.createElement("div", { className: "flex gap-2 mb-2 flex-wrap" }, /* @__PURE__ */ React.createElement(
    "input",
    {
      value: termInput,
      onChange: (e) => setTermInput(e.target.value),
      placeholder: "\u4F8B\uFF1A\u539F\u7530\u826F\u61B2 \u307E\u305F\u306F \uFF08\u4EE3\u8868\u8005\u306E\u5F79\u8077\u540D\u53CA\u3073\u6C0F\u540D\uFF09",
      className: "flex-1 min-w-[200px] border border-stone-300 rounded-lg px-3 py-2 text-sm"
    }
  ), /* @__PURE__ */ React.createElement(
    "input",
    {
      value: readingInput,
      onChange: (e) => setReadingInput(e.target.value),
      placeholder: "\u4F8B\uFF1A\u306F\u3089\u3060\u3088\u3057\u306E\u308A",
      disabled: skipInput,
      className: "flex-1 min-w-[160px] border border-stone-300 rounded-lg px-3 py-2 text-sm disabled:bg-stone-100 disabled:text-stone-400"
    }
  ), /* @__PURE__ */ React.createElement("button", { onClick: onAdd, className: "bg-teal-700 text-white px-4 py-2 rounded-lg text-sm font-semibold flex items-center gap-1" }, /* @__PURE__ */ React.createElement(Plus, { size: 14 }), " \u767B\u9332")), /* @__PURE__ */ React.createElement("label", { className: "flex items-center gap-2 text-xs text-stone-500 font-semibold mb-4" }, /* @__PURE__ */ React.createElement(
    "input",
    {
      type: "checkbox",
      checked: skipInput,
      onChange: (e) => setSkipInput(e.target.checked)
    }
  ), "\u3053\u306E\u5358\u8A9E\u30FB\u6587\u5B57\u5217\u306F\u8AAD\u307F\u4E0A\u3052\u304B\u3089\u9664\u5916\u3059\u308B\uFF08\u8AAD\u307E\u306A\u3044\uFF09"), entries.length === 0 ? /* @__PURE__ */ React.createElement("div", { className: "text-center text-stone-400 text-sm py-8 border-2 border-dashed border-stone-200 rounded-xl" }, "\u307E\u3060\u767B\u9332\u304C\u3042\u308A\u307E\u305B\u3093") : /* @__PURE__ */ React.createElement("div", { className: "space-y-1.5" }, entries.map(([term, reading]) => /* @__PURE__ */ React.createElement("div", { key: term, className: "flex items-center justify-between bg-stone-50 rounded-lg px-3 py-2 text-sm" }, /* @__PURE__ */ React.createElement("span", null, /* @__PURE__ */ React.createElement("span", { className: "font-semibold" }, term), /* @__PURE__ */ React.createElement("span", { className: "text-stone-400 mx-1" }, "\u2192"), reading === "" ? /* @__PURE__ */ React.createElement("span", { className: "text-red-500 font-semibold" }, "\uFF08\u8AAD\u307F\u4E0A\u3052\u3067\u6D88\u53BB\uFF09") : reading), /* @__PURE__ */ React.createElement("button", { onClick: () => onRemove(term), className: "text-stone-400 hover:text-red-500" }, /* @__PURE__ */ React.createElement(X, { size: 15 }))))));
}
function PrintView({ service, docType, version, onClose }) {
  if (!version) {
    return /* @__PURE__ */ React.createElement("div", { className: "fixed inset-0 bg-white flex items-center justify-center" }, /* @__PURE__ */ React.createElement("div", { className: "text-center" }, /* @__PURE__ */ React.createElement("p", { className: "text-stone-500 mb-4" }, "\u5370\u5237\u5BFE\u8C61\u304C\u898B\u3064\u304B\u308A\u307E\u305B\u3093\u3067\u3057\u305F\u3002"), /* @__PURE__ */ React.createElement("button", { onClick: onClose, className: "bg-teal-700 text-white px-4 py-2 rounded-lg text-sm" }, "\u623B\u308B")));
  }
  return /* @__PURE__ */ React.createElement("div", { className: "fixed inset-0 bg-white overflow-y-auto text-stone-900" }, /* @__PURE__ */ React.createElement("style", null, `
        @media print {
          .noprint { display: none !important; }
          @page { margin: 18mm 16mm; }
        }
      `), /* @__PURE__ */ React.createElement("div", { className: "noprint sticky top-0 bg-white border-b border-stone-200 px-6 py-3 flex items-center gap-3 z-10" }, /* @__PURE__ */ React.createElement("button", { onClick: onClose, className: "flex items-center gap-1 text-sm text-stone-500" }, /* @__PURE__ */ React.createElement(ArrowLeft, { size: 16 }), " \u623B\u308B"), /* @__PURE__ */ React.createElement("button", { onClick: () => window.print(), className: "ml-auto flex items-center gap-1.5 bg-teal-700 text-white px-4 py-2 rounded-lg text-sm font-semibold" }, /* @__PURE__ */ React.createElement(Printer, { size: 15 }), " \u5370\u5237\u3059\u308B")), /* @__PURE__ */ React.createElement("div", { className: "max-w-3xl mx-auto px-8 py-10" }, /* @__PURE__ */ React.createElement("div", { className: "mb-8 text-center border-b border-stone-300 pb-4" }, /* @__PURE__ */ React.createElement("div", { className: "text-sm text-stone-500 mb-1" }, service ? service.name : ""), /* @__PURE__ */ React.createElement("h1", { className: "text-xl font-bold" }, docType ? docType.label : ""), /* @__PURE__ */ React.createElement("div", { className: "text-xs text-stone-400 mt-2" }, "\u30D5\u30A1\u30A4\u30EB\u540D\uFF1A", version.fileName || "\uFF08\u624B\u5165\u529B\uFF09", "\u3000\uFF0F\u3000\u4FDD\u5B58\u65E5\u6642\uFF1A", formatDate(version.uploadedAt))), /* @__PURE__ */ React.createElement("div", { className: "space-y-3 leading-relaxed" }, version.sections.map((s) => /* @__PURE__ */ React.createElement("div", { key: s.id, className: s.isHeading ? "font-bold text-base mt-6" : "text-sm" }, s.text)))));
}
function PlaybackPanel({
  queue,
  current,
  isSpeaking,
  wordSpans,
  activeWordIdx,
  rate,
  setRate,
  autoAdvance,
  setAutoAdvance,
  onTogglePlay,
  onPrev,
  onNext,
  onRestart,
  onBack
}) {
  const item = queue[current];
  const stageRef = useRef(null);
  const textStageRef = useRef(null);
  const [isFullscreen, setIsFullscreen] = useState(false);
  useEffect(() => {
    const onChange = () => setIsFullscreen(!!document.fullscreenElement);
    document.addEventListener("fullscreenchange", onChange);
    return () => document.removeEventListener("fullscreenchange", onChange);
  }, []);
  useEffect(() => {
    if (textStageRef.current) textStageRef.current.scrollTop = 0;
  }, [current]);
  const toggleFullscreen = () => {
    if (!stageRef.current) return;
    if (document.fullscreenElement) {
      document.exitFullscreen().catch(() => {
      });
    } else if (stageRef.current.requestFullscreen) {
      stageRef.current.requestFullscreen().catch(() => {
      });
    }
  };
  return /* @__PURE__ */ React.createElement("div", { ref: stageRef, className: "fixed inset-0 bg-stone-50 flex flex-col text-stone-800" }, /* @__PURE__ */ React.createElement("style", null, `
        @media (orientation: landscape) {
          .yomiage-toolbar { padding-top: 0.5rem; padding-bottom: 0.5rem; }
          .yomiage-controls { padding-top: 0.5rem; padding-bottom: 0.75rem; }
        }
        .yomiage-text {
          font-size: clamp(50px, 8.8vw, 128px);
          line-height: 1.6;
        }
        .yomiage-word.active {
          background: #FDE68A;
          color: #134E4A;
          font-weight: 700;
          text-decoration: underline;
          text-decoration-color: #D97706;
          text-decoration-thickness: 4px;
          text-underline-offset: 6px;
          border-radius: 6px;
          padding: 0 2px;
        }
      `), /* @__PURE__ */ React.createElement("div", { className: "yomiage-toolbar flex items-center gap-3 px-4 pt-3 pb-2 shrink-0" }, /* @__PURE__ */ React.createElement("button", { onClick: onBack, className: "flex items-center gap-1 text-sm text-stone-500 shrink-0" }, /* @__PURE__ */ React.createElement(ArrowLeft, { size: 16 }), " \u623B\u308B"), /* @__PURE__ */ React.createElement("div", { className: "flex-1 min-w-0" }, /* @__PURE__ */ React.createElement("div", { className: "text-xs font-bold text-teal-700 truncate" }, item?.docLabel), /* @__PURE__ */ React.createElement("div", { className: "flex gap-1 mt-1" }, queue.map((_, i) => /* @__PURE__ */ React.createElement("div", { key: i, className: `flex-1 h-1.5 rounded-full ${i < current ? "bg-teal-700" : i === current ? "bg-amber-500" : "bg-stone-200"}` })))), /* @__PURE__ */ React.createElement("div", { className: "text-xs font-semibold text-stone-400 shrink-0" }, current + 1, " / ", queue.length), /* @__PURE__ */ React.createElement("button", { onClick: toggleFullscreen, className: "text-stone-400 shrink-0" }, isFullscreen ? /* @__PURE__ */ React.createElement(Minimize2, { size: 18 }) : /* @__PURE__ */ React.createElement(Maximize2, { size: 18 }))), /* @__PURE__ */ React.createElement("div", { ref: textStageRef, className: "flex-1 min-h-0 overflow-y-auto px-6 md:px-16 flex flex-col items-center" }, /* @__PURE__ */ React.createElement("div", { className: "yomiage-text w-full max-w-5xl text-center my-auto" }, wordSpans.map((w, i) => /* @__PURE__ */ React.createElement("span", { key: i, className: `yomiage-word ${i === activeWordIdx ? "active" : ""}` }, w)))), /* @__PURE__ */ React.createElement("div", { className: "yomiage-controls shrink-0 px-4 pb-4" }, /* @__PURE__ */ React.createElement("div", { className: "flex items-center justify-center gap-4 mb-3" }, /* @__PURE__ */ React.createElement("button", { onClick: onRestart, className: "w-11 h-11 rounded-2xl border border-stone-300 bg-white flex items-center justify-center text-stone-600" }, /* @__PURE__ */ React.createElement(RotateCcw, { size: 17 })), /* @__PURE__ */ React.createElement("button", { onClick: onPrev, disabled: current === 0, className: "w-14 h-14 rounded-2xl border border-stone-300 bg-white flex items-center justify-center text-stone-600 disabled:opacity-30" }, /* @__PURE__ */ React.createElement(SkipBack, { size: 20 })), /* @__PURE__ */ React.createElement("button", { onClick: onTogglePlay, className: "w-20 h-20 rounded-3xl bg-teal-700 text-white flex items-center justify-center shadow-lg" }, isSpeaking ? /* @__PURE__ */ React.createElement(Pause, { size: 30 }) : /* @__PURE__ */ React.createElement(Play, { size: 30 })), /* @__PURE__ */ React.createElement("button", { onClick: onNext, disabled: current === queue.length - 1, className: "w-14 h-14 rounded-2xl border border-stone-300 bg-white flex items-center justify-center text-stone-600 disabled:opacity-30" }, /* @__PURE__ */ React.createElement(SkipForward, { size: 20 }))), /* @__PURE__ */ React.createElement("div", { className: "flex items-center justify-center gap-4" }, /* @__PURE__ */ React.createElement("div", { className: "flex gap-1 bg-stone-100 rounded-xl p-1" }, [[0.75, "\u3086\u3063\u304F\u308A"], [1, "\u3075\u3064\u3046"], [1.25, "\u306F\u3084\u3044"]].map(([r, label]) => /* @__PURE__ */ React.createElement(
    "button",
    {
      key: r,
      onClick: () => setRate(r),
      className: `text-xs font-semibold px-3 py-1.5 rounded-lg ${rate === r ? "bg-teal-700 text-white" : "text-stone-500"}`
    },
    label
  ))), /* @__PURE__ */ React.createElement("label", { className: "flex items-center gap-2 text-xs text-stone-500 font-semibold" }, "\u81EA\u52D5\u3067\u6B21\u3078", /* @__PURE__ */ React.createElement("input", { type: "checkbox", checked: autoAdvance, onChange: (e) => setAutoAdvance(e.target.checked) })))));
}

// main-esm.jsx
function showFatalError(err) {
  const container = document.getElementById("root");
  if (!container) return;
  const msg = err && (err.stack || err.message) ? err.stack || err.message : String(err);
  container.innerHTML = '<div style="padding:20px;font-family:sans-serif;font-size:13px;color:#7f1d1d;background:#fef2f2;white-space:pre-wrap;word-break:break-word;">\u8D77\u52D5\u6642\u306B\u30A8\u30E9\u30FC\u304C\u767A\u751F\u3057\u307E\u3057\u305F\u3002\u3053\u306E\u5185\u5BB9\u3092\u30B9\u30AF\u30EA\u30FC\u30F3\u30B7\u30E7\u30C3\u30C8\u3067\u9001\u3063\u3066\u304F\u3060\u3055\u3044\u3002\n\n' + String(msg).replace(/</g, "&lt;") + "</div>";
}
window.addEventListener("error", (e) => {
  showFatalError(e.error && (e.error.stack || e.error.message) || e.message || e);
});
window.addEventListener("unhandledrejection", (e) => {
  showFatalError(e.reason && (e.reason.stack || e.reason.message) || e.reason);
});
try {
  const container = document.getElementById("root");
  const root = createRoot(container);
  root.render(React2.createElement(JuuyoujikouYomiageApp));
} catch (err) {
  showFatalError(err);
}
