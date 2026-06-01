// @ts-nocheck
import React, {
  useState,
  useEffect,
  useCallback,
  useRef,
  useMemo,
} from "react";
import {
  Plus,
  Trash2,
  Settings,
  ChevronRight,
  ChevronLeft,
  Layers,
  LayoutDashboard,
  X,
  ListPlus,
  CheckSquare,
  Square,
  AlertCircle,
  Printer,
  FileDown,
  Clipboard,
  ClipboardPaste,
  FolderPlus,
  Building2,
  Box,
  Wrench,
  Copy,
  FileSpreadsheet,
  Search,
  Calculator,
  Layout,
  Ruler,
} from "lucide-react";
import { exportRealExcelElegante } from "./excelExportHelper";
import { initializeApp } from "firebase/app";
import {
  getAuth,
  signInWithCustomToken,
  signInAnonymously,
  onAuthStateChanged,
} from "firebase/auth";
import {
  getFirestore,
  doc,
  setDoc,
  getDoc,
  deleteDoc,
} from "firebase/firestore";

const firebaseConfig =
  typeof __firebase_config !== "undefined" ? JSON.parse(__firebase_config) : {};
const app =
  Object.keys(firebaseConfig).length > 0 ? initializeApp(firebaseConfig) : null;
const auth = app ? getAuth(app) : null;
const db = app ? getFirestore(app) : null;
const appId = typeof __app_id !== "undefined" ? __app_id : "default-app-id";

const PESOS_VARILLA = {
  "3": 0.557,
  "4": 0.996,
  "5": 1.554,
  "6": 2.239,
  "8": 3.975,
  "10": 6.225,
};

const ANCLAJE_DEFAULT = {
  "3": "0.25",
  "4": "0.30",
  "5": "0.35",
  "6": "0.40",
  "8": "0.55",
  "10": "0.60",
};
const ACERO_MULT = {
  "3": 0.25,
  "4": 0.3,
  "5": "0.35",
  "6": "0.40",
  "8": "0.55",
  "10": "0.60",
};

const getSteelTypesForElement = (t) => {
  t = (t || "").toLowerCase();
  if (t === "muro" || t === "muro curvo")
    return ["Vertical", "Horizontal", "Refuerzo Adicional"];
  if (
    t === "losa" ||
    t === "losa nervada" ||
    t.includes("zapata") ||
    t === "escalera papelillo" ||
    t === "rampa de escalera"
  )
    return ["Longitudinal", "Transversal", "Bastones"];
  if (t === "columna circular" || t === "pilas" || t === "pila")
    return ["Longitudinal", "Zuncho", "Grapas"];
  if (t === "columna" || t === "dado")
    return ["Longitudinal", "Estribos", "Grapas"];
  if (["trabe", "nervadura", "contratrabe"].includes(t))
    return ["Longitudinal", "Bastones", "Estribos"];
  return ["Principal", "Secundario", "Estribos"];
};

const getEffectiveLargo = (r) => {
  let l = parseFloat(r.largo) || 0;
  const a = parseFloat(r.ancho) || 0;
  const tipo = (r.tipo || "").toLowerCase();
  if (r.i && ["trabe", "contratrabe", "zapata corrida"].includes(tipo))
    l = Math.max(0, l - a / 2);
  return l;
};

const getNextClaveValue = (lastClave) => {
  if (!lastClave) return "";
  const match = lastClave.match(/(.*?)(\d+)$/);
  if (!match) return lastClave;
  return (
    match[1] +
    (parseInt(match[2], 10) + 1).toString().padStart(match[2].length, "0")
  );
};

const exportExcelFile = (content, filename, includeBOM = true) => {
  const blob = new Blob([includeBOM ? "\ufeff" + content : content], {
    type: "application/vnd.ms-excel" + (includeBOM ? ";charset=utf-8" : ""),
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
};

const exportFormattedExcel = (tableHtml, filename) => {
  const template = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
    <head>
      <meta charset="utf-8">
      <style> 
        table { border-collapse: collapse; font-family: Arial, sans-serif; font-size: 10pt; margin-bottom: 20px; table-layout: fixed; } 
        th, td { border: 1px solid #cbd5e1; padding: 6px; vertical-align: middle; white-space: normal !important; word-wrap: break-word !important; } 
        th { font-weight: bold; text-align: center; }
        .wrap { white-space: normal !important; mso-height-source: auto; }
      </style>
    </head>
    <body>${tableHtml}</body>
  </html>`;
  const blob = new Blob(["\ufeff" + template], {
    type: "application/vnd.ms-excel;charset=utf-8",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename.endsWith(".xls") ? filename : filename + ".xls";
  link.click();
  URL.revokeObjectURL(url);
};

function LoadingScreen({ text }) {
  return (
    <div className="min-h-screen bg-[#fcfdfe] flex flex-col items-center justify-center p-8">
      <div className="w-16 h-16 border-4 border-blue-200 border-t-blue-600 rounded-full animate-spin mb-6"></div>
      <h2 className="text-xl font-black text-slate-800 uppercase tracking-widest animate-pulse">
        {text}
      </h2>
      <p className="text-xs text-slate-400 mt-4 uppercase tracking-widest">
        Sincronizando con la nube...
      </p>
    </div>
  );
}

const usePersistentState = (key, initialValue, user) => {
  const [state, setState] = useState(() => {
    const localData = localStorage.getItem(key);
    if (localData) {
      try {
        const parsed = JSON.parse(localData);
        return typeof initialValue === "object" && !Array.isArray(initialValue)
          ? { ...initialValue, ...parsed }
          : parsed;
      } catch (e) {
        return initialValue;
      }
    }
    return initialValue;
  });
  const [isLoaded, setIsLoaded] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const stateRef = useRef(state);
  useEffect(() => {
    stateRef.current = state;
  }, [state]);

  useEffect(() => {
    if (!user || !db) {
      setIsLoaded(true);
      return;
    }
    const docRef = doc(
      db,
      "artifacts",
      appId,
      "users",
      user.uid,
      "global",
      "settings",
    );
    getDoc(docRef)
      .then((snap) => {
        if (snap.exists() && snap.data()[key] !== undefined) {
          const fetchedVal = snap.data()[key];
          setState((prev) =>
            typeof initialValue === "object" && !Array.isArray(initialValue)
              ? { ...prev, ...fetchedVal }
              : fetchedVal,
          );
        }
        setIsLoaded(true);
      })
      .catch((e) => {
        console.error(e);
        setIsLoaded(true);
      });
  }, [key, user]);

  useEffect(() => {
    if (!isLoaded) return;

    const localTimeout = setTimeout(() => {
      try {
        localStorage.setItem(key, JSON.stringify(state));
      } catch (e) {
        console.warn("Storage warning:", e);
      }
    }, 500);

    let dbTimeout;
    if (user && db) {
      setIsSaving(true);
      dbTimeout = setTimeout(() => {
        const docRef = doc(
          db,
          "artifacts",
          appId,
          "users",
          user.uid,
          "global",
          "settings",
        );
        setDoc(docRef, { [key]: state }, { merge: true })
          .then(() => setIsSaving(false))
          .catch((e) => {
            console.error(e);
            setIsSaving(false);
          });
      }, 1500);
    }
    return () => {
      clearTimeout(localTimeout);
      if (dbTimeout) clearTimeout(dbTimeout);
    };
  }, [state, isLoaded, key, user]);

  useEffect(() => {
    const handleVisibility = () => {
      if (document.visibilityState === "hidden" && user && db) {
        const docRef = doc(
          db,
          "artifacts",
          appId,
          "users",
          user.uid,
          "global",
          "settings",
        );
        setDoc(docRef, { [key]: stateRef.current }, { merge: true });
      }
    };
    window.addEventListener("visibilitychange", handleVisibility);
    return () =>
      window.removeEventListener("visibilitychange", handleVisibility);
  }, [key, user]);

  return [state, setState, isLoaded, isSaving];
};

const useProjectState = (projectId, baseKey, initialValue, user) => {
  const localKey = `xdifica_proj_${projectId}_${baseKey}`;
  const [state, setState] = useState(() => {
    const localData = localStorage.getItem(localKey);
    if (localData) {
      try {
        const parsed = JSON.parse(localData);
        return typeof initialValue === "object" && !Array.isArray(initialValue)
          ? { ...initialValue, ...parsed }
          : parsed;
      } catch (e) {
        return initialValue;
      }
    }
    return initialValue;
  });
  const [isLoaded, setIsLoaded] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const stateRef = useRef(state);
  useEffect(() => {
    stateRef.current = state;
  }, [state]);

  useEffect(() => {
    if (!user || !db) {
      setIsLoaded(true);
      return;
    }
    const docRef = doc(
      db,
      "artifacts",
      appId,
      "users",
      user.uid,
      "projects",
      projectId,
    );
    getDoc(docRef)
      .then((snap) => {
        if (snap.exists() && snap.data()[baseKey] !== undefined) {
          const fetchedVal = snap.data()[baseKey];
          setState((prev) =>
            typeof initialValue === "object" && !Array.isArray(initialValue)
              ? { ...prev, ...fetchedVal }
              : fetchedVal,
          );
        }
        setIsLoaded(true);
      })
      .catch((e) => {
        console.error(e);
        setIsLoaded(true);
      });
  }, [projectId, baseKey, user]);

  useEffect(() => {
    if (!isLoaded) return;

    const localTimeout = setTimeout(() => {
      try {
        localStorage.setItem(localKey, JSON.stringify(state));
      } catch (e) {
        console.warn("Storage warning:", e);
      }
    }, 500);

    let dbTimeout;
    if (user && db) {
      setIsSaving(true);
      dbTimeout = setTimeout(() => {
        const docRef = doc(
          db,
          "artifacts",
          appId,
          "users",
          user.uid,
          "projects",
          projectId,
        );
        setDoc(docRef, { [baseKey]: state }, { merge: true })
          .then(() => setIsSaving(false))
          .catch((e) => {
            console.error(e);
            setIsSaving(false);
          });
      }, 1500);
    }
    return () => {
      clearTimeout(localTimeout);
      if (dbTimeout) clearTimeout(dbTimeout);
    };
  }, [state, isLoaded, projectId, baseKey, user, localKey]);

  useEffect(() => {
    const handleVisibility = () => {
      if (document.visibilityState === "hidden" && user && db) {
        const docRef = doc(
          db,
          "artifacts",
          appId,
          "users",
          user.uid,
          "projects",
          projectId,
        );
        setDoc(docRef, { [baseKey]: stateRef.current }, { merge: true });
      }
    };
    window.addEventListener("visibilitychange", handleVisibility);
    return () =>
      window.removeEventListener("visibilitychange", handleVisibility);
  }, [projectId, baseKey, user]);

  return [state, setState, isLoaded, isSaving];
};

const DebouncedCell = ({
  value,
  onChange,
  onDoubleClick,
  title,
  isTextArea = false,
  className = "",
  rows = 1,
  placeholder = "",
  readOnly = false,
  type = "text",
  min,
  step,
  allowMath = false,
}) => {
  const evaluateExpression = (expr) => {
    if (typeof expr !== "string") return expr;
    let cleanExpr = expr.trim();
    if (cleanExpr.startsWith("=")) cleanExpr = cleanExpr.slice(1);
    // Allow numbers, basic math operators, parentheses, and decimals.
    if (!/^[0-9+\-*/().\s]+$/.test(cleanExpr) || cleanExpr === "") return expr;
    try {
      const result = new Function(`"use strict"; return (${cleanExpr});`)();
      return isNaN(result) ? expr : result;
    } catch (e) {
      return expr;
    }
  };

  const formatValue = (val) => {
    if (val === null || val === undefined) return "";
    let v = val;
    if (allowMath && typeof v === "string") {
      v = evaluateExpression(v);
    }
    if ((type === "number" || allowMath) && step && v !== "") {
      const parsed = parseFloat(v);
      if (isNaN(parsed)) return val;
      if (step === "0.001") return parsed.toFixed(3);
      if (step === "0.01") return parsed.toFixed(2);
      return parsed.toString();
    }
    return val;
  };
  const [localValue, setLocalValue] = useState(formatValue(value));
  useEffect(() => {
    setLocalValue(formatValue(value));
  }, [value, type, step, allowMath]);
  const handleBlur = () => {
    let finalValue = localValue;
    if (allowMath && typeof finalValue === "string") {
      finalValue = evaluateExpression(finalValue);
    }
    if ((type === "number" || allowMath) && step && finalValue !== "") {
      const parsed = parseFloat(finalValue);
      if (!isNaN(parsed)) {
        if (step === "0.001") finalValue = parsed.toFixed(3);
        else if (step === "0.01") finalValue = parsed.toFixed(2);
        else finalValue = parsed.toString();
        setLocalValue(finalValue);
      }
    }
    if (String(finalValue) !== String(value)) onChange(finalValue);
  };
  if (isTextArea)
    return (
      <textarea
        className={`resize-y ${className}`}
        rows={rows}
        value={localValue !== undefined ? localValue : ""}
        onChange={(e) => setLocalValue(e.target.value)}
        onBlur={handleBlur}
        onFocus={(e) => e.target.select()}
        onDoubleClick={onDoubleClick}
        title={title}
        placeholder={placeholder}
        readOnly={readOnly}
        autoComplete="off"
        spellCheck="false"
      />
    );
  return (
    <input
      type={allowMath ? "text" : type}
      min={min}
      step={step}
      className={className}
      value={localValue !== undefined ? localValue : ""}
      onChange={(e) => setLocalValue(e.target.value)}
      onBlur={handleBlur}
      onFocus={(e) => e.target.select()}
      onDoubleClick={onDoubleClick}
      onKeyDown={(e) => {
        if (e.key === "Enter") e.target.blur();
      }}
      title={title}
      placeholder={placeholder}
      readOnly={readOnly}
      autoComplete="off"
      spellCheck="false"
    />
  );
};

const GeneratorRow = React.memo(
  ({
    row,
    index,
    updateRow,
    deleteRow,
    calculateVolume,
    isSelected,
    onToggleSelect,
  }) => {
    const [localEje, setLocalEje] = useState(row.eje || "");
    const [localClave, setLocalClave] = useState(row.claveLoc || "");
    useEffect(() => {
      setLocalEje(row.eje || "");
      setLocalClave(row.claveLoc || "");
    }, [row.eje, row.claveLoc]);
    const handleBlur = (field, value) => {
      if (row[field] !== value) updateRow(row.id, field, value);
    };
    return (
      <tr
        className={`transition-colors ${isSelected ? "bg-blue-50/50" : "hover:bg-slate-50"}`}
      >
        <td className="px-2 py-1.5 border-r border-slate-300 text-center relative">
          <div className="flex items-center justify-center gap-2">
            <button
              onClick={() => onToggleSelect(row.id)}
              className={`${isSelected ? "text-blue-600" : "text-slate-300 hover:text-slate-400"}`}
            >
              {isSelected ? <CheckSquare size={16} /> : <Square size={16} />}
            </button>
            <span className="font-black text-slate-300 text-[10px] w-4">
              {index + 1}
            </span>
          </div>
        </td>
        <td className="px-2 py-1.5 border-r border-slate-300">
          <div className="flex gap-2 group/row">
            <input
              placeholder="Eje/Ref"
              className="w-[40%] p-1.5 bg-white border border-slate-200 rounded text-[11px] font-bold uppercase outline-none focus:ring-1 focus:ring-blue-500"
              value={localEje}
              onChange={(e) => setLocalEje(e.target.value)}
              onBlur={(e) => handleBlur("eje", e.target.value)}
            />
            <input
              placeholder="Clave"
              className="w-[60%] p-1.5 bg-white border border-slate-200 rounded text-[11px] font-bold uppercase outline-none focus:ring-1 focus:ring-blue-500"
              value={localClave}
              onChange={(e) => setLocalClave(e.target.value)}
              onBlur={(e) => handleBlur("claveLoc", e.target.value)}
            />
          </div>
        </td>
        <td className="px-2 py-1.5 border-r border-slate-300 text-center font-black">
          <DebouncedCell
            type="number"
            allowMath={true}
            step="0.01"
            className="w-full p-1 bg-transparent text-center outline-none text-[11px]"
            value={row.largo}
            onChange={(v) => updateRow(row.id, "largo", v)}
          />
        </td>
        <td className="px-2 py-1.5 border-r border-slate-300 text-center font-black">
          <DebouncedCell
            type="number"
            allowMath={true}
            step="0.01"
            className="w-full p-1 bg-transparent text-center outline-none text-[11px]"
            value={row.ancho}
            onChange={(v) => updateRow(row.id, "ancho", v)}
          />
        </td>
        <td className="px-2 py-1.5 border-r border-slate-300 text-center font-black">
          <DebouncedCell
            type="number"
            allowMath={true}
            step="0.001"
            className="w-full p-1 bg-transparent text-center outline-none text-[11px]"
            value={row.kg_ml}
            onChange={(v) => updateRow(row.id, "kg_ml", v)}
          />
        </td>
        <td className="px-2 py-1.5 border-r border-slate-300 text-center font-black">
          <DebouncedCell
            type="number"
            allowMath={true}
            step="0.01"
            className="w-full p-1 bg-transparent text-center outline-none text-[11px]"
            value={row.alto}
            onChange={(v) => updateRow(row.id, "alto", v)}
          />
        </td>
        <td className="px-2 py-1.5 border-r border-slate-300 text-center font-black bg-slate-50/50 text-blue-800 text-[11px]">
          {(() => {
            const l = parseFloat(row.largo);
            const lAbs = isNaN(l) || l === 0 ? 0 : Math.abs(l);
            const a = parseFloat(row.ancho);
            const aAbs = isNaN(a) || a === 0 ? 0 : Math.abs(a);
            const h = parseFloat(row.alto);
            const hAbs = isNaN(h) || h === 0 ? 0 : Math.abs(h);
            const k = parseFloat(row.kg_ml);
            const kAbs = isNaN(k) || k === 0 ? 0 : Math.abs(k);
            const isActuallyZero =
              lAbs === 0 && aAbs === 0 && hAbs === 0 && kAbs === 0;
            const vol = isActuallyZero
              ? 0
              : (lAbs > 0 ? lAbs : 1) *
                (aAbs > 0 ? aAbs : 1) *
                (hAbs > 0 ? hAbs : 1) *
                (kAbs > 0 ? kAbs : 1);
            return vol.toLocaleString(undefined, {
              minimumFractionDigits: 2,
              maximumFractionDigits: 2,
            });
          })()}
        </td>
        <td className="px-2 py-1.5 border-r border-slate-300 text-center font-black">
          <DebouncedCell
            type="number"
            allowMath={true}
            className="w-full p-1 bg-transparent text-center outline-none text-[11px]"
            value={row.piezas}
            onChange={(v) => updateRow(row.id, "piezas", v)}
          />
        </td>
        <td className="px-2 py-1.5 border-r border-slate-300 text-right pr-4 font-black bg-slate-50/50 text-blue-900 text-[12px]">
          {calculateVolume(row).toLocaleString(undefined, {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2,
          })}
        </td>
        <td className="px-2 py-1.5 text-center">
          <button
            onClick={() => deleteRow(row.id)}
            className="text-slate-300 hover:text-red-500 transition-colors p-1"
          >
            <Trash2 size={16} />
          </button>
        </td>
      </tr>
    );
  },
);

const POLIN_DATA = {
  '3"': { dim: '3" x 1.5" x 0.5"', pesos: { "16": 2.08, "14": 2.61 } },
  '4"': {
    dim: '4" x 2.0" x 0.75"',
    pesos: { "16": 2.71, "14": 3.4, "12": 4.8, "10": 5.9 },
  },
  '5"': {
    dim: '5" x 2.0" x 0.75"',
    pesos: { "16": 3.0, "14": 3.8, "12": 5.3, "10": 6.6 },
  },
  '6"': {
    dim: '6" x 2.0" x 0.75"',
    pesos: { "16": 4.2, "14": 4.19, "12": 5.86, "10": 7.3 },
  },
  '7"': {
    dim: '7" x 2.75" x 0.75"',
    pesos: { "16": 4.2, "14": 5.3, "10": 9.2 },
  },
  '8"': {
    dim: '8" x 2.75" x 0.75"',
    pesos: { "16": 4.5, "14": 5.7, "12": 7.96, "10": 9.9 },
  },
  '10"': { dim: '10" x 2.75"', pesos: { "14": 6.49, "12": 9.06, "10": 11.2 } },
  '12"': {
    dim: '12" x 3.5" x 0.75"',
    pesos: { "14": 7.87, "12": 10.99, "10": 13.6 },
  },
};

const IPR_DATA = {
  '4" X 4"': [{ lb: 13.0, kg: 19.3, h: 105.0 }],
  '5" X 5"': [
    { lb: 16.0, kg: 23.8, h: 127.0 },
    { lb: 19.0, kg: 28.1, h: 131.0 },
  ],
  '6" X 4"': [
    { lb: 9.0, kg: 13.4, h: 150.0 },
    { lb: 12.0, kg: 17.9, h: 153.0 },
    { lb: 16.0, kg: 23.8, h: 152.0 },
  ],
  '6" X 6"': [
    { lb: 15.0, kg: 22.3, h: 152.0 },
    { lb: 20.0, kg: 29.8, h: 157.0 },
    { lb: 25.0, kg: 37.2, h: 162.0 },
  ],
  '8" X 4"': [
    { lb: 10.0, kg: 14.9, h: 200.0 },
    { lb: 13.0, kg: 19.3, h: 203.0 },
    { lb: 15.0, kg: 22.3, h: 206.0 },
  ],
  '8" X 5 1/4"': [
    { lb: 14.0, kg: 20.8, h: 203.0 },
    { lb: 18.0, kg: 26.8, h: 207.0 },
    { lb: 21.0, kg: 31.3, h: 210.0 },
  ],
  '8" X 6 1/2"': [
    { lb: 24.0, kg: 35.7, h: 201.0 },
    { lb: 28.0, kg: 41.7, h: 205.0 },
  ],
  '8" X 8"': [
    { lb: 31.0, kg: 46.1, h: 203.0 },
    { lb: 35.0, kg: 52.1, h: 206.0 },
    { lb: 67.0, kg: 99.7, h: 229.0 },
  ],
  '10" X 4"': [
    { lb: 12.0, kg: 17.9, h: 251.0 },
    { lb: 15.0, kg: 22.3, h: 254.0 },
    { lb: 17.0, kg: 25.3, h: 257.0 },
    { lb: 19.0, kg: 28.3, h: 260.0 },
  ],
  '10" X 5 3/4"': [
    { lb: 18.0, kg: 23.8, h: 253.0 },
    { lb: 22.0, kg: 32.7, h: 258.0 },
    { lb: 26.0, kg: 38.7, h: 262.0 },
    { lb: 30.0, kg: 44.6, h: 266.0 },
  ],
  '10" X 8"': [
    { lb: 33.0, kg: 49.1, h: 247.0 },
    { lb: 39.0, kg: 58.0, h: 252.0 },
    { lb: 45.0, kg: 67.0, h: 257.0 },
  ],
  '10" X 10"': [
    { lb: 49.0, kg: 72.9, h: 253.0 },
    { lb: 54.0, kg: 80.4, h: 256.0 },
    { lb: 60.0, kg: 89.3, h: 260.0 },
    { lb: 68.0, kg: 101.2, h: 264.0 },
    { lb: 77.0, kg: 114.6, h: 269.0 },
    { lb: 88.0, kg: 131.0, h: 275.0 },
    { lb: 100.0, kg: 148.8, h: 282.0 },
    { lb: 112.0, kg: 166.7, h: 289.0 },
  ],
  '12" X 4"': [
    { lb: 14.0, kg: 20.8, h: 303.0 },
    { lb: 16.0, kg: 23.8, h: 305.0 },
    { lb: 19.0, kg: 28.3, h: 309.0 },
    { lb: 22.0, kg: 32.7, h: 313.0 },
  ],
  '12" X 6 1/2"': [
    { lb: 26.0, kg: 38.7, h: 310.0 },
    { lb: 30.0, kg: 44.6, h: 313.0 },
    { lb: 35.0, kg: 52.1, h: 317.0 },
  ],
  '12" X 8"': [
    { lb: 40.0, kg: 59.5, h: 303.0 },
    { lb: 45.0, kg: 67.0, h: 306.0 },
    { lb: 50.0, kg: 74.4, h: 310.0 },
  ],
  '12" X 10"': [
    { lb: 53.0, kg: 78.9, h: 306.0 },
    { lb: 58.0, kg: 86.3, h: 310.0 },
  ],
  '12" X 12"': [
    { lb: 65.0, kg: 96.7, h: 308.0 },
    { lb: 72.0, kg: 107.1, h: 311.0 },
    { lb: 79.0, kg: 117.6, h: 314.0 },
    { lb: 87.0, kg: 129.5, h: 318.0 },
    { lb: 96.0, kg: 142.9, h: 323.0 },
    { lb: 106.0, kg: 157.7, h: 327.0 },
    { lb: 120.0, kg: 178.6, h: 333.0 },
    { lb: 136.0, kg: 202.37, h: 341.0 },
    { lb: 152.0, kg: 226.2, h: 348.0 },
    { lb: 170.0, kg: 253.0, h: 356.0 },
    { lb: 190.0, kg: 282.8, h: 365.0 },
  ],
  '12" X 12 1/2"': [
    { lb: 210.0, kg: 312.5, h: 374.0 },
    { lb: 230.0, kg: 342.3, h: 382.0 },
    { lb: 252.0, kg: 375.0, h: 391.0 },
    { lb: 279.0, kg: 415.2, h: 403.0 },
    { lb: 305.0, kg: 453.9, h: 415.0 },
    { lb: 336.0, kg: 500.0, h: 427.0 },
  ],
  '14" X 5"': [
    { lb: 22.0, kg: 32.7, h: 349.0 },
    { lb: 26.0, kg: 38.7, h: 353.0 },
  ],
  '14" X 6 3/4"': [
    { lb: 30.0, kg: 44.6, h: 352.0 },
    { lb: 34.0, kg: 50.6, h: 355.0 },
    { lb: 38.0, kg: 56.5, h: 358.0 },
  ],
  '14" X 8"': [
    { lb: 43.0, kg: 64.0, h: 347.0 },
    { lb: 48.0, kg: 71.4, h: 350.0 },
    { lb: 53.0, kg: 78.9, h: 354.0 },
  ],
  '14" X 10"': [
    { lb: 61.0, kg: 90.8, h: 353.0 },
    { lb: 68.0, kg: 101.2, h: 357.0 },
    { lb: 74.0, kg: 110.1, h: 360.0 },
    { lb: 82.0, kg: 122.0, h: 363.0 },
  ],
  '14" X 14 1/2"': [
    { lb: 90.0, kg: 133.9, h: 356.0 },
    { lb: 99.0, kg: 147.3, h: 360.0 },
    { lb: 109.0, kg: 162.2, h: 364.0 },
    { lb: 120.0, kg: 178.6, h: 368.0 },
    { lb: 132.0, kg: 196.4, h: 372.0 },
  ],
  '14" X 16"': [
    { lb: 145.0, kg: 215.8, h: 375.0 },
    { lb: 159.0, kg: 236.6, h: 380.0 },
    { lb: 176.0, kg: 261.9, h: 387.0 },
    { lb: 193.0, kg: 287.2, h: 393.0 },
    { lb: 211.0, kg: 314.0, h: 399.0 },
    { lb: 233.0, kg: 346.7, h: 407.0 },
    { lb: 257.0, kg: 382.5, h: 416.0 },
    { lb: 283.0, kg: 421.1, h: 425.0 },
    { lb: 311.0, kg: 462.8, h: 435.0 },
    { lb: 342.0, kg: 509.0, h: 446.0 },
    { lb: 370.0, kg: 550.5, h: 455.0 },
    { lb: 398.0, kg: 592.3, h: 465.0 },
    { lb: 426.0, kg: 634.0, h: 474.0 },
  ],
  '16" X 5 1/2"': [
    { lb: 26.0, kg: 38.7, h: 399.0 },
    { lb: 31.0, kg: 46.1, h: 403.0 },
  ],
  '16" X 7"': [
    { lb: 36.0, kg: 53.6, h: 403.0 },
    { lb: 40.0, kg: 59.5, h: 407.0 },
    { lb: 45.0, kg: 67.0, h: 410.0 },
    { lb: 50.0, kg: 74.4, h: 413.0 },
    { lb: 57.0, kg: 84.8, h: 417.0 },
  ],
  '16" X 10 1/4"': [
    { lb: 67.0, kg: 99.7, h: 415.0 },
    { lb: 77.0, kg: 114.6, h: 420.0 },
    { lb: 89.0, kg: 132.4, h: 425.0 },
  ],
  '18" X 7 1/2"': [{ lb: 71.0, kg: 105.7, h: 469.0 }],
  '18" X 11"': [
    { lb: 76.0, kg: 113.1, h: 463.0 },
    { lb: 86.0, kg: 128.0, h: 467.0 },
    { lb: 97.0, kg: 144.4, h: 472.0 },
    { lb: 106.0, kg: 157.7, h: 476.0 },
    { lb: 119.0, kg: 177.1, h: 482.0 },
    { lb: 130.0, kg: 193.5, h: 485.0 },
    { lb: 143.0, kg: 212.8, h: 495.0 },
    { lb: 158.0, kg: 235.1, h: 501.0 },
    { lb: 175.0, kg: 260.4, h: 508.0 },
    { lb: 192.0, kg: 285.7, h: 517.0 },
    { lb: 211.0, kg: 314.0, h: 525.0 },
    { lb: 234.0, kg: 348.2, h: 535.0 },
    { lb: 258.0, kg: 383.9, h: 545.0 },
    { lb: 283.0, kg: 421.1, h: 555.0 },
    { lb: 311.0, kg: 462.8, h: 567.0 },
  ],
  '21" X 6 1/2"': [
    { lb: 44.0, kg: 65.5, h: 525.0 },
    { lb: 50.0, kg: 74.4, h: 529.0 },
    { lb: 57.0, kg: 84.8, h: 535.0 },
  ],
  '21" X 8 1/4"': [
    { lb: 62.0, kg: 92.3, h: 533.0 },
    { lb: 68.0, kg: 101.2, h: 537.0 },
    { lb: 73.0, kg: 108.6, h: 539.0 },
    { lb: 83.0, kg: 123.5, h: 544.0 },
    { lb: 93.0, kg: 138.4, h: 549.0 },
  ],
  '21" X 12 1/2"': [
    { lb: 101.0, kg: 150.3, h: 543.0 },
    { lb: 111.0, kg: 165.2, h: 546.0 },
    { lb: 122.0, kg: 181.5, h: 551.0 },
    { lb: 132.0, kg: 196.4, h: 554.0 },
    { lb: 147.0, kg: 218.8, h: 560.0 },
    { lb: 166.0, kg: 247.0, h: 571.0 },
    { lb: 182.0, kg: 270.8, h: 577.0 },
    { lb: 201.0, kg: 299.1, h: 585.0 },
  ],
  '24" X 7"': [{ lb: 55.0, kg: 81.8, h: 599.0 }],
};

const HSS_DATA = {
  '4" x 4"': {
    dim: "101.6 x 101.6",
    pesos: {
      '1/8"': 9.61,
      '3/16"': 14.02,
      '1/4"': 18.17,
      '5/16"': 22.07,
      '3/8"': 25.7,
      '1/2"': 32.19,
    },
  },
  '4 1/2" x 4 1/2"': {
    dim: "114.3 x 114.3",
    pesos: {
      '1/8"': 10.86,
      '3/16"': 15.92,
      '1/4"': 20.7,
      '5/16"': 25.24,
      '3/8"': 29.5,
      '1/2"': 37.25,
    },
  },
  '5" x 5"': {
    dim: "127.0 x 127.0",
    pesos: {
      '1/8"': 12.13,
      '3/16"': 17.81,
      '1/4"': 23.25,
      '5/16"': 28.39,
      '3/8"': 33.29,
      '1/2"': 42.31,
    },
  },
  '5 1/2" x 5 1/2"': {
    dim: "139.7 x 139.7",
    pesos: { '3/16"': 19.72, '1/4"': 25.77, '5/16"': 31.56, '3/8"': 37.1 },
  },
  '6" x 6"': {
    dim: "152.4 x 152.4",
    pesos: {
      '3/16"': 21.62,
      '1/4"': 28.3,
      '5/16"': 34.73,
      '3/8"': 40.89,
      '1/2"': 52.44,
      '5/8"': 63.1,
    },
  },
  '7" x 7"': {
    dim: "177.8 x 177.8",
    pesos: {
      '3/16"': 25.42,
      '1/4"': 33.36,
      '5/16"': 41.06,
      '3/8"': 48.48,
      '1/2"': 62.58,
    },
  },
  '7 1/2" x 7 1/2"': {
    dim: "190.5 x 190.5",
    pesos: {
      '3/16"': 27.31,
      '1/4"': 35.89,
      '5/16"': 44.23,
      '3/8"': 52.28,
      '1/2"': 67.64,
    },
  },
  '8" x 8"': {
    dim: "203.2 x 203.2",
    pesos: {
      '3/16"': 29.21,
      '1/4"': 38.42,
      '5/16"': 47.38,
      '3/8"': 56.09,
      '1/2"': 72.7,
      '5/8"': 88.4,
    },
  },
  '9" x 9"': {
    dim: "228.6 x 228.6",
    pesos: {
      '3/16"': 33.01,
      '1/4"': 43.5,
      '5/16"': 53.72,
      '3/8"': 63.68,
      '1/2"': 82.83,
    },
  },
  '10" x 10"': {
    dim: "254.0 x 254.0",
    pesos: {
      '3/16"': 36.8,
      '1/4"': 48.56,
      '5/16"': 60.05,
      '3/8"': 71.28,
      '1/2"': 92.95,
      '5/8"': 113.8,
    },
  },
  '12" x 12"': {
    dim: "304.8 x 304.8",
    pesos: {
      '3/16"': 44.41,
      '1/4"': 58.68,
      '5/16"': 72.71,
      '3/8"': 86.46,
      '1/2"': 113.2,
      '5/8"': 139.1,
    },
  },
  '14" x 14"': {
    dim: "355.6 x 355.6",
    pesos: { '5/16"': 85.36, '3/8"': 101.66, '1/2"': 133.46, '5/8"': 166.82 },
  },
  '16" x 16"': {
    dim: "406.4 x 406.4",
    pesos: { '5/16"': 98.03, '3/8"': 116.85, '1/2"': 153.73, '5/8"': 189.01 },
  },
};

const HSS_RECT_DATA = {
  '8" x 4"': {
    dim: "203.2 x 101.6",
    pesos: {
      '3/16"': 21.64,
      '1/4"': 28.36,
      '5/16"': 34.75,
      '3/8"': 40.98,
      '1/2"': 52.46,
    },
  },
  '8" x 6"': { dim: "203.2 x 152.4", pesos: { '1/4"': 33.44, '3/8"': 48.52 } },
  '10" x 4"': {
    dim: "254.0 x 101.6",
    pesos: { '3/16"': 25.42, '1/4"': 33.36 },
  },
  '10" x 6"': {
    dim: "254.0 x 152.4",
    pesos: {
      '3/16"': 29.34,
      '1/4"': 38.52,
      '5/16"': 47.54,
      '3/8"': 56.23,
      '1/2"': 72.79,
    },
  },
  '12" x 4"': {
    dim: "304.8 x 101.6",
    pesos: { '3/16"': 29.21, '1/4"': 38.42, '3/8"': 56.09 },
  },
  '12" x 6"': {
    dim: "304.8 x 152.4",
    pesos: { '3/16"': 33.01, '1/4"': 43.5, '3/8"': 63.77 },
  },
  '12" x 8"': {
    dim: "304.8 x 203.2",
    pesos: { '1/4"': 48.69, '3/8"': 71.31, '1/2"': 92.95, '5/8"': 113.59 },
  },
  '14" x 6"': {
    dim: "355.6 x 152.4",
    pesos: { '5/16"': 60.05, '3/8"': 71.28 },
  },
};

const PTR_DATA = {
  '1" x 1"': {
    dim: "25.4 x 25.4",
    pesos: { "14": 1.46, "12": 1.99, "11": 2.24, "10": 2.49 },
  },
  '1 1/4" x 1 1/4"': {
    dim: "31.75 x 31.75",
    pesos: { "16": 1.48, "14": 1.87, "12": 2.57, "10": 3.2 },
  },
  '1 1/2" x 1 1/2"': {
    dim: "38.1 x 38.1",
    pesos: { "16": 1.84, "14": 2.27, "12": 3.12, "11": 3.5, "10": 3.91 },
  },
  '2" x 2"': {
    dim: "50.8 x 50.8",
    pesos: { "14": 3.02, "12": 4.18, "11": 4.75, "10": 5.31 },
  },
  '2 1/2" x 2 1/2"': {
    dim: "64 x 64",
    pesos: {
      "14": 3.8,
      "12": 5.27,
      "11": 5.8,
      "10": 6.17,
      "9": 6.75,
      "8": 7.44,
      '3/16"': 8.31,
    },
  },
  '3" x 3"': {
    dim: "76 x 76",
    pesos: {
      "14": 4.35,
      "11": 7.06,
      "10": 7.54,
      "9": 8.25,
      "8": 9.1,
      '3/16"': 10.21,
    },
  },
  '3 1/2" x 3 1/2"': {
    dim: "89 x 89",
    pesos: {
      "14": 5.11,
      "11": 8.32,
      "10": 8.89,
      "9": 9.73,
      "8": 10.76,
      '3/16"': 12.11,
      "5": 13.97,
      '1/4"': 15.62,
    },
  },
  '4" x 4"': {
    dim: "102 x 102",
    pesos: {
      "14": 5.87,
      "11": 9.6,
      "10": 10.24,
      "9": 11.24,
      "8": 12.44,
      '3/16"': 14.0,
      "5": 16.19,
      '1/4"': 18.15,
    },
  },
  '4 1/2" x 4 1/2"': {
    dim: "114 x 114",
    pesos: {
      "14": 6.63,
      "11": 10.24,
      "10": 11.52,
      "9": 13.02,
      "8": 14.03,
      '3/16"': 15.9,
      "5": 18.4,
      '1/4"': 20.67,
    },
  },
  '5" x 5"': {
    dim: "127 x 127",
    pesos: {
      "14": 7.39,
      "11": 12.13,
      "10": 13.18,
      "9": 14.22,
      "8": 16.13,
      '3/16"': 17.79,
      "5": 20.66,
      '1/4"': 23.21,
    },
  },
};

const CANALES_C_DATA = {
  '3"': { dim: "76.2 mm", pesos: { "6.10": 6.1, "7.44": 7.44 } },
  '4"': { dim: "101.6 mm", pesos: { "8.04": 8.04, "10.8": 10.8 } },
  '5"': { dim: "127.0 mm", pesos: { "9.97": 9.97, "13.39": 13.39 } },
  '6"': {
    dim: "152.4 mm",
    pesos: { "12.20": 12.2, "15.62": 15.62, "19.34": 19.34 },
  },
  '8"': {
    dim: "203.2 mm",
    pesos: { "17.11": 17.11, "20.46": 20.46, "27.90": 27.9 },
  },
  '10"': {
    dim: "254.0 mm",
    pesos: { "22.76": 22.76, "30.00": 30.0, "45.00": 45.0 },
  },
  '12"': {
    dim: "304.8 mm",
    pesos: { "30.80": 30.8, "37.20": 37.2, "44.60": 44.6 },
  },
};

const VIGA_IR_DATA = {
  "152": [12.7, 13.6, 18.0, 24.0, 22.4, 29.7, 37.2],
  "203": [
    15.0, 19.4, 22.5, 26.6, 31.2, 35.9, 41.8, 46.2, 52.2, 59.3, 71.4, 86.6,
    99.8,
  ],
  "254": [
    17.9, 22.3, 25.3, 28.5, 32.9, 38.5, 44.8, 49.2, 58.2, 67.4, 72.9, 80.0,
    89.1, 101.3, 114.5, 131.2, 148.9, 166.6,
  ],
  "305": [
    21.1, 23.9, 28.2, 32.8, 38.7, 44.5, 52.2, 59.8, 66.9, 74.4, 79.0, 86.1,
    96.7, 106.9, 117.5, 129.7, 142.8, 158.0, 178.8,
  ],
  "356": [
    32.9, 38.9, 44.8, 50.6, 56.7, 63.8, 71.4, 79.0, 90.7, 101.3, 110.4, 122.1,
  ],
  "406": [38.9, 46.2, 53.7, 59.8, 67.4, 74.4],
};

const PERFILES_OC_DATA = {
  '20"': { dim: "508.0 mm", pesos: { '0.500"': 229, '0.375"': 173 } },
  '18"': { dim: "457.2 mm", pesos: { '0.500"': 206, '0.375"': 156 } },
  '16"': {
    dim: "406.4 mm",
    pesos: {
      '0.625"': 227,
      '0.500"': 183,
      '0.438"': 161,
      '0.375"': 138,
      '0.312"': 115,
      '0.250"': 92.8,
    },
  },
  '14"': {
    dim: "355.6 mm",
    pesos: {
      '0.625"': 197,
      '0.500"': 159,
      '0.375"': 120,
      '0.312"': 101,
      '0.250"': 81.0,
    },
  },
  '12.750"': {
    dim: "323.8 mm",
    pesos: { '0.500"': 144, '0.375"': 109, '0.250"': 73.6 },
  },
  '10.750"': {
    dim: "273.0 mm",
    pesos: { '0.500"': 121, '0.375"': 91.7, '0.250"': 61.9 },
  },
  '10"': {
    dim: "254.0 mm",
    pesos: {
      '0.625"': 138,
      '0.500"': 112,
      '0.375"': 85.1,
      '0.312"': 71.2,
      '0.250"': 57.4,
      '0.188"': 43.5,
    },
  },
  '9.625"': {
    dim: "244.5 mm",
    pesos: {
      '0.500"': 108,
      '0.375"': 81.8,
      '0.312"': 68.5,
      '0.250"': 55.2,
      '0.188"': 41.8,
    },
  },
  '8.625"': {
    dim: "219.1 mm",
    pesos: {
      '0.625"': 118,
      '0.500"': 95.7,
      '0.375"': 72.9,
      '0.322"': 63.0,
      '0.250"': 49.3,
      '0.188"': 37.4,
    },
  },
  '7.625"': { dim: "193.7 mm", pesos: { '0.375"': 64.1, '0.328"': 56.4 } },
  '7.500"': {
    dim: "190.5 mm",
    pesos: {
      '0.500"': 82.5,
      '0.375"': 63.0,
      '0.312"': 52.9,
      '0.250"': 42.7,
      '0.188"': 32.4,
    },
  },
  '7"': { dim: "177.8 mm", pesos: { '0.500"': 76.6, '0.375"': 58.6 } },
};

const LAMINA_DATA = {
  "RN-100/35": { pesos: { "26": 4.69, "24": 5.42, "22": 7.6 } },
  "R-72": {
    pesos: { "30": 2.44, "28": 2.97, "26": 3.52, "24": 4.06, "22": 5.7 },
  },
  "R-101": { pesos: { "28": 3.46, "26": 4.69, "24": 5.42, "22": 7.6 } },
  "O-725": { pesos: { "30": 2.44, "28": 2.97, "26": 3.52, "24": 4.06 } },
  "O-100": { pesos: { "28": 3.96, "26": 4.69, "24": 5.42 } },
  "SECC-25": { pesos: { "24": 5.42, "22": 7.6, "20": 9.06, "18": 11.96 } },
  "R-90": { pesos: { "26": 5.21, "24": 6.02, "22": 8.44 } },
};

const PLACAS_DATA = {
  '1"': 199.21,
  '15/16"': 186.75,
  '7/8"': 174.31,
  '13/16"': 161.85,
  '3/4"': 149.41,
  '11/16"': 136.96,
  '5/8"': 124.51,
  '9/16"': 112.06,
  '1/2"': 99.61,
  '7/16"': 87.15,
  '3/8"': 74.7,
  '5/16"': 62.25,
  '1/4"': 49.8,
  '3/16"': 37.35,
};

function ProjectWorkspace({
  projectId,
  onBack,
  onUpdateMetadata,
  user,
  catalogoConceptos,
  setCatalogoConceptos,
  isCatalogoLoaded,
  isSavingGlobal,
}) {
  const [obraInfo, setObraInfo, l1, s1] = useProjectState(
    projectId,
    "obraInfo",
    {
      nombre: "NUEVA OBRA",
      ubicacion: "",
      fecha: new Date().toLocaleDateString(),
    },
    user,
  );
  const [partidas, setPartidas, l3, s3] = useProjectState(
    projectId,
    "partidas",
    [],
    user,
  );
  const [partidaToDelete, setPartidaToDelete] = useState(null);
  const [colWidths, setColWidths, l4, s4] = usePersistentState(
    "xdifica_colWidths_v4",
    {
      select: 30,
      id: 70,
      clave: 45,
      cc: 40,
      justificacion: 120,
      descripcion: 200,
      unidad: 50,
      niveles: 65,
      totales: 75,
    },
    user,
  );
  const [wallColWidths, setWallColWidths, l5, s5] = usePersistentState(
    "xdifica_wallColWidths",
    {
      no: 40,
      eje: 60,
      clave: 80,
      largo: 60,
      ancho: 60,
      alto: 60,
      bruto: 70,
      huecos: 70,
      castillos: 70,
      neta: 80,
      aplanadoC1: 90,
      aplanadoC2: 90,
      recubrimientoC1: 90,
      recubrimientoC2: 90,
      action: 40,
    },
    user,
  );
  const [structColWidths, setStructColWidths, l6, s6] = usePersistentState(
    "xdifica_structColWidths",
    {
      select: 40,
      no: 40,
      eje: 70,
      clave: 80,
      tipo: 140,
      i: 30,
      largo: 80,
      ancho: 80,
      alto: 80,
      piezas: 60,
      concreto: 100,
      cimbra: 100,
      cimbraFrontera: 100,
      acero: 100,
      action: 90,
    },
    user,
  );
  const [steelColWidths, setSteelColWidths, l7, s7] = usePersistentState(
    "xdifica_steelColWidths",
    {
      no: 35,
      tipo: 120,
      varilla: 60,
      tramo: 80,
      separacion: 60,
      piezas: 55,
      longitud: 75,
      ganchos: 65,
      anclaje: 65,
      traslapes: 65,
      mlPza: 70,
      totalMl: 70,
      totalKg: 70,
      action: 40,
    },
    user,
  );
  const [genColWidths, setGenColWidths, l8, s8] = usePersistentState(
    "xdifica_genColWidths",
    {
      select: 60,
      localizacion: 350,
      largo: 100,
      ancho: 100,
      kg_ml: 100,
      alto: 100,
      volPza: 100,
      piezas: 90,
      volumen: 130,
      action: 80,
    },
    user,
  );
  const [catColWidths, setCatColWidths, l10, s10] = usePersistentState(
    "xdifica_catColWidths",
    {
      id: 100,
      clave: 80,
      cc: 80,
      justificacion: 150,
      descripcion: 350,
      unidad: 80,
      action: 50,
    },
    user,
  );

  const [activePartidaId, setActivePartidaId] = useState(null);
  const [activeLevelId, setActiveLevelId] = useState(null);

  const [editingModal, setEditingModal] = useState(null);
  const [activeLevelReport, setActiveLevelReport] = useState(null);
  const [showSettings, setShowSettings] = useState(false);
  const [showWallGenerator, setShowWallGenerator] = useState(false);
  const [showStructureGenerator, setShowStructureGenerator] = useState(false);
  const [showMetalStructureGenerator, setShowMetalStructureGenerator] =
    useState(false);
  const [showAnalysisNervadura, setShowAnalysisNervadura] = useState(false);
  const [nervaduraAnalysisSelection, setNervaduraAnalysisSelection] = useState(
    [],
  );
  const [nervaduraAnalysisM2, setNervaduraAnalysisM2] = useState("");

  const [activeWallSubmodal, setActiveWallSubmodal] = useState(null);
  const [activeSteelSubmodal, setActiveSteelSubmodal] = useState(null);
  const [activeMetalSubmodal, setActiveMetalSubmodal] =
    useState("Polín Monten");
  const [activeCasetonSubmodal, setActiveCasetonSubmodal] = useState(null);
  const [conceptSelectorFor, setConceptSelectorFor] = useState(null);
  const [conceptSearchTerm, setConceptSearchTerm] = useState("");

  const [selectedIds, setSelectedIds] = useState([]);
  const [clipboardRows, setClipboardRows] = useState(null);
  const [copyStatus, setCopyStatus] = useState("");
  const [copiedStructureRow, setCopiedStructureRow] = useState(null);
  const [selectedGeneratorRows, setSelectedGeneratorRows] = useState([]);
  const [selectedStructureRows, setSelectedStructureRows] = useState([]);
  const [selectedMetalRows, setSelectedMetalRows] = useState([]);
  const [copiedStructureRowsList, setCopiedStructureRowsList] = useState(null);
  const [selectedWallRows, setSelectedWallRows] = useState([]);
  const [copiedWallRowsList, setCopiedWallRowsList] = useState(null);
  const [copiedWallRow, setCopiedWallRow] = useState(null);
  const [showGlobalSteelReport, setShowGlobalSteelReport] = useState(false);
  const [globalReportLevel, setGlobalReportLevel] = useState("todos");

  const [structAdjModal, setStructAdjModal] = useState(null);

  const isFullyLoaded =
    l1 && l3 && l4 && l5 && l6 && l7 && l8 && isCatalogoLoaded && l10;

  const isSavingAny =
    s1 || s3 || s4 || s5 || s6 || s7 || s8 || s10 || isSavingGlobal;

  useEffect(() => {
    const handleBeforeUnload = (e) => {
      if (isSavingAny) {
        e.preventDefault();
        e.returnValue = "";
      }
    };
    window.addEventListener("beforeunload", handleBeforeUnload);
    return () => window.removeEventListener("beforeunload", handleBeforeUnload);
  }, [isSavingAny]);

  const activePartida = useMemo(
    () => partidas.find((p) => p.id === activePartidaId),
    [partidas, activePartidaId],
  );

  const niveles = useMemo(() => {
    if (!activePartida) return [];
    return activePartida.niveles || [{ id: "n1", nombre: "N1 +1.00" }];
  }, [activePartida]);

  useEffect(() => {
    if (
      l3 &&
      niveles.length > 0 &&
      !niveles.find((n) => n.id === activeLevelId)
    ) {
      setActiveLevelId(niveles[0].id);
    }
  }, [niveles, activeLevelId, l3]);
  const conceptos = activePartida ? activePartida.conceptos : [];
  const generadores = activePartida ? activePartida.generadores : {};
  const murosData = activePartida ? activePartida.muros || {} : {};
  const muros = Array.isArray(murosData)
    ? murosData
    : murosData[activeLevelId] || [];
  const estructurasData = activePartida ? activePartida.estructuras || {} : {};
  const estructuras = Array.isArray(estructurasData)
    ? estructurasData
    : estructurasData[activeLevelId] || [];

  const metalData = activePartida ? activePartida.metal || {} : {};
  const metalEstructuras = Array.isArray(metalData)
    ? metalData
    : metalData[activeLevelId] || [];

  const todasLasEstructuras = useMemo(() => {
    if (!activePartida || !activePartida.estructuras) return [];
    const data = activePartida.estructuras;
    if (Array.isArray(data))
      return data.map((e) => ({
        ...e,
        _nivelInfo: "General",
        _nivelId: "general",
      }));
    let all = [];
    Object.keys(data).forEach((lvlId) => {
      const nivelName = niveles.find((n) => n.id === lvlId)?.nombre || lvlId;
      all = [
        ...all,
        ...(data[lvlId] || []).map((e) => ({
          ...e,
          _nivelInfo: nivelName,
          _nivelId: lvlId,
        })),
      ];
    });
    return all;
  }, [activePartida, niveles]);

  // Actualizadores de estado
  const updateActiveNiveles = useCallback(
    (upd) =>
      setPartidas((prev) =>
        prev.map((p) =>
          p.id === activePartidaId
            ? {
                ...p,
                niveles:
                  typeof upd === "function"
                    ? upd(p.niveles || [{ id: "n1", nombre: "N1 +1.00" }])
                    : upd,
              }
            : p,
        ),
      ),
    [activePartidaId, setPartidas],
  );

  const setNiveles = updateActiveNiveles;

  const updateActiveConceptos = useCallback(
    (upd) =>
      setPartidas((prev) =>
        prev.map((p) =>
          p.id === activePartidaId
            ? {
                ...p,
                conceptos: typeof upd === "function" ? upd(p.conceptos || []) : upd,
              }
            : p,
        ),
      ),
    [activePartidaId, setPartidas],
  );
  const updateActiveGeneradores = useCallback(
    (upd) =>
      setPartidas((prev) =>
        prev.map((p) =>
          p.id === activePartidaId
            ? {
                ...p,
                generadores:
                  typeof upd === "function" ? upd(p.generadores) : upd,
              }
            : p,
        ),
      ),
    [activePartidaId, setPartidas],
  );
  const updateActiveMuros = useCallback(
    (upd) => {
      setPartidas((prev) =>
        prev.map((p) => {
          if (p.id !== activePartidaId) return p;
          const data = p.muros || {};
          const isArr = Array.isArray(data);
          const nextData =
            typeof upd === "function"
              ? upd(isArr ? data : data[activeLevelId] || [])
              : upd;
          return {
            ...p,
            muros: isArr
              ? { [activeLevelId]: nextData }
              : { ...data, [activeLevelId]: nextData },
          };
        }),
      );
    },
    [activePartidaId, activeLevelId, setPartidas],
  );
  const updateActiveEstructuras = useCallback(
    (upd) => {
      setPartidas((prev) =>
        prev.map((p) => {
          if (p.id !== activePartidaId) return p;
          const data = p.estructuras || {};
          const isArr = Array.isArray(data);
          const nextData =
            typeof upd === "function"
              ? upd(isArr ? data : data[activeLevelId] || [])
              : upd;
          return {
            ...p,
            estructuras: isArr
              ? { [activeLevelId]: nextData }
              : { ...data, [activeLevelId]: nextData },
          };
        }),
      );
    },
    [activePartidaId, activeLevelId, setPartidas],
  );

  const saveStructAdjustment = useCallback(
    ({ tipoLosa, esquinaLosa, espesorLosa, espesorZapata }) => {
      if (!structAdjModal) return;
      const { id } = structAdjModal;

      updateActiveEstructuras((prev) =>
        prev.map((s) => {
          if (s.id !== id) return s;
          return {
            ...s,
            structTipoLosa: tipoLosa,
            structEsquinaLosa: esquinaLosa,
            structEspesorLosa: espesorLosa,
            structEspesorZapata: espesorZapata,
          };
        }),
      );
      setStructAdjModal(null);
    },
    [structAdjModal, updateActiveEstructuras],
  );

  const updateActiveMetalEstructuras = useCallback(
    (upd) => {
      setPartidas((prev) =>
        prev.map((p) => {
          if (p.id !== activePartidaId) return p;
          const data = p.metal || {};
          const isArr = Array.isArray(data);
          const nextData =
            typeof upd === "function"
              ? upd(isArr ? data : data[activeLevelId] || [])
              : upd;
          return {
            ...p,
            metal: isArr
              ? { [activeLevelId]: nextData }
              : { ...data, [activeLevelId]: nextData },
          };
        }),
      );
    },
    [activePartidaId, activeLevelId, setPartidas],
  );

  const handleCreatePartida = () =>
    setPartidas((prev) => [
      ...prev,
      {
        id: `P-${Date.now().toString(36).toUpperCase()}`,
        nombre: "NUEVA PARTIDA",
        niveles: [{ id: "n1", nombre: "N1 +1.00" }],
        conceptos: [],
        generadores: {},
        muros: {},
        estructuras: {},
      },
    ]);
  const handleDeletePartida = (id) => {
    setPartidaToDelete(id);
  };
  const confirmDeletePartida = (id) => {
    setPartidas((prev) => prev.filter((p) => p.id !== id));
    if (activePartidaId === id) setActivePartidaId(null);
    setPartidaToDelete(null);
  };
  const openPartida = (id) => {
    setActivePartidaId(id);
    setSelectedIds([]);
    setShowSettings(false);
    setShowWallGenerator(false);
    setShowStructureGenerator(false);
    setShowMetalStructureGenerator(false);
    setShowAnalysisNervadura(false);
    setActiveWallSubmodal(null);
    setActiveSteelSubmodal(null);
    setActiveCasetonSubmodal(null);
    setActiveLevelReport(null);
  };

  const resizingRef = useRef(null);
  const startResizing = (colId, e, type = "summary") => {
    e.preventDefault();
    let w = 0;
    if (type === "summary") w = colWidths[colId];
    else if (type === "walls") w = wallColWidths[colId];
    else if (type === "structures") w = structColWidths[colId];
    else if (type === "steel") w = steelColWidths[colId];
    else if (type === "generator") w = genColWidths[colId];
    else if (type === "catalog") w = catColWidths[colId];
    resizingRef.current = { colId, startX: e.pageX, startWidth: w, type };
    document.addEventListener("mousemove", handleMouseMove);
    document.addEventListener("mouseup", stopResizing);
  };
  const handleMouseMove = useCallback(
    (e) => {
      if (!resizingRef.current) return;
      const { colId, startX, startWidth, type } = resizingRef.current;
      const newWidth = Math.max(30, startWidth + (e.pageX - startX));
      if (type === "summary")
        setColWidths((p) => ({ ...p, [colId]: newWidth }));
      else if (type === "walls")
        setWallColWidths((p) => ({ ...p, [colId]: newWidth }));
      else if (type === "structures")
        setStructColWidths((p) => ({ ...p, [colId]: newWidth }));
      else if (type === "steel")
        setSteelColWidths((p) => ({ ...p, [colId]: newWidth }));
      else if (type === "generator")
        setGenColWidths((p) => ({ ...p, [colId]: newWidth }));
      else if (type === "catalog")
        setCatColWidths((p) => ({ ...p, [colId]: newWidth }));
    },
    [
      setColWidths,
      setWallColWidths,
      setStructColWidths,
      setSteelColWidths,
      setGenColWidths,
      setCatColWidths,
    ],
  );
  const stopResizing = useCallback(() => {
    resizingRef.current = null;
    document.removeEventListener("mousemove", handleMouseMove);
    document.removeEventListener("mouseup", stopResizing);
  }, [handleMouseMove]);

  const calculateVolumeRow = useCallback((row) => {
    if (typeof row.overrideTotal === "number") return row.overrideTotal;
    const rawL = parseFloat(row.largo);
    const rawA = parseFloat(row.ancho);
    const rawH = parseFloat(row.alto);
    const rawK = parseFloat(row.kg_ml);
    const l = isNaN(rawL) || rawL === 0 ? 0 : Math.abs(rawL);
    const a = isNaN(rawA) || rawA === 0 ? 0 : Math.abs(rawA);
    const h = isNaN(rawH) || rawH === 0 ? 0 : Math.abs(rawH);
    const k = isNaN(rawK) || rawK === 0 ? 0 : Math.abs(rawK);
    const p = parseFloat(row.piezas) || 1;
    if (l === 0 && a === 0 && h === 0 && k === 0) return 0;

    let volPza = 0;
    if (row.isPasos) {
      volPza = Math.floor(l / 4);
    } else {
      volPza =
        (l > 0 ? l : 1) * (a > 0 ? a : 1) * (h > 0 ? h : 1) * (k > 0 ? k : 1);
    }

    return Math.round(volPza * p * 100) / 100;
  }, []);
  const getVolumenNivel = (idC, idN) =>
    generadores[idC]?.[idN]?.rows?.reduce(
      (acc, item) => acc + calculateVolumeRow(item),
      0,
    ) || 0;
  const getTotalConcepto = (idC) =>
    niveles.reduce((acc, n) => acc + getVolumenNivel(idC, n.id), 0);

  const updateConceptoField = useCallback(
    (id, field, value) =>
      updateActiveConceptos((prev) =>
        prev.map((c) => (c.id === id ? { ...c, [field]: value } : c)),
      ),
    [updateActiveConceptos],
  );

  const updateConceptoId = useCallback(
    (oldId, newId) => {
      if (oldId === newId || !newId.trim()) return;
      updateActiveConceptos((prev) => {
        if (prev.some((c) => c.id === newId)) return prev;
        const matched = catalogoConceptos.find((cat) => (cat.codigo || cat.id) === newId);
        return prev.map((c) =>
          c.id === oldId
            ? {
                ...c,
                id: newId,
                ...(matched
                  ? {
                      clave: matched.clave || "",
                      cc: matched.cc || "",
                      justificacion: matched.justificacion || "",
                      descripcion: matched.descripcion || "",
                      unidad: matched.unidad || "m2",
                    }
                  : {}),
              }
            : c,
        );
      });
      updateActiveGeneradores((prev) => {
        if (!prev[oldId]) return prev;
        const newState = { ...prev };
        newState[newId] = newState[oldId];
        delete newState[oldId];
        return newState;
      });
      setSelectedIds((prev) => prev.map((id) => (id === oldId ? newId : id)));
    },
    [updateActiveConceptos, updateActiveGeneradores, catalogoConceptos],
  );

  const handlePasteProjectConceptos = useCallback(
    (e) => {
      if (editingModal) return; // Don't paste if we are editing a generador or details
      const clipboardData = e.clipboardData;
      if (!clipboardData) return;
      
      const text = clipboardData.getData("text/plain");
      if (!text) return;
      
      const lines = text.split(/\r?\n/).filter((line) => line.trim() !== "");
      if (lines.length > 0 && text.includes("\t")) {
        e.preventDefault();
        const newRows = lines.map((line) => {
          const cols = line.split("\t");
          if (cols.length >= 6) {
             return {
               id: cols[0]?.trim() || `C-${Date.now().toString(36)}-${Math.floor(Math.random()*1000)}`,
               clave: cols[1]?.trim() || "",
               cc: cols[2]?.trim() || "",
               justificacion: cols[3]?.trim() || "",
               descripcion: cols[4]?.trim() || "",
               unidad: cols[5]?.trim() || "m2",
             };
          } else {
             return {
               id: cols[0]?.trim() || `C-${Date.now().toString(36)}-${Math.floor(Math.random()*1000)}`,
               clave: "",
               cc: "",
               justificacion: "",
               descripcion: cols[1]?.trim() || "",
               unidad: cols[2]?.trim() || "m2",
             };
          }
        });

        updateActiveConceptos((prev) => {
          const newState = [...prev];
          newRows.forEach((r) => {
            if (!newState.some((x) => x.id === r.id)) {
              newState.push(r);
            }
          });
          return newState;
        });
      }
    },
    [updateActiveConceptos, editingModal]
  );

  const updateRowGenerador = useCallback(
    (rowId, field, value) => {
      if (!editingModal) return;
      const { concepto, nivel } = editingModal;
      updateActiveGeneradores((prev) => {
        const currentConcept = prev[concepto.id] || {},
          currentLevel = currentConcept[nivel.id] || { rows: [] };
        return {
          ...prev,
          [concepto.id]: {
            ...currentConcept,
            [nivel.id]: {
              ...currentLevel,
              rows: currentLevel.rows.map((r) =>
                r.id === rowId ? { ...r, [field]: value } : r,
              ),
            },
          },
        };
      });
    },
    [editingModal, updateActiveGeneradores],
  );
  const deleteRowGenerador = useCallback(
    (rowId) => {
      if (!editingModal) return;
      const { concepto, nivel } = editingModal;
      updateActiveGeneradores((prev) => {
        const currentConcept = prev[concepto.id] || {},
          currentLevel = currentConcept[nivel.id];
        if (!currentLevel) return prev;
        return {
          ...prev,
          [concepto.id]: {
            ...currentConcept,
            [nivel.id]: {
              ...currentLevel,
              rows: currentLevel.rows.filter(
                (r) => r.id !== rowId && !selectedGeneratorRows.includes(r.id),
              ),
            },
          },
        };
      });
      setSelectedGeneratorRows((prev) => prev.filter((id) => id !== rowId));
    },
    [editingModal, updateActiveGeneradores, selectedGeneratorRows],
  );

  const handlePasteImage = (e) => {
    const items = e.clipboardData.items;
    for (let i = 0; i < items.length; i++) {
      if (items[i].type.indexOf("image") !== -1) {
        const file = items[i].getAsFile();
        const reader = new FileReader();
        reader.onload = (event) => {
          if (!editingModal) return;
          const img = new Image();
          img.onload = () => {
            const canvas = document.createElement("canvas");
            let width = img.width;
            let height = img.height;
            const maxDimension = 600;
            if (width > height) {
              if (width > maxDimension) {
                height = Math.round((height * maxDimension) / width);
                width = maxDimension;
              }
            } else {
              if (height > maxDimension) {
                width = Math.round((width * maxDimension) / height);
                height = maxDimension;
              }
            }
            canvas.width = width;
            canvas.height = height;
            const ctx = canvas.getContext("2d");
            ctx.drawImage(img, 0, 0, width, height);
            
            // Compress image substantially
            const compressedBase64 = canvas.toDataURL("image/jpeg", 0.4);

            updateActiveGeneradores((prev) => {
              const currentConcept = prev[editingModal.concepto.id] || {},
                currentLevel = currentConcept[editingModal.nivel.id] || {
                  rows: [],
                  images: [],
                };
              
              let currentImages = currentLevel.images || [];
              if (currentLevel.image && currentImages.length === 0) {
                 currentImages = [currentLevel.image];
              }

              let newImages = [...currentImages];
              if (newImages.length < 3) {
                  newImages.push(compressedBase64);
              }

              return {
                ...prev,
                [editingModal.concepto.id]: {
                  ...currentConcept,
                  [editingModal.nivel.id]: {
                    ...currentLevel,
                    images: newImages,
                    image: null,
                  },
                },
              };
            });
          };
          img.src = event.target.result;
        };
        reader.readAsDataURL(file);
      }
    }
  };

    const copyRowsToExcelClipboard = useCallback((rows) => {
      const tsvString = rows
        .map((r) => {
          const rawL = parseFloat(r.largo);
          const rawA = parseFloat(r.ancho);
          const rawH = parseFloat(r.alto);
          const rawK = parseFloat(r.kg_ml);
          const rawP = parseFloat(r.piezas);
          const L = isNaN(rawL) || rawL === 0 ? 0 : Math.abs(rawL);
          const A = isNaN(rawA) || rawA === 0 ? 0 : Math.abs(rawA);
          const H = isNaN(rawH) || rawH === 0 ? 0 : Math.abs(rawH);
          const K = isNaN(rawK) || rawK === 0 ? 0 : Math.abs(rawK);
          const pzas = isNaN(rawP) || rawP === 0 ? 1 : rawP;

          const volPza =
            (L > 0 ? L : 1) * (A > 0 ? A : 1) * (H > 0 ? H : 1) * (K > 0 ? K : 1);
          const isActuallyZero = L === 0 && A === 0 && H === 0 && K === 0;
          let finalVolPza = isActuallyZero ? 0 : volPza;
          if (r.isPasos) finalVolPza = Math.floor(L / 4);
          else if (typeof r.overrideVolPza === "number")
            finalVolPza = r.overrideVolPza;
          const volTotal =
            typeof r.overrideTotal === "number"
              ? r.overrideTotal
              : finalVolPza * pzas;

          const formatDim = (val) => {
            if (isNaN(parseFloat(val)) || parseFloat(val) === 0) return "";
            return parseFloat(val).toFixed(2);
          };
          const formatKgMl = (val) => {
            if (isNaN(parseFloat(val)) || parseFloat(val) === 0) return "";
            return parseFloat(val).toFixed(3);
          };
          const formatPzas = (val) => {
            if (isNaN(parseFloat(val)) || parseFloat(val) === 0) return "1.00";
            return parseFloat(val).toFixed(2);
          };
          const claveCompleta = [r.eje, r.claveLoc].filter(Boolean).join(" - ");
          return `${claveCompleta}\t${formatDim(r.largo)}\t${formatDim(r.ancho)}\t${formatKgMl(r.kg_ml)}\t${formatDim(r.alto)}\t${finalVolPza.toFixed(2)}\t${formatPzas(r.piezas)}\t${volTotal.toFixed(2)}`;
        })
        .join("\n");
      if (navigator.clipboard) {
        navigator.clipboard.writeText(tsvString).catch(() => {});
      } else {
        const textArea = document.createElement("textarea");
        textArea.value = tsvString;
        textArea.style.position = "fixed";
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        try {
          // eslint-disable-next-line
          document.execCommand("copy");
        } catch (err) {}
        document.body.removeChild(textArea);
      }
    }, []);

  const handleCopyRows = () => {
    if (!editingModal) return;
    const allRows =
      generadores[editingModal.concepto.id]?.[editingModal.nivel.id]?.rows ||
      [];
    const rowsToCopy =
      selectedGeneratorRows.length > 0
        ? allRows.filter((r) => selectedGeneratorRows.includes(r.id))
        : allRows;
    if (rowsToCopy.length === 0) return;
    setClipboardRows(rowsToCopy);
    copyRowsToExcelClipboard(rowsToCopy);
    setCopyStatus("copied");
    setTimeout(() => setCopyStatus(""), 2000);
  };
  const handlePasteRows = () => {
    if (!editingModal || !clipboardRows || clipboardRows.length === 0) return;
    const { concepto, nivel } = editingModal;
    updateActiveGeneradores((prev) => {
      const currentConcept = prev[concepto.id] || {};
      const currentLevel = currentConcept[nivel.id] || { rows: [] };
      let updatedRows = [...currentLevel.rows];
      if (selectedGeneratorRows.length > 0) {
        const sortedSelectedIds = updatedRows
          .filter((r) => selectedGeneratorRows.includes(r.id))
          .map((r) => r.id);
        updatedRows = updatedRows.map((row) => {
          if (selectedGeneratorRows.includes(row.id)) {
            const clipData =
              clipboardRows.length === 1
                ? clipboardRows[0]
                : clipboardRows[
                    sortedSelectedIds.indexOf(row.id) % clipboardRows.length
                  ];
            return {
              ...row,
              eje: clipData.eje !== undefined ? clipData.eje : row.eje,
              claveLoc:
                clipData.claveLoc !== undefined
                  ? clipData.claveLoc
                  : row.claveLoc,
              largo: clipData.largo !== undefined ? clipData.largo : row.largo,
              ancho: clipData.ancho !== undefined ? clipData.ancho : row.ancho,
              alto: clipData.alto !== undefined ? clipData.alto : row.alto,
              kg_ml: clipData.kg_ml !== undefined ? clipData.kg_ml : row.kg_ml,
              piezas:
                clipData.piezas !== undefined ? clipData.piezas : row.piezas,
              overrideVolPza:
                clipData.overrideVolPza !== undefined
                  ? clipData.overrideVolPza
                  : row.overrideVolPza,
              overrideTotal:
                clipData.overrideTotal !== undefined
                  ? clipData.overrideTotal
                  : row.overrideTotal,
              isPasos:
                clipData.isPasos !== undefined ? clipData.isPasos : row.isPasos,
            };
          }
          return row;
        });
      } else {
        const newRows = clipboardRows.map((row, i) => ({
          ...row,
          id: `R-${Date.now()}-${i}-${Math.random().toString(36).substr(2, 5)}`,
        }));
        updatedRows = [...updatedRows, ...newRows];
      }
      return {
        ...prev,
        [editingModal.concepto.id]: {
          ...currentConcept,
          [editingModal.nivel.id]: { ...currentLevel, rows: updatedRows },
        },
      };
    });
  };

  const toggleAllGeneratorRows = () => {
    if (!editingModal) return;
    const allRows =
      generadores[editingModal.concepto.id]?.[editingModal.nivel.id]?.rows ||
      [];
    if (selectedGeneratorRows.length === allRows.length && allRows.length > 0)
      setSelectedGeneratorRows([]);
    else setSelectedGeneratorRows(allRows.map((r) => r.id));
  };
  const toggleStructureRow = (rowId) =>
    setSelectedStructureRows((prev) =>
      prev.includes(rowId)
        ? prev.filter((id) => id !== rowId)
        : [...prev, rowId],
    );
  const toggleSelectConcept = (id) => {
    setSelectedIds((prev) =>
      prev.includes(id) ? prev.filter((i) => i !== id) : [...prev, id]
    );
  };
  const toggleAllStructureRows = () => {
    if (
      selectedStructureRows.length === estructuras.length &&
      estructuras.length > 0
    )
      setSelectedStructureRows([]);
    else setSelectedStructureRows(estructuras.map((e) => e.id));
  };
  const toggleWallRow = (rowId) =>
    setSelectedWallRows((prev) =>
      prev.includes(rowId)
        ? prev.filter((id) => id !== rowId)
        : [...prev, rowId],
    );
  const toggleAllWallRows = () => {
    if (selectedWallRows.length === muros.length && muros.length > 0)
      setSelectedWallRows([]);
    else setSelectedWallRows(muros.map((m) => m.id));
  };

  const handleCopySelectedWalls = () => {
    const rowsToCopy =
      selectedWallRows.length > 0
        ? muros.filter((r) => selectedWallRows.includes(r.id))
        : muros;
    if (rowsToCopy.length === 0) return;
    setCopiedWallRowsList(rowsToCopy);
    setCopyStatus("copied-wall-list");
    setTimeout(() => setCopyStatus(""), 2000);
  };
  const handlePasteSelectedWalls = () => {
    if (!copiedWallRowsList || copiedWallRowsList.length === 0) return;
    updateActiveMuros((prev) => {
      let updatedRows = [...prev];
      if (selectedWallRows.length > 0) {
        const sortedSelectedIds = updatedRows
          .filter((r) => selectedWallRows.includes(r.id))
          .map((r) => r.id);
        updatedRows = updatedRows.map((row) => {
          if (selectedWallRows.includes(row.id)) {
            const clipData =
              copiedWallRowsList.length === 1
                ? copiedWallRowsList[0]
                : copiedWallRowsList[
                    sortedSelectedIds.indexOf(row.id) %
                      copiedWallRowsList.length
                  ];
            const clonedHuecos = (clipData.huecos || []).map((h) => ({
              ...h,
              id: `H-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
            }));
            const clonedCastillos = (clipData.castillos || []).map((c) => ({
              ...c,
              id: `C-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
            }));
            return {
              ...row,
              eje: clipData.eje !== undefined ? clipData.eje : row.eje,
              clave: clipData.clave !== undefined ? clipData.clave : row.clave,
              largo: clipData.largo !== undefined ? clipData.largo : row.largo,
              ancho: clipData.ancho !== undefined ? clipData.ancho : row.ancho,
              alto: clipData.alto !== undefined ? clipData.alto : row.alto,
              tipoAplanadoC1:
                clipData.tipoAplanadoC1 !== undefined
                  ? clipData.tipoAplanadoC1
                  : row.tipoAplanadoC1,
              tipoAplanadoC2:
                clipData.tipoAplanadoC2 !== undefined
                  ? clipData.tipoAplanadoC2
                  : row.tipoAplanadoC2,
              tipoRecubrimientoC1:
                clipData.tipoRecubrimientoC1 !== undefined
                  ? clipData.tipoRecubrimientoC1
                  : row.tipoRecubrimientoC1,
              tipoRecubrimientoC2:
                clipData.tipoRecubrimientoC2 !== undefined
                  ? clipData.tipoRecubrimientoC2
                  : row.tipoRecubrimientoC2,
              huecos: clonedHuecos,
              castillos: clonedCastillos,
            };
          }
          return row;
        });
      } else {
        const newRows = copiedWallRowsList.map((row, i) => {
          const clonedHuecos = (row.huecos || []).map((h) => ({
            ...h,
            id: `H-${Date.now()}-${i}-${Math.random().toString(36).substr(2, 5)}`,
          }));
          const clonedCastillos = (row.castillos || []).map((c) => ({
            ...c,
            id: `C-${Date.now()}-${i}-${Math.random().toString(36).substr(2, 5)}`,
          }));
          return {
            ...row,
            id: `M-${Date.now()}-${i}-${Math.random().toString(36).substr(2, 5)}`,
            huecos: clonedHuecos,
            castillos: clonedCastillos,
          };
        });
        updatedRows = [...updatedRows, ...newRows];
      }
      return updatedRows;
    });
    setCopyStatus("pasted-wall-list");
    setTimeout(() => setCopyStatus(""), 2000);
  };

  const handleCopyWallRow = useCallback((row) => {
    setCopiedWallRow(row);
    setCopyStatus(`copied-wall-${row.id}`);
    setTimeout(() => setCopyStatus(""), 2000);
  }, []);
  const handlePasteWallRow = useCallback(
    (targetId) => {
      if (!copiedWallRow) return;
      updateActiveMuros((prev) =>
        prev.map((m) => {
          if (m.id !== targetId) return m;
          const clonedHuecos = (copiedWallRow.huecos || []).map((h) => ({
            ...h,
            id: `H-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
          }));
          const clonedCastillos = (copiedWallRow.castillos || []).map((c) => ({
            ...c,
            id: `C-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
          }));
          return {
            ...m,
            largo: copiedWallRow.largo || "",
            ancho: copiedWallRow.ancho || "",
            alto: copiedWallRow.alto || "",
            tipoAplanadoC1: copiedWallRow.tipoAplanadoC1 || "",
            tipoAplanadoC2: copiedWallRow.tipoAplanadoC2 || "",
            tipoRecubrimientoC1: copiedWallRow.tipoRecubrimientoC1 || "",
            tipoRecubrimientoC2: copiedWallRow.tipoRecubrimientoC2 || "",
            huecos: clonedHuecos,
            castillos: clonedCastillos,
          };
        }),
      );
      setCopyStatus(`pasted-wall-${targetId}`);
      setTimeout(() => setCopyStatus(""), 2000);
    },
    [copiedWallRow, updateActiveMuros],
  );

  const handleCopySelectedStructures = () => {
    const rowsToCopy =
      selectedStructureRows.length > 0
        ? estructuras.filter((r) => selectedStructureRows.includes(r.id))
        : estructuras;
    if (rowsToCopy.length === 0) return;
    setCopiedStructureRowsList(rowsToCopy);
    setCopyStatus("copied-structure-list");
    setTimeout(() => setCopyStatus(""), 2000);
  };
  const handlePasteSelectedStructures = () => {
    if (!copiedStructureRowsList || copiedStructureRowsList.length === 0)
      return;
    updateActiveEstructuras((prev) => {
      let updatedRows = [...prev];
      if (selectedStructureRows.length > 0) {
        const sortedSelectedIds = updatedRows
          .filter((r) => selectedStructureRows.includes(r.id))
          .map((r) => r.id);
        updatedRows = updatedRows.map((row) => {
          if (selectedStructureRows.includes(row.id)) {
            const clipData =
              copiedStructureRowsList.length === 1
                ? copiedStructureRowsList[0]
                : copiedStructureRowsList[
                    sortedSelectedIds.indexOf(row.id) %
                      copiedStructureRowsList.length
                  ];
            const clonedAceros = (clipData.aceros || []).map((a) => ({
              ...a,
              id: `AC-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
            }));
            const clonedCasetones = (clipData.casetones || []).map((c) => ({
              ...c,
              id: `CAS-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
            }));
            return {
              ...row,
              eje: clipData.eje !== undefined ? clipData.eje : row.eje,
              clave: clipData.clave !== undefined ? clipData.clave : row.clave,
              tipo: clipData.tipo !== undefined ? clipData.tipo : row.tipo,
              i: clipData.i !== undefined ? clipData.i : row.i,
              largo: clipData.largo !== undefined ? clipData.largo : row.largo,
              ancho: clipData.ancho !== undefined ? clipData.ancho : row.ancho,
              alto: clipData.alto !== undefined ? clipData.alto : row.alto,
              piezas:
                clipData.piezas !== undefined ? clipData.piezas : row.piezas,
              aceros: clonedAceros,
              casetones: clonedCasetones,
            };
          }
          return row;
        });
      } else {
        const newRows = copiedStructureRowsList.map((row, i) => {
          const clonedAceros = (row.aceros || []).map((a) => ({
            ...a,
            id: `AC-${Date.now()}-${i}-${Math.random().toString(36).substr(2, 5)}`,
          }));
          const clonedCasetones = (row.casetones || []).map((c) => ({
            ...c,
            id: `CAS-${Date.now()}-${i}-${Math.random().toString(36).substr(2, 5)}`,
          }));
          return {
            ...row,
            id: `E-${Date.now()}-${i}-${Math.random().toString(36).substr(2, 5)}`,
            aceros: clonedAceros,
            casetones: clonedCasetones,
          };
        });
        updatedRows = [...updatedRows, ...newRows];
      }
      return updatedRows;
    });
    setCopyStatus("pasted-structure-list");
    setTimeout(() => setCopyStatus(""), 2000);
  };

  const getWallHuecosTotal = (huecos = []) => {
    const sum = huecos.reduce(
      (acc, h) =>
        acc +
        (parseFloat(h.largo) || 0) *
          (parseFloat(h.alto) || 0) *
          (parseFloat(h.piezas) || 1),
      0,
    );
    return Math.round(sum * 100) / 100;
  };
  const getWallCastillosTotal = (castillos = [], wallAlto) => {
    const sum = castillos.reduce(
      (acc, c) =>
        acc +
        (parseFloat(c.ancho) || 0) *
          (parseFloat(wallAlto) || 0) *
          (parseFloat(c.piezas) || 1),
      0,
    );
    return Math.round(sum * 100) / 100;
  };
  const calcWallArea = (largo, alto) => {
    const area = (parseFloat(largo) || 0) * (parseFloat(alto) || 0);
    return Math.round(area * 100) / 100;
  };
  const calcWallNetArea = (row) =>
    Math.max(
      0,
      calcWallArea(row.largo, row.alto) -
        getWallHuecosTotal(row.huecos) -
        getWallCastillosTotal(row.castillos, row.alto),
    );

  const getCasetonesTotalVol = (casetones = [], structPiezas = 1) => {
    const sum = casetones.reduce(
      (acc, c) =>
        acc +
        (parseFloat(c.largo) || 0) *
          (parseFloat(c.ancho) || 0) *
          (parseFloat(c.alto) || 0) *
          (parseFloat(c.piezas) || 1),
      0,
    );
    return Math.round(sum * structPiezas * 100) / 100;
  };

  const wallSummary = useMemo(() => {
    let totalMuros = 0;
    const aplanados = {};
    const recubrimientos = {};
    const murosPorAncho = {};
    muros.forEach((w) => {
      const bruto = calcWallArea(w.largo, w.alto);
      const dedH = getWallHuecosTotal(w.huecos);
      const dedC = getWallCastillosTotal(w.castillos, w.alto);
      const neto = Math.max(0, bruto - dedH - dedC);
      const areaAcabados = Math.max(0, bruto - dedH);
      totalMuros += neto;
      const anchoStr = (!isNaN(parseFloat(w.ancho)) ? parseFloat(w.ancho) : 0).toFixed(2);
      murosPorAncho[anchoStr] = (murosPorAncho[anchoStr] || 0) + neto;
      const addAreaAcabado = (dict, tipo) => {
        if (tipo?.trim()) {
          const key = tipo.trim().toUpperCase();
          dict[key] = (dict[key] || 0) + areaAcabados;
        }
      };
      addAreaAcabado(aplanados, w.tipoAplanadoC1);
      addAreaAcabado(aplanados, w.tipoAplanadoC2);
      addAreaAcabado(recubrimientos, w.tipoRecubrimientoC1);
      addAreaAcabado(recubrimientos, w.tipoRecubrimientoC2);
    });
    return { totalMuros, murosPorAncho, aplanados, recubrimientos };
  }, [muros]);

  const calcConcreto = (r) => {
    const tipo = r.tipo?.toLowerCase() || "";
    if (
      tipo === "nervadura" ||
      tipo.includes("vigueta h=") ||
      tipo.includes("vigueta h35") ||
      tipo.includes("vigueta h25")
    )
      return 0;

    const hOrig = parseFloat(r.alto) || 0;
    const p = parseFloat(r.piezas) || 1;
    const l = getEffectiveLargo(r);
    const a = parseFloat(r.ancho) || 0;

    const thickness = parseFloat(r.structEspesorLosa) || 0;
    const thicknessZapata = parseFloat(r.structEspesorZapata) || 0;
    const slabTypeInput = (r.structTipoLosa || "").toLowerCase();
    const isCorner = r.structEsquinaLosa === true;

    // Start with height minus footing thickness
    let effectiveH = Math.max(0, hOrig - thicknessZapata);

    if (slabTypeInput) {
      if (slabTypeInput !== "vigueta") {
        if (isCorner && (tipo.includes("muro") || tipo.includes("columna") || tipo === "dado" || tipo === "trabe" || tipo === "contratrabe")) {
          effectiveH -= thickness / 2;
        } else {
          effectiveH -= thickness;
        }
      }
    } else {
      const simpleDiscount = parseFloat(r.descuentoLosa) || 0;
      effectiveH = Math.max(0, effectiveH - simpleDiscount);
    }
    
    effectiveH = Math.max(0, effectiveH);

    if (tipo === "columna circular" || tipo === "pilas" || tipo === "pila") {
      const radio = l / 2;
      return Math.round(Math.PI * radio * radio * effectiveH * p * 100) / 100;
    }

    const volBruto = l * a * effectiveH * p;
    if (tipo === "losa nervada") {
      const volCasetones = getCasetonesTotalVol(r.casetones || [], r.piezas);
      return Math.max(0, Math.round((volBruto - volCasetones) * 100) / 100);
    }
    return Math.round(volBruto * 100) / 100;
  };
  const calcCimbra = (r) => {
    const l = getEffectiveLargo(r),
      a = parseFloat(r.ancho) || 0,
      hOrig = parseFloat(r.alto) || 0,
      p = parseFloat(r.piezas) || 1,
      t = (r.tipo || "").toLowerCase();

    const thicknessLosa = parseFloat(r.structEspesorLosa) || 0;
    const thicknessZapata = parseFloat(r.structEspesorZapata) || 0;
    const isEsquina = r.structEsquinaLosa === true;
    const hasTipoLosa = !!r.structTipoLosa;

    if (
      t.includes("vigueta h=") ||
      t.includes("vigueta h35") ||
      t.includes("vigueta h25")
    )
      return 0;

    // Base height for cimbra after zapata discount
    const hBase = Math.max(0, hOrig - thicknessZapata);

    let area = 0;
    if (t === "columna" || t === "dado") {
      let perimeter_eff;
      if (hasTipoLosa) {
        if (isEsquina) {
          // (l+a) at hBase + (l+a) at (hBase - thicknessLosa)
          perimeter_eff = (l + a) * hBase + (l + a) * (hBase - thicknessLosa);
        } else {
          perimeter_eff = (l + a) * 2 * (hBase - thicknessLosa);
        }
      } else {
        const simpleDiscount = parseFloat(r.descuentoLosa) || 0;
        perimeter_eff = (l + a) * 2 * (hBase - simpleDiscount);
      }
      area = Math.max(0, perimeter_eff) * p;
    } else if (t === "columna circular") {
      let h_eff = hBase;
      if (hasTipoLosa) {
        if (isEsquina) h_eff = hBase - thicknessLosa / 2;
        else h_eff = hBase - thicknessLosa;
      } else {
        const simpleDiscount = parseFloat(r.descuentoLosa) || 0;
        h_eff = Math.max(0, hBase - simpleDiscount);
      }
      area = Math.PI * l * h_eff * p;
    } else if (t === "trabe" || t === "contratrabe") {
      let sideSum;
      if (hasTipoLosa) {
        if (isEsquina) sideSum = hBase + (hBase - thicknessLosa);
        else sideSum = (hBase - thicknessLosa) * 2;
      } else {
        const simpleDiscount = parseFloat(r.descuentoLosa) || 0;
        sideSum = (hBase - simpleDiscount) * 2;
      }
      area = (l * a + l * sideSum) * p;
    } else if (
      t === "losa" ||
      t === "losa nervada" ||
      t === "losa de vigueta" ||
      t === "escalera papelillo" ||
      t === "rampa de escalera"
    )
      area = l * a * p;
    else if (t === "muro" || t === "muro curvo") {
      let areaMuro;
      if (hasTipoLosa) {
        if (isEsquina) {
          // One face full hBase, other hBase - thicknessLosa
          areaMuro = l * hBase + l * (hBase - thicknessLosa);
        } else {
          areaMuro = l * (hBase - thicknessLosa) * 2;
        }
      } else {
        const simpleDiscount = parseFloat(r.descuentoLosa) || 0;
        const h_eff = Math.max(0, hBase - simpleDiscount);
        areaMuro = l * h_eff * 2;
      }
      area = Math.max(0, areaMuro) * p;
    } else if (t === "zapata aislada" || t === "zapata corrida") {
      area = (l + a) * 2 * hBase * p;
    } else {
      area = (l + a) * 2 * hBase * p;
    }
    return Math.round(area * 100) / 100;
  };
  const calcCimbraFrontera = (r) => {
    const l = getEffectiveLargo(r),
      a = parseFloat(r.ancho) || 0,
      p = parseFloat(r.piezas) || 1,
      t = (r.tipo || "").toLowerCase();
    let front = 0;
    if (
      t === "losa" ||
      t === "losa nervada" ||
      t.includes("losa vigueta") ||
      t.includes("losa de vigueta") ||
      t.includes("losa doble vigueta") ||
      t === "escalera papelillo" ||
      t === "rampa de escalera"
    )
      front = (l + a) * 2 * p;
    return Math.round(front * 100) / 100;
  };
  const calcAceroItem = (acero) => {
    const p = parseFloat(acero.piezas) || 0,
      l = parseFloat(acero.longitud) || 0,
      g = parseFloat(acero.ganchos) || 0,
      anc = parseFloat(acero.anclaje) || 0,
      t = parseFloat(acero.traslapes) || 0;
    const tipo = (acero.tipo || "").toLowerCase();
    const isEstriboGroup =
      tipo === "estribos" || tipo === "grapas" || tipo === "zuncho";
    let mlPorPieza;
    if (isEstriboGroup) {
      mlPorPieza = l;
    } else {
      const multi = ACERO_MULT[acero.numVarilla] || 0.5;
      mlPorPieza = l + g * multi + anc + t;
    }
    const mlTotal = mlPorPieza * p;
    const pesoKgM = PESOS_VARILLA[acero.numVarilla] || 0;
    return { mlPorPieza, ml: mlTotal, kg: mlTotal * pesoKgM };
  };
  const calcAceroTotalKg = (aceros = [], piezasElem = 1) => {
    const pz = parseFloat(piezasElem) || 1;
    return aceros.reduce(
      (acc, a) => acc + Math.round(calcAceroItem(a).kg * pz * 100) / 100,
      0,
    );
  };

  const estructuraSummary = useMemo(() => {
    let totalConcreto = 0,
      totalConcretoCimentacion = 0,
      totalConcretoEstructura = 0,
      totalConcretoColumnas = 0,
      totalConcretoMuros = 0,
      totalCimbra = 0,
      totalCimbraCimentacion = 0,
      totalCimbraFronteraCimentacion = 0,
      totalAceroGeneral = 0,
      totalAceroCimentacion = 0,
      totalAceroEstructuraGlobal = 0,
      totalPlantilla = 0,
      totalExcavacion = 0,
      totalAfine = 0,
      totalRelleno = 0,
      totalCasetonesGeneral = 0,
      totalViguetaH25 = 0,
      totalViguetaH35 = 0,
      totalDobleViguetaH25 = 0,
      totalDobleViguetaH35 = 0,
      totalAreaViguetasGlobal = 0;
    let totalCimbraFrontera = 0;
    let totalCimbraFronteraViguetas = 0;
    const breakdown = {},
      aceroCimentacionDetalle = {},
      aceroEstructuraDetalle = {},
      aceroEstructuraGlobalDetalle = {};
    estructuras.forEach((e) => {
      let acerosCorregidos = e.aceros || [];
      const validTypesOptions = getSteelTypesForElement(e.tipo);
      const validTypes = validTypesOptions.map((t) => t.toLowerCase());
      if (validTypes.length > 0) {
        acerosCorregidos = acerosCorregidos.map((a) => {
          if (!validTypes.includes((a.tipo || "").toLowerCase())) {
            return { ...a, tipo: validTypesOptions[0] };
          }
          return a;
        });
      }

      const c = calcConcreto(e),
        cim = calcCimbra(e),
        cimFrontera = calcCimbraFrontera(e),
        aceroKg = calcAceroTotalKg(acerosCorregidos, e.piezas);
      const l = getEffectiveLargo(e);
      const a = parseFloat(e.ancho) || 0;
      const piezasMultiplier = parseFloat(e.piezas || 1) || 1;
      const tipo = (e.tipo || "Sin Seleccionar").toLowerCase().trim();
      const isCimentacion = [
        "zapata aislada",
        "zapata corrida",
        "contratrabe",
        "dado",
        "pilas",
        "pila",
      ].includes(tipo);
      const isEstructuraGlobal = [
        "losa",
        "losa nervada",
        "losa de vigueta",
        "losa vigueta h=35",
        "losa vigueta h=25",
        "losa doble vigueta h=25",
        "losa doble vigueta h=35",
        "trabe",
        "nervadura",
        "columna",
        "columna circular",
        "escalera papelillo",
        "rampa de escalera",
      ].includes(tipo);
      const isEstructuraGpo = [
        "losa",
        "losa nervada",
        "losa de vigueta",
        "losa vigueta h=35",
        "losa vigueta h=25",
        "losa doble vigueta h=25",
        "losa doble vigueta h=35",
        "trabe",
        "nervadura",
        "escalera papelillo",
        "rampa de escalera",
      ].includes(tipo);

      if (isCimentacion) {
        totalConcretoCimentacion += c;
        totalAceroCimentacion += aceroKg;
        totalCimbraCimentacion += cim;
        totalCimbraFronteraCimentacion += cimFrontera;
      } else {
        totalConcreto += c;
        totalAceroGeneral += aceroKg;
        if (isEstructuraGpo) totalConcretoEstructura += c;
        else if (
          tipo === "columna" ||
          tipo === "columna circular" ||
          tipo === "columnas" ||
          tipo === "columnas circulares"
        )
          totalConcretoColumnas += c;
        else if (
          tipo === "muro" ||
          tipo === "muro curvo" ||
          tipo === "muros" ||
          tipo === "muros curvos"
        )
          totalConcretoMuros += c;
      }

      totalCimbra += cim;
      totalCimbraFrontera += cimFrontera;
      if (tipo.includes("vigueta")) {
        totalCimbraFronteraViguetas += cimFrontera;
        totalAreaViguetasGlobal += l * a * piezasMultiplier;
      }

      if (tipo === "losa vigueta h=25") totalViguetaH25 += l * a * piezasMultiplier;
      else if (tipo === "losa vigueta h=35") totalViguetaH35 += l * a * piezasMultiplier;
      else if (tipo === "losa doble vigueta h=25")
        totalDobleViguetaH25 += l * a * piezasMultiplier;
      else if (tipo === "losa doble vigueta h=35")
        totalDobleViguetaH35 += l * a * piezasMultiplier;

      const isConcEstGroup = [
        "losa",
        "losa nervada",
        "losa de vigueta",
        "trabe",
        "escalera papelillo",
        "rampa de escalera",
      ].includes(tipo);

      let groupOtros = tipo;
      if (tipo === "columna" || tipo === "columnas") groupOtros = "Columnas";
      else if (tipo === "columna circular" || tipo === "columnas circulares")
        groupOtros = "Columnas Circulares";
      else if (tipo === "pilas" || tipo === "pila") groupOtros = "Pilas";
      else if (
        tipo === "muro" ||
        tipo === "muro curvo" ||
        tipo === "muros" ||
        tipo === "muros curvos"
      )
        groupOtros = "Muros";
      else if (tipo === "losa" || tipo === "losas" || tipo === "losa nervada") groupOtros = "Losas";
      else if (tipo.includes("vigueta")) {
        groupOtros = "Losas de Vigueta";
      }
      else if (tipo === "escalera papelillo" || tipo === "rampa de escalera")
        groupOtros = "Losas";
      else if (tipo === "trabe") groupOtros = "Trabes";
      else if (tipo === "nervadura") groupOtros = "Nervaduras";
      else if (tipo === "zapata aislada")
        groupOtros = `Zapatas Aisladas (Esp: ${(!isNaN(parseFloat(e.alto)) ? parseFloat(e.alto) : 0).toFixed(2)}m)`;
      else if (tipo === "zapata corrida")
        groupOtros = `Zapatas Corridas (Esp: ${(!isNaN(parseFloat(e.alto)) ? parseFloat(e.alto) : 0).toFixed(2)}m)`;
      else if (tipo === "contratrabe") groupOtros = "Contratrabes";
      else if (tipo === "dado") groupOtros = "Dados";

      const ensureBucket = (key, customTipo) => {
        if (!breakdown[key]) {
          breakdown[key] = {
            tipoOriginal: customTipo || tipo,
            hasLosaNervada: tipo === "losa nervada" || customTipo === "losa nervada",
            isCimentacion: isCimentacion,
            concreto: 0,
            cimbra: 0,
            cimbraFrontera: 0,
            aceroKg: 0,
            aceroDetalle: {},
            plantilla: 0,
            excavacion: 0,
            afine: 0,
            relleno: 0,
            casetones: 0,
            areaViguetasH25: 0,
            areaViguetasH35: 0,
            areaDobleViguetaH25: 0,
            areaDobleViguetaH35: 0,
            areaViguetasGENERIC: 0,
            cimbraEscaleraPapelillo: 0,
            cimbraRampaEscalera: 0,
            obturacionMuros: 0,
            mallaRefuerzo: 0,
            pasosMuros: 0,
            cimbraMuros_0_3: 0,
            cimbraMuros_3_6: 0,
            cimbraMuros_6_9: 0,
            cimbraMuroCurvo_0_3: 0,
            cimbraMuroCurvo_3_6: 0,
            cimbraMuroCurvo_6_9: 0,
            m2MuroTotal: 0,
            emplayerColumnas: 0,
            andamiajeColumnas: 0,
            cimbraColumnas_0_3: 0,
            cimbraColumnas_3_6: 0,
            cimbraColumnas_6_9: 0,
            andamiajeTrabes: 0,
            anclajesDetalle: {},
          };
        }
      };

      if (isConcEstGroup) {
        ensureBucket("Concreto en Estructuras", "concreto_estructura");
        breakdown["Concreto en Estructuras"].concreto += c;
        ensureBucket(groupOtros, tipo);
      } else if (tipo === "columna circular") {
        ensureBucket("Columnas", "columna");
        breakdown["Columnas"].concreto += c;
        ensureBucket(groupOtros, tipo);
      } else {
        ensureBucket(groupOtros, tipo);
        breakdown[groupOtros].concreto += c;
      }

      if (tipo === "losa nervada") {
        const volCasetones = getCasetonesTotalVol(e.casetones || [], piezasMultiplier);
        breakdown[groupOtros].casetones += volCasetones;
        totalCasetonesGeneral += volCasetones;
        breakdown[groupOtros].hasLosaNervada = true;
      }

      if (tipo === "escalera papelillo") {
        ensureBucket("Losas", "losa");
        breakdown["Losas"].cimbraFrontera += cimFrontera;
        breakdown["Losas"].cimbraEscaleraPapelillo += cim;
      } else if (tipo === "rampa de escalera") {
        ensureBucket("Losas", "losa");
        breakdown["Losas"].cimbraFrontera += cimFrontera;
        breakdown["Losas"].cimbraRampaEscalera += cim;
      } else if (tipo.includes("vigueta")) {
        ensureBucket("Losas de Vigueta", "Losas de Vigueta");
        breakdown["Losas de Vigueta"].cimbraFrontera += cimFrontera;
        breakdown["Losas de Vigueta"].cimbra += cim;
        if (/\b(h=?25)\b/i.test(tipo)) {
          if (tipo.includes("doble")) breakdown["Losas de Vigueta"].areaDobleViguetaH25 += l * a * piezasMultiplier;
          else breakdown["Losas de Vigueta"].areaViguetasH25 += l * a * piezasMultiplier;
        } else if (/\b(h=?35)\b/i.test(tipo)) {
          if (tipo.includes("doble")) breakdown["Losas de Vigueta"].areaDobleViguetaH35 += l * a * piezasMultiplier;
          else breakdown["Losas de Vigueta"].areaViguetasH35 += l * a * piezasMultiplier;
        } else {
          breakdown["Losas de Vigueta"].areaViguetasGENERIC += l * a * piezasMultiplier;
        }
      } else {
        breakdown[groupOtros].cimbra += cim;
        breakdown[groupOtros].cimbraFrontera += cimFrontera;
      }

      const p = parseFloat(e.piezas) || 1;
      const h = parseFloat(e.alto) || 0;
      const tLower = tipo;
      const hasTipoLosa = !!e.structTipoLosa;
      const isEsquina = e.structEsquinaLosa === true;
      const thicknessZapata = parseFloat(e.structEspesorZapata) || 0;
      const thicknessLosa = parseFloat(e.structEspesorLosa) || 0;
      const hOrig = parseFloat(e.alto) || 0;
      let hBase = Math.max(0, hOrig - thicknessZapata);
      let h_eff_cimbra = hBase;
      if (hasTipoLosa) {
        if (isEsquina && (tLower.includes("muro") || tLower.includes("columna") || /^(dados?|trabes?|contratrabes?)$/i.test(tLower))) {
          h_eff_cimbra = hBase - (thicknessLosa / 2);
        } else {
          h_eff_cimbra = hBase - thicknessLosa;
        }
      } else {
        const simpleDiscount = parseFloat(e.descuentoLosa) || 0;
        h_eff_cimbra = Math.max(0, hBase - simpleDiscount);
      }
      h_eff_cimbra = Math.max(0, h_eff_cimbra);

      if (
        tipo === "muro" ||
        tipo === "muro curvo" ||
        tipo === "muros" ||
        tipo === "muros curvos"
      ) {
        breakdown[groupOtros].obturacionMuros += cim;
        breakdown[groupOtros].mallaRefuerzo += cim;
        breakdown[groupOtros].m2MuroTotal += l * h * p;
        breakdown[groupOtros].pasosMuros += Math.floor(l / 4) * p;

        if (tipo === "muro curvo") {
          if (h_eff_cimbra <= 3) {
            breakdown[groupOtros].cimbraMuroCurvo_0_3 += cim;
          } else if (h_eff_cimbra <= 6) {
            breakdown[groupOtros].cimbraMuroCurvo_0_3 += (cim / h_eff_cimbra) * 3;
            breakdown[groupOtros].cimbraMuroCurvo_3_6 += (cim / h_eff_cimbra) * (h_eff_cimbra - 3);
          } else {
            breakdown[groupOtros].cimbraMuroCurvo_0_3 += (cim / h_eff_cimbra) * 3;
            breakdown[groupOtros].cimbraMuroCurvo_3_6 += (cim / h_eff_cimbra) * 3;
            breakdown[groupOtros].cimbraMuroCurvo_6_9 += (cim / h_eff_cimbra) * (h_eff_cimbra - 6);
          }
        } else {
          if (h_eff_cimbra <= 3) {
            breakdown[groupOtros].cimbraMuros_0_3 += cim;
          } else if (h_eff_cimbra <= 6) {
            breakdown[groupOtros].cimbraMuros_0_3 += (cim / h_eff_cimbra) * 3;
            breakdown[groupOtros].cimbraMuros_3_6 += (cim / h_eff_cimbra) * (h_eff_cimbra - 3);
          } else {
            breakdown[groupOtros].cimbraMuros_0_3 += (cim / h_eff_cimbra) * 3;
            breakdown[groupOtros].cimbraMuros_3_6 += (cim / h_eff_cimbra) * 3;
            breakdown[groupOtros].cimbraMuros_6_9 += (cim / h_eff_cimbra) * (h_eff_cimbra - 6);
          }
        }
      } else if (tipo === "columna" || tipo === "columna circular") {
        ensureBucket("Columnas", "columna");
        breakdown["Columnas"].emplayerColumnas += cim;

        if (h_eff_cimbra <= 3) {
          breakdown[groupOtros].cimbraColumnas_0_3 += cim;
        } else if (h_eff_cimbra <= 6) {
          breakdown[groupOtros].cimbraColumnas_0_3 += (cim / h_eff_cimbra) * 3;
          breakdown[groupOtros].cimbraColumnas_3_6 += (cim / h_eff_cimbra) * (h_eff_cimbra - 3);
        } else {
          breakdown[groupOtros].cimbraColumnas_0_3 += (cim / h_eff_cimbra) * 3;
          breakdown[groupOtros].cimbraColumnas_3_6 += (cim / h_eff_cimbra) * 3;
          breakdown[groupOtros].cimbraColumnas_6_9 += (cim / h_eff_cimbra) * (h_eff_cimbra - 6);
        }

        if (h_eff_cimbra > 3) breakdown[groupOtros].andamiajeColumnas += p;
      } else if (tipo === "trabe") {
        breakdown[groupOtros].andamiajeTrabes += l * p;
      }

      if (["zapata aislada", "zapata corrida", "contratrabe"].includes(tipo)) {
        const plantillaItem = (l + 0.5) * (a + 0.5) * p,
          excavacionItem = plantillaItem * 1.5,
          afineItem = plantillaItem,
          rellenoItem = Math.max(0, excavacionItem - c);
        breakdown[groupOtros].plantilla += plantillaItem;
        breakdown[groupOtros].excavacion += excavacionItem;
        breakdown[groupOtros].afine += afineItem;
        breakdown[groupOtros].relleno += rellenoItem;
        totalPlantilla += plantillaItem;
        totalExcavacion += excavacionItem;
        totalAfine += afineItem;
        totalRelleno += rellenoItem;
      }

      if (tipo === "pilas" || tipo === "pila") {
        const diam = getEffectiveLargo(e);
        const diamExc = diam + 1.0;
        const radioExc = diamExc / 2;
        const areaExc = Math.PI * radioExc * radioExc;
        const excavacionItem = areaExc * h * p;
        const rellenoItem = Math.max(0, excavacionItem - c);
        breakdown[groupOtros].excavacion += excavacionItem;
        breakdown[groupOtros].relleno += rellenoItem;
        totalExcavacion += excavacionItem;
        totalRelleno += rellenoItem;
      }

      (acerosCorregidos || []).forEach((a) => {
        const numVar = a.numVarilla || "-";
        const kgRounded =
          Math.round(calcAceroItem(a).kg * (parseFloat(e.piezas) || 1) * 100) /
          100;
        if (isCimentacion) {
          aceroCimentacionDetalle[numVar] =
            (aceroCimentacionDetalle[numVar] || 0) + kgRounded;
        } else if (isEstructuraGlobal) {
          totalAceroEstructuraGlobal += kgRounded;
          aceroEstructuraGlobalDetalle[numVar] =
            (aceroEstructuraGlobalDetalle[numVar] || 0) + kgRounded;
        } else {
          aceroEstructuraDetalle[numVar] =
            (aceroEstructuraDetalle[numVar] || 0) + kgRounded;
        }

        breakdown[groupOtros].aceroKg += kgRounded;
        if (!breakdown[groupOtros].aceroDetalle[numVar])
          breakdown[groupOtros].aceroDetalle[numVar] = 0;
        breakdown[groupOtros].aceroDetalle[numVar] += kgRounded;

        const aTipo = (a.tipo || "").toLowerCase().trim();
        const isAnclajeTarget =
          aTipo.includes("vertical") ||
          aTipo.includes("remetido") ||
          aTipo.includes("longitudinal") ||
          aTipo.includes("principal") ||
          aTipo.includes("bastones") ||
          aTipo === "";
        if (
          (tipo === "muro" ||
            tipo === "muro curvo" ||
            tipo === "muros" ||
            tipo === "muros curvos") &&
          isAnclajeTarget
        ) {
          const nv = a.numVarilla || "-";
          if (!breakdown[groupOtros].anclajesDetalle)
            breakdown[groupOtros].anclajesDetalle = {};
          if (!breakdown[groupOtros].anclajesDetalle[nv])
            breakdown[groupOtros].anclajesDetalle[nv] = 0;
          breakdown[groupOtros].anclajesDetalle[nv] +=
            (parseFloat(a.piezas) || 0) * (parseFloat(e.piezas) || 1);
        }
      });
    });

    Object.values(breakdown).forEach((group) => {
      // Pasos en muros is now calculated element-by-element
    });

    return {
      totalConcreto,
      totalConcretoCimentacion,
      totalConcretoEstructura,
      totalConcretoColumnas,
      totalConcretoMuros,
      totalCimbra,
      totalCimbraCimentacion,
      totalCimbraFronteraCimentacion,
      totalAceroGeneral,
      totalAceroCimentacion,
      totalAceroEstructuraGlobal,
      breakdown,
      aceroCimentacionDetalle,
      aceroEstructuraDetalle,
      aceroEstructuraGlobalDetalle,
      totalCimbraFrontera,
      totalCimbraFronteraViguetas,
      totalViguetaH25,
      totalViguetaH35,
      totalDobleViguetaH25,
      totalDobleViguetaH35,
      totalAreaViguetasGlobal,
      totalPlantilla,
      totalExcavacion,
      totalAfine,
      totalRelleno,
      totalCasetonesGeneral,
    };
  }, [estructuras]);

  const updateWallSubItem = useCallback(
    (wallId, type, itemId, field, value) => {
      updateActiveMuros((prev) =>
        prev.map((m) =>
          m.id === wallId
            ? {
                ...m,
                [type]: (m[type] || []).map((item) =>
                  item.id === itemId ? { ...item, [field]: value } : item,
                ),
              }
            : m,
        ),
      );
    },
    [updateActiveMuros],
  );
  const addWallSubItem = useCallback(
    (wallId, type) => {
      updateActiveMuros((prev) =>
        prev.map((m) =>
          m.id === wallId
            ? {
                ...m,
                [type]: [
                  ...(m[type] || []),
                  type === "huecos"
                    ? {
                        id: `H-${Date.now()}`,
                        tipo: "",
                        largo: "",
                        alto: "",
                        piezas: "1",
                      }
                    : { id: `C-${Date.now()}`, ancho: "", piezas: "1" },
                ],
              }
            : m,
        ),
      );
    },
    [updateActiveMuros],
  );
  const removeWallSubItem = useCallback(
    (wallId, type, itemId) => {
      updateActiveMuros((prev) =>
        prev.map((m) =>
          m.id === wallId
            ? {
                ...m,
                [type]: (m[type] || []).filter((item) => item.id !== itemId),
              }
            : m,
        ),
      );
    },
    [updateActiveMuros],
  );

  const updateCasetonItem = useCallback(
    (structId, itemId, field, value) => {
      updateActiveEstructuras((prev) =>
        prev.map((s) =>
          s.id === structId
            ? {
                ...s,
                casetones: (s.casetones || []).map((c) =>
                  c.id === itemId ? { ...c, [field]: value } : c,
                ),
              }
            : s,
        ),
      );
    },
    [updateActiveEstructuras],
  );
  const addCasetonItem = useCallback(
    (structId) => {
      updateActiveEstructuras((prev) =>
        prev.map((s) =>
          s.id === structId
            ? {
                ...s,
                casetones: [
                  ...(s.casetones || []),
                  {
                    id: `CAS-${Date.now()}`,
                    clave: "",
                    largo: "",
                    ancho: "",
                    alto: "",
                    piezas: "1",
                  },
                ],
              }
            : s,
        ),
      );
    },
    [updateActiveEstructuras],
  );
  const removeCasetonItem = useCallback(
    (structId, itemId) => {
      updateActiveEstructuras((prev) =>
        prev.map((s) =>
          s.id === structId
            ? {
                ...s,
                casetones: (s.casetones || []).filter((c) => c.id !== itemId),
              }
            : s,
        ),
      );
    },
    [updateActiveEstructuras],
  );

  const updateCatalogoField = useCallback(
    (id, field, value) => {
      setCatalogoConceptos((prev) =>
        prev.map((c) => (c.id === id ? { ...c, [field]: value } : c)),
      );
    },
    [setCatalogoConceptos],
  );
  const updateCatalogoId = useCallback(
    (oldId, newCodigo) => {
      if (!newCodigo.trim()) return;
      setCatalogoConceptos((prev) => {
        const matched = prev.find((cat) => (cat.codigo || cat.id) === newCodigo && cat.id !== oldId);
        return prev.map((c) =>
          c.id === oldId
            ? {
                ...c,
                codigo: newCodigo,
                ...(matched
                  ? {
                      clave: matched.clave || "",
                      cc: matched.cc || "",
                      justificacion: matched.justificacion || "",
                      descripcion: matched.descripcion || "",
                      unidad: matched.unidad || "",
                    }
                  : {}),
              }
            : c
        );
      });
    },
    [setCatalogoConceptos],
  );

  const handleCopyWalls = (type, value) => {
    let rowsToCopy = [];
    const normalize = (val) => (val || "").trim().toUpperCase();
    muros.forEach((m) => {
      const l = m.largo,
        h = m.alto,
        base = { eje: m.eje || "", claveLoc: m.clave || "" };
      let multiplier = 0,
        suffix = "";
      if (type === "TOTAL_NETO") {
        const anchoStr = parseFloat(m.ancho || 0).toFixed(2);
        if (value && anchoStr !== value) return;
        multiplier = 1;
        suffix = "(Área Bruta)";
      } else if (type === "APLANADO") {
        multiplier =
          (normalize(m.tipoAplanadoC1) === value ? 1 : 0) +
          (normalize(m.tipoAplanadoC2) === value ? 1 : 0);
        suffix = `(Área Bruta - ${multiplier} Caras)`;
      } else if (type === "RECUBRIMIENTO") {
        multiplier =
          (normalize(m.tipoRecubrimientoC1) === value ? 1 : 0) +
          (normalize(m.tipoRecubrimientoC2) === value ? 1 : 0);
        suffix = `(Área Bruta - ${multiplier} Caras)`;
      }
      if (multiplier === 0) return;

      rowsToCopy.push({
        ...base,
        claveLoc: (base.claveLoc ? base.claveLoc + " " : "") + suffix,
        largo: l,
        ancho: "",
        alto: h,
        piezas: multiplier,
      });
      (m.huecos || []).forEach((dh) =>
        rowsToCopy.push({
          ...base,
          claveLoc:
            (base.claveLoc ? base.claveLoc + " " : "") +
            `(- ${dh.tipo || "Hueco"})`,
          largo: dh.largo,
          ancho: "",
          alto: dh.alto,
          piezas: -(parseFloat(dh.piezas) || 1) * multiplier,
        }),
      );
      if (type === "TOTAL_NETO") {
        (m.castillos || []).forEach((dc) =>
          rowsToCopy.push({
            ...base,
            claveLoc:
              (base.claveLoc ? base.claveLoc + " " : "") + "(- Castillo)",
            largo: dc.ancho,
            ancho: "",
            alto: m.alto,
            piezas: -(parseFloat(dc.piezas) || 1) * multiplier,
          }),
        );
      }
    });
    setClipboardRows(rowsToCopy);
    copyRowsToExcelClipboard(rowsToCopy);
    setCopyStatus(`copied-wall-${type}-${value || "all"}`);
    setTimeout(() => setCopyStatus(""), 2000);
  };

  const handleCopyStructure = (tipoKey, material) => {
    let filtered = [];
    if (tipoKey === "GLOBAL_CIMENTACION")
      filtered = estructuras.filter((e) =>
        [
          "zapata aislada",
          "zapata corrida",
          "contratrabe",
          "dado",
          "pilas",
          "pila",
        ].includes((e.tipo || "Sin Seleccionar").toLowerCase()),
      );
    else if (tipoKey === "GLOBAL_ESTRUCTURA_ACERO")
      filtered = estructuras.filter((e) =>
        [
          "losa",
          "losa nervada",
          "losa de vigueta",
          "trabe",
          "nervadura",
          "columna",
          "columna circular",
          "escalera papelillo",
          "rampa de escalera",
        ].includes((e.tipo || "Sin Seleccionar").toLowerCase()),
      );
    else if (tipoKey === "GLOBAL_ESTRUCTURA")
      filtered = estructuras.filter(
        (e) =>
          ![
            "zapata aislada",
            "zapata corrida",
            "contratrabe",
            "dado",
            "muro",
            "muro curvo",
          ].includes((e.tipo || "Sin Seleccionar").toLowerCase()),
      );
    else if (tipoKey === "Estructura (Losas y Trabes)")
      filtered = estructuras.filter((e) =>
        ["losa", "losa nervada", "trabe", "nervadura"].includes(
          (e.tipo || "").toLowerCase(),
        ),
      );
    else if (tipoKey === "Concreto en Estructuras")
      filtered = estructuras.filter((e) =>
        [
          "losa",
          "losa nervada",
          "losa de vigueta",
          "trabe",
          "nervadura",
          "escalera papelillo",
          "rampa de escalera",
        ].includes((e.tipo || "").toLowerCase()),
      );
    else if (tipoKey === "Columnas")
      filtered = estructuras.filter((e) =>
        ["columna", "columna circular"].includes((e.tipo || "").toLowerCase()),
      );
    else if (tipoKey === "Acero en Estructuras")
      filtered = estructuras.filter((e) =>
        [
          "losa",
          "losa nervada",
          "losa de vigueta",
          "trabe",
          "nervadura",
          "columna",
          "columna circular",
          "escalera papelillo",
          "rampa de escalera",
        ].includes((e.tipo || "").toLowerCase()),
      );
    else if (tipoKey === "Cimentación")
      filtered = estructuras.filter((e) =>
        [
          "zapata aislada",
          "zapata corrida",
          "contratrabe",
          "dado",
          "pilas",
          "pila",
        ].includes((e.tipo || "").toLowerCase()),
      );
    else
      filtered = estructuras.filter((e) => {
        const t = (e.tipo || "Sin Seleccionar").toLowerCase();
        const tk = tipoKey.toLowerCase();

        let eGroup = t;
        if (t === "columna" || t === "columnas") eGroup = "columnas";
        else if (t === "columna circular" || t === "columnas circulares") eGroup = "columnas circulares";
        else if (t === "muro" || t === "muro curvo" || t === "muros" || t === "muros curvos") eGroup = "muros";
        else if (
          t === "losa" ||
          t === "losas" ||
          t === "losa nervada" ||
          t === "escalera papelillo" ||
          t === "rampa de escalera"
        )
          eGroup = "losas";
        else if (t.includes("vigueta")) eGroup = "losas de vigueta";
        else if (t === "trabe" || t === "trabes") eGroup = "trabes";
        else if (t === "nervadura" || t === "nervaduras") eGroup = "nervaduras";
        else if (t === "zapata aislada")
          eGroup = `zapatas aisladas (esp: ${parseFloat(e.alto || 0).toFixed(2)}m)`;
        else if (t === "zapata corrida")
          eGroup = `zapatas corridas (esp: ${parseFloat(e.alto || 0).toFixed(2)}m)`;
        else if (t === "contratrabe") eGroup = "contratrabes";
        else if (t === "dado") eGroup = "dados";
        else if (t === "pila" || t === "pilas") eGroup = "pilas";

        return eGroup === tk;
      });

    let rowsToCopy = [];
    if (material === "concreto") {
      rowsToCopy = filtered.flatMap((e) => {
        const base = { eje: e.eje || "", claveLoc: e.clave || "" };
        const pMain = parseFloat(e.piezas) || 1;
        let lValue = getEffectiveLargo(e);
        let aValue = e.ancho;

        const t = e.tipo?.toLowerCase() || "";

        let effectiveH = parseFloat(e.alto) || 0;
        if (
          t === "nervadura" ||
          t.includes("vigueta h=") ||
          t.includes("vigueta h35") ||
          t.includes("vigueta h25")
        )
          return [];
        const thicknessZapata = parseFloat(e.structEspesorZapata) || 0;
        const thicknessLosa = parseFloat(e.structEspesorLosa) || 0;
        const slabTypeInput = (e.structTipoLosa || "").toLowerCase();
        const isEsquina = e.structEsquinaLosa === true;
        effectiveH = Math.max(0, effectiveH - thicknessZapata);
        if (slabTypeInput) {
          if (slabTypeInput !== "vigueta") {
            if (isEsquina && (t.includes("muro") || t.includes("columna") || t === "dado" || t === "trabe" || t === "contratrabe")) {
              effectiveH -= thicknessLosa / 2;
            } else {
              effectiveH -= thicknessLosa;
            }
          }
        } else {
          const simpleDiscount = parseFloat(e.descuentoLosa) || 0;
          effectiveH = Math.max(0, effectiveH - simpleDiscount);
        }
        effectiveH = Math.max(0, effectiveH);

        if (t === "columna circular" || t === "pilas" || t === "pila") {
          const radio = lValue / 2;
          lValue = Math.PI * radio * radio;
          aValue = "";
        }

        let rows = [
          {
            ...base,
            largo: lValue,
            ancho: aValue,
            alto: effectiveH,
            piezas: e.piezas,
          },
        ];
        if (
          e.tipo?.toLowerCase() === "losa nervada" &&
          e.casetones &&
          e.casetones.length > 0
        ) {
          const casetonesRows = e.casetones.map((c) => ({
            ...base,
            claveLoc:
              (base.claveLoc ? base.claveLoc + " " : "") +
              `(- Casetón ${c.clave || ""})`.trim(),
            largo: c.largo,
            ancho: c.ancho,
            alto: c.alto,
            piezas: -(parseFloat(c.piezas) || 1) * pMain,
          }));
          rows = rows.concat(casetonesRows);
        }
        return rows;
      });
    } else if (
      [
        "cimbra",
        "cimbraEscaleraPapelillo",
        "cimbraRampaEscalera",
        "obturacionMuros",
        "mallaRefuerzo",
        "emplayerColumnas",
        "andamiajeColumnas",
        "andamiajeTrabes",
      ].includes(material) ||
      material.startsWith("cimbraMuros_") ||
      material.startsWith("cimbraColumnas_") ||
      material.startsWith("cimbraMuroCurvo_")
    ) {
      rowsToCopy = filtered.flatMap((e) => {
        const l = getEffectiveLargo(e),
          a = e.ancho,
          hOrig = parseFloat(e.alto) || 0,
          p = parseFloat(e.piezas) || 1,
          t = (e.tipo || "").toLowerCase(),
          base = { eje: e.eje || "", claveLoc: e.clave || "" };

        const thicknessZapata = parseFloat(e.structEspesorZapata) || 0;
        const thicknessLosa = parseFloat(e.structEspesorLosa) || 0;
        const isEsquina = e.structEsquinaLosa === true;
        
        let hBase = Math.max(0, hOrig - thicknessZapata);
        let h_eff = hBase;
        if (e.structTipoLosa) {
          if (isEsquina && (t.includes("muro") || t.includes("columna") || t === "dado" || t === "trabe" || t === "contratrabe")) {
            h_eff = hBase - (thicknessLosa / 2);
          } else {
            h_eff = hBase - thicknessLosa;
          }
        } else {
          const discount = parseFloat(e.descuentoLosa) || 0;
          h_eff = Math.max(0, hBase - discount);
        }
        h_eff = Math.max(0, h_eff);

        let shouldInclude = true;
        let outputAlto = h_eff;

        if (
          material.startsWith("cimbraMuros_") &&
          t !== "muro" &&
          t !== "muros"
        )
          return [];
        if (material === "cimbraEscaleraPapelillo" && t !== "escalera papelillo") return [];
        if (material === "cimbraRampaEscalera" && t !== "rampa de escalera") return [];
        if (
          material.startsWith("cimbraMuroCurvo_") &&
          t !== "muro curvo" &&
          t !== "muros curvos"
        )
          return [];

        if (
          material.startsWith("cimbraMuros_") ||
          material.startsWith("cimbraColumnas_") ||
          material.startsWith("cimbraMuroCurvo_")
        ) {
          if (material.endsWith("0_3")) {
            outputAlto = Math.min(3, h_eff);
            shouldInclude = outputAlto > 0;
          } else if (material.endsWith("3_6")) {
            if (h_eff <= 3) return [];
            outputAlto = Math.min(3, h_eff - 3);
          } else if (material.endsWith("6_9")) {
            if (h_eff <= 6) return [];
            outputAlto = h_eff - 6;
          }
        }
        if (!shouldInclude) return [];

        let labelSuffix = "";
        if (material === "obturacionMuros") labelSuffix = " (Obturación)";
        else if (material === "mallaRefuerzo") labelSuffix = " (Malla)";
        else if (material === "emplayerColumnas") labelSuffix = " (Emplayer)";
        else if (material === "pasosMuros") labelSuffix = " (Pasos)";
        else if (material.startsWith("andamiaje")) labelSuffix = " (Andamiaje)";
        else if (material.startsWith("cimbra")) labelSuffix = " (Cimbra)";

        if (t === "muro" || t === "muro curvo")
          return [
            {
              ...base,
              claveLoc: base.claveLoc
                ? base.claveLoc + labelSuffix
                : labelSuffix.trim(),
              largo: l,
              ancho: 2,
              alto: outputAlto,
              piezas: p,
            },
          ];
        if (
          t === "losa" ||
          t === "losa nervada" ||
          t.includes("vigueta") ||
          t === "escalera papelillo" ||
          t === "rampa de escalera"
        )
          return [
            {
              ...base,
              claveLoc: base.claveLoc ? base.claveLoc + " (Fondo)" : "Fondo",
              largo: l,
              ancho: a,
              alto: "",
              piezas: p,
            },
          ];
        if (t === "columna circular") {
          return [
            {
              ...base,
              claveLoc: base.claveLoc
                ? base.claveLoc + labelSuffix
                : labelSuffix.trim(),
              largo: Math.round(Math.PI * l * 100) / 100,
              ancho: "",
              alto: outputAlto,
              piezas: p,
            },
          ];
        }
        if (t === "columna" || t === "dado") {
          return [
            {
              ...base,
              claveLoc: base.claveLoc
                ? base.claveLoc +
                  " (Lados X" +
                  (labelSuffix ? labelSuffix : "") +
                  ")"
                : "Lados X" + (labelSuffix ? labelSuffix : ""),
              largo: l,
              ancho: 2,
              alto: outputAlto,
              piezas: p,
            },
            {
              ...base,
              claveLoc: base.claveLoc
                ? base.claveLoc +
                  " (Lados Y" +
                  (labelSuffix ? labelSuffix : "") +
                  ")"
                : "Lados Y" + (labelSuffix ? labelSuffix : ""),
              largo: a,
              ancho: 2,
              alto: outputAlto,
              piezas: p,
            },
          ];
        }
        if (t === "zapata aislada" || t === "zapata corrida")
          return [
            {
              ...base,
              claveLoc: base.claveLoc
                ? base.claveLoc + " (Lados X" + labelSuffix + ")"
                : "Lados X" + labelSuffix,
              largo: l,
              ancho: 2,
              alto: h_eff,
              piezas: p,
            },
            {
              ...base,
              claveLoc: base.claveLoc
                ? base.claveLoc + " (Lados Y" + labelSuffix + ")"
                : "Lados Y" + labelSuffix,
              largo: a,
              ancho: 2,
              alto: h_eff,
              piezas: p,
            },
          ];
        if (t === "trabe" || t === "contratrabe")
          return [
            {
              ...base,
              claveLoc: base.claveLoc
                ? base.claveLoc + " (Fondo" + labelSuffix + ")"
                : "Fondo" + labelSuffix,
              largo: l,
              ancho: a,
              alto: "",
              piezas: p,
            },
            {
              ...base,
              claveLoc: base.claveLoc
                ? base.claveLoc + " (Laterales" + labelSuffix + ")"
                : "Laterales" + labelSuffix,
              largo: l,
              ancho: 2,
              alto: h_eff,
              piezas: p,
            },
          ];
        return [];
      });
    } else if (material && material.startsWith("areaVigueta") || material === "areaDobleViguetaH25" || material === "areaDobleViguetaH35") {
      rowsToCopy = filtered.flatMap((e) => {
        const l = getEffectiveLargo(e);
        const a = parseFloat(e.ancho) || 0;
        const p = parseFloat(e.piezas) || 1;
        
        const isH25 = material.includes("H25");
        const isH35 = material.includes("H35");
        const isDoble = material.includes("Doble");
        
        const tipoMatch = (e.tipo || "").toLowerCase();
        
        let match = false;
        const hasH25 = tipoMatch.includes("h=25") || tipoMatch.includes("h25");
        const hasH35 = tipoMatch.includes("h=35") || tipoMatch.includes("h35");
        
        if (isDoble && isH25 && tipoMatch.includes("doble") && hasH25) match = true;
        else if (isDoble && isH35 && tipoMatch.includes("doble") && hasH35) match = true;
        else if (!isDoble && isH25 && !tipoMatch.includes("doble") && hasH25) match = true;
        else if (!isDoble && isH35 && !tipoMatch.includes("doble") && hasH35) match = true;
        else if (material === "areaViguetasGENERIC" && !hasH25 && !hasH35) match = true;
        else if (material === "areaViguetas") match = true; // fallback
        
        if (!match) return [];
        return [
          {
            eje: e.eje || "",
            claveLoc: e.clave || "",
            largo: l,
            ancho: a,
            alto: "",
            piezas: p,
          },
        ];
      });
    } else if (material === "pasosMuros") {
      rowsToCopy = filtered.flatMap((e) => {
        if (
          (e.tipo || "").toLowerCase() === "muro" ||
          (e.tipo || "").toLowerCase() === "muro curvo"
        ) {
          const l = getEffectiveLargo(e);
          const p = parseFloat(e.piezas) || 1;
          const pasosPiece = Math.floor(l / 4);
          if (pasosPiece > 0) {
            return [
              {
                eje: e.eje || "",
                claveLoc: (e.clave ? e.clave + " " : "") + "(Pasos)",
                largo: l,
                ancho: "",
                alto: "",
                piezas: p,
                isPasos: true,
              },
            ];
          }
        }
        return [];
      });
    } else if (material === "andamiajeColumnas") {
      rowsToCopy = filtered.flatMap((e) => {
        const h = parseFloat(e.alto) || 0;
        const p = parseFloat(e.piezas) || 1;
        const t = (e.tipo || "").toLowerCase();
        if ((t === "columna" || t === "columna circular") && h > 3) {
          return [
            {
              eje: e.eje || "",
              claveLoc: (e.clave || "") + " (Andamiaje)",
              largo: 1,
              ancho: "",
              alto: "",
              piezas: p,
            },
          ];
        }
        return [];
      });
    } else if (material === "andamiajeTrabes") {
      rowsToCopy = filtered.flatMap((e) => {
        const t = (e.tipo || "").toLowerCase();
        if (t === "trabe") {
          const p = parseFloat(e.piezas) || 1;
          const l = getEffectiveLargo(e);
          return [
            {
              eje: e.eje || "",
              claveLoc: (e.clave || "") + " (Andamiaje)",
              largo: l,
              ancho: "",
              alto: "",
              piezas: p,
            },
          ];
        }
        return [];
      });
    } else if (
      typeof material === "string" &&
      material.startsWith("anclajeVar_")
    ) {
      const targetNumV = material.split("_")[1] || "";
      if (!targetNumV) return [];
      rowsToCopy = filtered.flatMap((e) => {
        const t = (e.tipo || "").toLowerCase();
        if (t === "muro" || t === "muro curvo") {
          const pElem = parseFloat(e.piezas) || 1;
          let pzas = 0;
          let acerosCorregidos = e.aceros || [];
          const validTypesOptions = getSteelTypesForElement(e.tipo);
          const validTypes = validTypesOptions.map((xt) => xt.toLowerCase());
          if (validTypes.length > 0) {
            acerosCorregidos = acerosCorregidos.map((a) => {
              if (!validTypes.includes((a.tipo || "").toLowerCase())) {
                return { ...a, tipo: validTypesOptions[0] };
              }
              return a;
            });
          }
          acerosCorregidos.forEach((a) => {
            const aTipo = (a.tipo || "").toLowerCase().trim();
            const isAnclajeTarget =
              aTipo.includes("vertical") ||
              aTipo.includes("remetido") ||
              aTipo.includes("longitudinal") ||
              aTipo.includes("principal") ||
              aTipo.includes("bastones") ||
              aTipo === "";
            const nv = (a.numVarilla || "").toString();
            if (isAnclajeTarget && nv === targetNumV) {
              pzas += parseFloat(a.piezas) || 0;
            }
          });
          if (pzas > 0) {
            return [
              {
                eje: e.eje || "",
                claveLoc: (e.clave || "") + ` (Anclaje Varilla #${targetNumV})`,
                largo: pzas.toFixed(2),
                ancho: "",
                alto: "",
                piezas: pElem,
              },
            ];
          }
        }
        return [];
      });
    } else if (material === "cimbraFrontera") {
      rowsToCopy = filtered.flatMap((e) => {
        const l = getEffectiveLargo(e),
          a = e.ancho,
          p = parseFloat(e.piezas) || 1,
          t = (e.tipo || "").toLowerCase(),
          base = { eje: e.eje || "", claveLoc: e.clave || "" };
        if (
          t === "losa" ||
          t === "losa nervada" ||
          t.includes("vigueta") ||
          t === "escalera papelillo" ||
          t === "rampa de escalera"
        )
          return [
            {
              ...base,
              claveLoc: base.claveLoc
                ? base.claveLoc + " (Frontera X)"
                : "Frontera X",
              largo: l,
              ancho: 2,
              alto: "",
              piezas: p,
            },
            {
              ...base,
              claveLoc: base.claveLoc
                ? base.claveLoc + " (Frontera Y)"
                : "Frontera Y",
              largo: a,
              ancho: 2,
              alto: "",
              piezas: p,
            },
          ];
        return [];
      });
    } else if (
      ["excavacion", "plantilla", "afine", "relleno"].includes(material)
    ) {
      rowsToCopy = filtered.flatMap((e) => {
        const l = getEffectiveLargo(e),
          a = parseFloat(e.ancho) || 0,
          h = parseFloat(e.alto) || 0,
          p = parseFloat(e.piezas) || 1,
          base = { eje: e.eje || "", claveLoc: e.clave || "" },
          tipo = (e.tipo || "").toLowerCase();

        let excavacionVol = 0;
        let pLargo = 0,
          pAncho = 0,
          eLargo = 0,
          eAncho = 0,
          eAlto = 0;
        if (tipo === "pilas" || tipo === "pila") {
          const diamExc = l + 1.0;
          const radioExc = diamExc / 2;
          const areaExc = Math.PI * radioExc * radioExc;
          excavacionVol = areaExc * h;
          pLargo = diamExc; // Not used for plantilla for pilas, but setting to avoid bugs
          pAncho = diamExc;
          eLargo = Math.PI * radioExc * radioExc;
          eAncho = "";
          eAlto = h;
        } else {
          const plantillaArea = (l + 0.5) * (a + 0.5);
          excavacionVol = plantillaArea * 1.5;
          pLargo = l + 0.5;
          pAncho = a + 0.5;
          eLargo = l + 0.5;
          eAncho = a + 0.5;
          eAlto = 1.5;
        }

        const c = calcConcreto(e) / p;
        const rellenoVol = Math.max(0, excavacionVol - c);

        if (material === "plantilla") {
          if (tipo === "pilas" || tipo === "pila") return []; // No plantilla for pilas
          return [
            {
              ...base,
              claveLoc: base.claveLoc
                ? base.claveLoc + " (Plantilla)"
                : "Plantilla",
              largo: pLargo,
              ancho: pAncho,
              alto: "",
              piezas: p,
            },
          ];
        }
        if (material === "afine") {
          if (tipo === "pilas" || tipo === "pila") return []; // No afine for pilas
          return [
            {
              ...base,
              claveLoc: base.claveLoc ? base.claveLoc + " (Afine)" : "Afine",
              largo: pLargo,
              ancho: pAncho,
              alto: "",
              piezas: p,
            },
          ];
        }
        if (material === "excavacion")
          return [
            {
              ...base,
              claveLoc: base.claveLoc
                ? base.claveLoc + " (Excavación)"
                : "Excavación",
              largo:
                tipo === "pilas" || tipo === "pila"
                  ? excavacionVol.toFixed(2)
                  : eLargo,
              ancho: tipo === "pilas" || tipo === "pila" ? "" : eAncho,
              alto: tipo === "pilas" || tipo === "pila" ? "" : eAlto,
              piezas: p,
            },
          ];
        if (material === "relleno")
          return [
            {
              ...base,
              claveLoc: base.claveLoc
                ? base.claveLoc + " (Relleno)"
                : "Relleno",
              largo: rellenoVol.toFixed(2),
              ancho: "",
              alto: "",
              piezas: p,
            },
          ];
        return [];
      });
    } else if (material === "caseton") {
      rowsToCopy = filtered.flatMap((e) => {
        const p = parseFloat(e.piezas) || 1,
          base = { eje: e.eje || "", claveLoc: e.clave || "" };
        return (e.casetones || []).map((c) => ({
          ...base,
          claveLoc:
            (base.claveLoc ? base.claveLoc + " " : "") +
            `(Casetón ${c.clave || ""})`.trim(),
          largo: c.largo,
          ancho: c.ancho,
          alto: c.alto,
          piezas: (parseFloat(c.piezas) || 1) * p,
        }));
      });
    } else if (material.startsWith("acero-")) {
      const numVar = material.split("-")[1];
      rowsToCopy = filtered.flatMap((e) => {
        let acerosCorregidos = e.aceros || [];
        const validTypesOptions = getSteelTypesForElement(e.tipo);
        const validTypes = validTypesOptions.map((xt) => xt.toLowerCase());
        if (validTypes.length > 0) {
          acerosCorregidos = acerosCorregidos.map((a) => {
            if (!validTypes.includes((a.tipo || "").toLowerCase())) {
              return { ...a, tipo: validTypesOptions[0] };
            }
            return a;
          });
        }
        return acerosCorregidos
          .filter((a) => numVar === "total" || a.numVarilla === numVar)
          .map((a) => ({
            eje: e.eje || "",
            claveLoc: (
              (e.clave ? e.clave + " " : "") +
              `${a.tipo || "Acero"} #${a.numVarilla || "-"}`
            ).trim(),
            largo: calcAceroItem(a).mlPorPieza,
            ancho: a.piezas,
            kg_ml: PESOS_VARILLA[a.numVarilla || "-"] || 0,
            alto: "",
            piezas: e.piezas || 1,
          }));
      });
    }
    setClipboardRows(rowsToCopy);
    copyRowsToExcelClipboard(rowsToCopy);
    setCopyStatus(`copied-${tipoKey}-${material}`);
    setTimeout(() => setCopyStatus(""), 2000);
  };

  const updateEstructuraField = useCallback(
    (structId, field, value) => {
      updateActiveEstructuras((prev) =>
        prev.map((s) => {
          if (s.id !== structId) return s;
          const newS = { ...s, [field]: value };
          if (["largo", "ancho", "alto", "tipo"].includes(field)) {
            newS.aceros = (newS.aceros || []).map((a) => {
              const updatedA = { ...a };

              if (field === "tipo") {
                const validTypesOptions = getSteelTypesForElement(newS.tipo);
                const validTypes = validTypesOptions.map((t) =>
                  t.toLowerCase(),
                );
                if (
                  validTypes.length > 0 &&
                  !validTypes.includes((updatedA.tipo || "").toLowerCase())
                ) {
                  updatedA.tipo = validTypesOptions[0];
                }
              }

              let tAcero = (updatedA.tipo || "").toLowerCase();
              const tEst = (newS.tipo || "").toLowerCase(),
                l = parseFloat(newS.largo) || 0,
                an = parseFloat(newS.ancho) || 0,
                al = parseFloat(newS.alto) || 0;
              if (tAcero === "longitudinal") {
                if (
                  [
                    "columna",
                    "columna circular",
                    "dado",
                    "pilas",
                    "pila",
                  ].includes(tEst)
                )
                  updatedA.longitud = al > 0 ? al.toString() : "";
                else if (["trabe", "nervadura", "contratrabe"].includes(tEst))
                  updatedA.longitud = l > 0 ? l.toString() : "";
                else if (
                  [
                    "zapata aislada",
                    "zapata corrida",
                    "losa",
                    "losa nervada",
                  ].includes(tEst)
                ) {
                  updatedA.anchoCalc = an > 0 ? an.toString() : "";
                  updatedA.longitud = l > 0 ? l.toString() : "";
                }
                const longCalc = parseFloat(updatedA.longitud) || 0;
                updatedA.traslapes =
                  longCalc >= 12 ? Math.floor(longCalc / 12).toString() : "0";
              }
              if (
                tAcero === "transversal" &&
                [
                  "zapata aislada",
                  "zapata corrida",
                  "losa",
                  "losa nervada",
                ].includes(tEst)
              ) {
                updatedA.anchoCalc = l > 0 ? l.toString() : "";
                updatedA.longitud = an > 0 ? an.toString() : "";
              }
              if (tAcero === "estribos" || tAcero === "zuncho") {
                updatedA.ganchos = "0";
                updatedA.anclaje = "0";
                updatedA.traslapes = "0";
                let perimetro = 0;
                if (["columna", "dado"].includes(tEst)) {
                  perimetro = (l + an) * 2;
                  if (
                    !updatedA.anchoCalc ||
                    parseFloat(updatedA.anchoCalc) === parseFloat(s.alto)
                  )
                    updatedA.anchoCalc = al > 0 ? al.toString() : "";
                } else if (
                  tEst === "columna circular" ||
                  tEst === "pilas" ||
                  tEst === "pila"
                ) {
                  perimetro = l * Math.PI;
                  if (
                    !updatedA.anchoCalc ||
                    parseFloat(updatedA.anchoCalc) === parseFloat(s.alto)
                  )
                    updatedA.anchoCalc = al > 0 ? al.toString() : "";
                } else if (
                  ["trabe", "nervadura", "contratrabe"].includes(tEst)
                ) {
                  perimetro = (an + al) * 2;
                  if (
                    !updatedA.anchoCalc ||
                    parseFloat(updatedA.anchoCalc) === parseFloat(s.largo)
                  )
                    updatedA.anchoCalc = l > 0 ? l.toString() : "";
                }
                if (perimetro > 0)
                  updatedA.longitud = (perimetro - 0.06 + 0.12).toFixed(2);
              }
              if (tEst === "muro") {
                if (tAcero === "vertical") {
                  updatedA.anchoCalc = l > 0 ? l.toString() : "";
                  updatedA.longitud = al > 0 ? al.toString() : "";
                } else if (tAcero === "horizontal") {
                  updatedA.anchoCalc = al > 0 ? al.toString() : "";
                  updatedA.longitud = l > 0 ? l.toString() : "";
                }
              }
              if (tAcero === "bastones" || tAcero === "refuerzo adicional")
                updatedA.longitud = updatedA.anchoCalc || "";
              const isPorPiezas =
                tAcero === "bastones" ||
                tAcero === "refuerzo adicional" ||
                (tAcero === "longitudinal" &&
                  [
                    "columna",
                    "columna circular",
                    "pilas",
                    "pila",
                    "dado",
                    "trabe",
                    "nervadura",
                    "contratrabe",
                  ].includes(tEst));
              if (!isPorPiezas) {
                const anchoCalc = parseFloat(updatedA.anchoCalc),
                  sep = parseFloat(updatedA.separacion);
                if (
                  !isNaN(anchoCalc) &&
                  anchoCalc > 0 &&
                  !isNaN(sep) &&
                  sep > 0
                )
                  updatedA.piezas = Math.floor(anchoCalc / (sep / 100)) + 2;
              }
              return updatedA;
            });
          }
          return newS;
        }),
      );
    },
    [updateActiveEstructuras],
  );

  const handleStructAltoDoubleClick = useCallback(
    (structId) => {
      const struct = estructuras.find((s) => s.id === structId);
      if (!struct) return;

      const tipo = (struct.tipo || "").toLowerCase();
      if (
        !["trabe", "contratrabe", "columna", "columna circular", "muro", "muro curvo", "nervadura"].includes(tipo.trim())
      ) {
        return;
      }
      setStructAdjModal({
        id: structId,
        tipo: tipo.trim(),
        tipoLosa: struct.structTipoLosa || "maciza",
        esquinaLosa: struct.structEsquinaLosa || false,
        espesorLosa: struct.structEspesorLosa || "0.10",
        espesorZapata: struct.structEspesorZapata || "0",
      });
    },
    [estructuras],
  );

  const updateSteelItem = useCallback(
    (structId, itemId, field, value) => {
      updateActiveEstructuras((prev) =>
        prev.map((s) => {
          if (s.id !== structId) return s;
          return {
            ...s,
            aceros: (s.aceros || []).map((a) => {
              if (a.id !== itemId) return a;
              const updatedA = { ...a, [field]: value },
                tipoAcero = (updatedA.tipo || "").toLowerCase(),
                tipoEstructura = (s.tipo || "").toLowerCase();
              const isPorPiezas =
                tipoAcero === "bastones" ||
                tipoAcero === "refuerzo adicional" ||
                (tipoAcero === "longitudinal" &&
                  [
                    "columna",
                    "columna circular",
                    "pilas",
                    "pila",
                    "dado",
                    "trabe",
                    "nervadura",
                    "contratrabe",
                  ].includes(tipoEstructura));
              if (field === "numVarilla") {
                const oldDefault = ANCLAJE_DEFAULT[a.numVarilla] || "0.30";
                if (
                  !a.anclaje ||
                  parseFloat(a.anclaje) === parseFloat(oldDefault)
                ) {
                  updatedA.anclaje = ANCLAJE_DEFAULT[value] || "";
                }
              }
              if (field === "tipo") {
                const l = parseFloat(s.largo) || 0,
                  an = parseFloat(s.ancho) || 0,
                  al = parseFloat(s.alto) || 0;
                if (
                  tipoAcero === "estribos" ||
                  tipoAcero === "zuncho" ||
                  tipoAcero === "grapas"
                ) {
                  updatedA.ganchos = "0";
                  updatedA.anclaje = "0";
                  updatedA.traslapes = "0";
                  let perimetro = 0;
                  if (["columna", "dado"].includes(tipoEstructura)) {
                    perimetro = (l + an) * 2;
                    if (!updatedA.anchoCalc && al > 0)
                      updatedA.anchoCalc = al.toString();
                  } else if (
                    tipoEstructura === "columna circular" ||
                    tipoEstructura === "pilas" ||
                    tipoEstructura === "pila"
                  ) {
                    perimetro = l * Math.PI;
                    if (!updatedA.anchoCalc && al > 0)
                      updatedA.anchoCalc = al.toString();
                  } else if (
                    ["trabe", "nervadura", "contratrabe"].includes(
                      tipoEstructura,
                    )
                  ) {
                    perimetro = (an + al) * 2;
                    if (!updatedA.anchoCalc && l > 0)
                      updatedA.anchoCalc = l.toString();
                  }
                  if (perimetro > 0)
                    updatedA.longitud = (perimetro - 0.06 + 0.12).toFixed(2);
                } else if (tipoAcero === "longitudinal") {
                  if (
                    [
                      "columna",
                      "columna circular",
                      "pilas",
                      "pila",
                      "dado",
                    ].includes(tipoEstructura)
                  )
                    updatedA.longitud = al > 0 ? al.toString() : "";
                  else if (
                    ["trabe", "nervadura", "contratrabe"].includes(
                      tipoEstructura,
                    )
                  )
                    updatedA.longitud = l > 0 ? l.toString() : "";
                  else if (
                    [
                      "zapata aislada",
                      "zapata corrida",
                      "losa",
                      "losa nervada",
                      "escalera papelillo",
                      "rampa de escalera",
                    ].includes(tipoEstructura)
                  ) {
                    updatedA.anchoCalc = an > 0 ? an.toString() : "";
                    updatedA.longitud = l > 0 ? l.toString() : "";
                  }
                } else if (
                  tipoAcero === "transversal" &&
                  [
                    "zapata aislada",
                    "zapata corrida",
                    "losa",
                    "losa nervada",
                    "escalera papelillo",
                    "rampa de escalera",
                  ].includes(tipoEstructura)
                ) {
                  updatedA.anchoCalc = l > 0 ? l.toString() : "";
                  updatedA.longitud = an > 0 ? an.toString() : "";
                } else if (
                  tipoEstructura === "muro" ||
                  tipoEstructura === "muro curvo"
                ) {
                  if (tipoAcero === "vertical") {
                    updatedA.anchoCalc = l > 0 ? l.toString() : "";
                    updatedA.longitud = al > 0 ? al.toString() : "";
                  } else if (tipoAcero === "horizontal") {
                    updatedA.anchoCalc = al > 0 ? al.toString() : "";
                    updatedA.longitud = l > 0 ? l.toString() : "";
                  }
                }
                if (
                  tipoAcero === "bastones" ||
                  tipoAcero === "refuerzo adicional"
                )
                  updatedA.longitud = updatedA.anchoCalc || "";
              }
              if (
                (tipoAcero === "bastones" ||
                  tipoAcero === "refuerzo adicional") &&
                field === "anchoCalc"
              )
                updatedA.longitud = value;
              if (
                !isPorPiezas &&
                (field === "anchoCalc" ||
                  field === "separacion" ||
                  field === "tipo")
              ) {
                const ancho = parseFloat(updatedA.anchoCalc),
                  sep = parseFloat(updatedA.separacion);
                if (!isNaN(ancho) && ancho > 0 && !isNaN(sep) && sep > 0)
                  updatedA.piezas = Math.floor(ancho / (sep / 100)) + 2;
              }
              if (
                field === "longitud" ||
                field === "tipo" ||
                field === "anchoCalc"
              ) {
                const longCalc = parseFloat(updatedA.longitud) || 0;
                updatedA.traslapes =
                  longCalc >= 12 ? Math.floor(longCalc / 12).toString() : "0";
              }
              if (
                tipoAcero === "estribos" ||
                tipoAcero === "zuncho" ||
                tipoAcero === "grapas"
              ) {
                updatedA.ganchos = "0";
                updatedA.anclaje = "0";
                updatedA.traslapes = "0";
              }
              return updatedA;
            }),
          };
        }),
      );
    },
    [updateActiveEstructuras],
  );

  const handleCopyStructureRow = useCallback((row) => {
    setCopiedStructureRow(row);
    setCopyStatus(`copied-struct-${row.id}`);
    setTimeout(() => setCopyStatus(""), 2000);
  }, []);
  const handlePasteStructureRow = useCallback(
    (targetId) => {
      if (!copiedStructureRow) return;
      updateActiveEstructuras((prev) =>
        prev.map((s) => {
          if (s.id !== targetId) return s;
          const clonedAceros = (copiedStructureRow.aceros || []).map((a) => ({
            ...a,
            id: `AC-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
          }));
          const clonedCasetones = (copiedStructureRow.casetones || []).map(
            (c) => ({
              ...c,
              id: `CAS-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
            }),
          );
          return {
            ...s,
            tipo: copiedStructureRow.tipo || "",
            largo: copiedStructureRow.largo || "",
            ancho: copiedStructureRow.ancho || "",
            alto: copiedStructureRow.alto || "",
            piezas: copiedStructureRow.piezas || "1",
            i: copiedStructureRow.i || false,
            aceros: clonedAceros,
            casetones: clonedCasetones,
          };
        }),
      );
      setCopyStatus(`pasted-struct-${targetId}`);
      setTimeout(() => setCopyStatus(""), 2000);
    },
    [copiedStructureRow, updateActiveEstructuras],
  );

  const resetSteelFieldToDefault = useCallback(
    (structId, itemId, field) => {
      const s = estructuras.find((e) => e.id === structId);
      if (!s) return;
      const a = s.aceros?.find((x) => x.id === itemId);
      if (!a) return;
      const tipoEst = (s.tipo || "").toLowerCase(),
        tipoAcero = (a.tipo || "").toLowerCase(),
        l = parseFloat(s.largo) || 0,
        an = parseFloat(s.ancho) || 0,
        al = parseFloat(s.alto) || 0;
      let defVal = "";
      if (field === "anchoCalc") {
        if (
          tipoAcero === "estribos" ||
          tipoAcero === "zuncho" ||
          tipoAcero === "grapas"
        ) {
          if (
            ["columna", "dado", "columna circular", "pilas", "pila"].includes(
              tipoEst,
            )
          )
            defVal = al > 0 ? al.toString() : "";
          else if (["trabe", "nervadura", "contratrabe"].includes(tipoEst))
            defVal = l > 0 ? l.toString() : "";
        } else if (tipoEst === "muro" || tipoEst === "muro curvo") {
          if (tipoAcero === "vertical") defVal = l > 0 ? l.toString() : "";
          else if (tipoAcero === "horizontal")
            defVal = al > 0 ? al.toString() : "";
        } else if (tipoAcero === "longitudinal") {
          if (
            ["columna", "columna circular", "pilas", "pila", "dado"].includes(
              tipoEst,
            )
          )
            defVal = al > 0 ? al.toString() : "";
          else if (["trabe", "nervadura", "contratrabe"].includes(tipoEst))
            defVal = l > 0 ? l.toString() : "";
          else if (
            [
              "zapata aislada",
              "zapata corrida",
              "losa",
              "losa nervada",
              "escalera papelillo",
              "rampa de escalera",
            ].includes(tipoEst)
          )
            defVal = an > 0 ? an.toString() : "";
        } else if (
          tipoAcero === "transversal" &&
          [
            "zapata aislada",
            "zapata corrida",
            "losa",
            "losa nervada",
            "escalera papelillo",
            "rampa de escalera",
          ].includes(tipoEst)
        )
          defVal = l > 0 ? l.toString() : "";
        else if (tipoAcero === "bastones" || tipoAcero === "refuerzo adicional")
          defVal = l > 0 ? l.toString() : "";
      } else if (field === "longitud") {
        if (
          tipoAcero === "estribos" ||
          tipoAcero === "zuncho" ||
          tipoAcero === "grapas"
        ) {
          let perimetro = 0;
          if (["columna", "dado"].includes(tipoEst)) perimetro = (l + an) * 2;
          else if (
            tipoEst === "columna circular" ||
            tipoEst === "pilas" ||
            tipoEst === "pila"
          )
            perimetro = l * Math.PI;
          else if (["trabe", "nervadura", "contratrabe"].includes(tipoEst))
            perimetro = (an + al) * 2;
          if (perimetro > 0) defVal = (perimetro - 0.06 + 0.12).toFixed(2);
        } else if (tipoAcero === "longitudinal") {
          if (
            ["columna", "columna circular", "pilas", "pila", "dado"].includes(
              tipoEst,
            )
          )
            defVal = al > 0 ? al.toString() : "";
          else if (
            [
              "trabe",
              "nervadura",
              "contratrabe",
              "zapata aislada",
              "zapata corrida",
              "losa",
              "losa nervada",
              "escalera papelillo",
              "rampa de escalera",
            ].includes(tipoEst)
          )
            defVal = l > 0 ? l.toString() : "";
        } else if (
          tipoAcero === "transversal" &&
          [
            "zapata aislada",
            "zapata corrida",
            "losa",
            "losa nervada",
            "escalera papelillo",
            "rampa de escalera",
          ].includes(tipoEst)
        )
          defVal = an > 0 ? an.toString() : "";
        else if (tipoEst === "muro" || tipoEst === "muro curvo") {
          if (tipoAcero === "vertical") defVal = al > 0 ? al.toString() : "";
          else if (tipoAcero === "horizontal")
            defVal = l > 0 ? l.toString() : "";
        } else if (
          tipoAcero === "bastones" ||
          tipoAcero === "refuerzo adicional"
        )
          defVal = a.anchoCalc || "";
      } else if (field === "piezas") {
        const isPorPiezas =
          tipoAcero === "bastones" ||
          tipoAcero === "refuerzo adicional" ||
          (tipoAcero === "longitudinal" &&
            [
              "columna",
              "columna circular",
              "pilas",
              "pila",
              "dado",
              "trabe",
              "nervadura",
              "contratrabe",
            ].includes(tipoEst));
        if (!isPorPiezas) {
          const anchoC = parseFloat(a.anchoCalc),
            sep = parseFloat(a.separacion);
          if (!isNaN(anchoC) && anchoC > 0 && !isNaN(sep) && sep > 0)
            defVal = (Math.floor(anchoC / (sep / 100)) + 2).toString();
        }
      } else if (field === "traslapes") {
        if (
          tipoAcero === "estribos" ||
          tipoAcero === "zuncho" ||
          tipoAcero === "grapas"
        ) {
          defVal = "0";
        } else {
          const longCalc = parseFloat(a.longitud) || 0;
          defVal = longCalc >= 12 ? Math.floor(longCalc / 12).toString() : "0";
        }
      } else if (field === "anclaje") {
        if (
          tipoAcero === "estribos" ||
          tipoAcero === "zuncho" ||
          tipoAcero === "grapas"
        ) {
          defVal = "0";
        } else {
          defVal = ANCLAJE_DEFAULT[a.numVarilla] || "";
        }
      } else if (field === "ganchos") {
        if (
          tipoAcero === "estribos" ||
          tipoAcero === "zuncho" ||
          tipoAcero === "grapas"
        ) {
          defVal = "0";
        } else {
          defVal = "2";
        }
      } else if (field === "separacion") {
        if (!isPorPiezas) {
          defVal = "0.20"; // Standard default if none provided
        } else {
          defVal = "";
        }
      }
      if (defVal !== "" && defVal !== a[field])
        updateSteelItem(structId, itemId, field, defVal);
    },
    [estructuras, updateSteelItem],
  );

  const addSteelItem = useCallback(
    (structId) => {
      updateActiveEstructuras((prev) =>
        prev.map((s) => {
          if (s.id !== structId) return s;
          let tramoSugerido = "";
          const tipo = (s.tipo || "").toLowerCase();

          const validTypes = getSteelTypesForElement(s.tipo);
          const initialTipo = validTypes.length > 0 ? validTypes[0] : "";
          const tAcero = initialTipo.toLowerCase();

          if (
            ["columna", "columna circular", "pilas", "pila", "dado"].includes(
              tipo,
            ) &&
            parseFloat(s.alto) > 0
          )
            tramoSugerido = s.alto.toString();
          else if (
            [
              "trabe",
              "nervadura",
              "contratrabe",
              "muro",
              "muro curvo",
            ].includes(tipo) &&
            parseFloat(s.largo) > 0
          )
            tramoSugerido = s.largo.toString();

          let initialLongitud = "";
          let initialAnchoCalc = tramoSugerido;
          const al = parseFloat(s.alto) || 0,
            l = parseFloat(s.largo) || 0,
            an = parseFloat(s.ancho) || 0;

          if (tAcero === "longitudinal") {
            if (
              ["columna", "columna circular", "pilas", "pila", "dado"].includes(
                tipo,
              )
            )
              initialLongitud = al > 0 ? al.toString() : "";
            else if (["trabe", "nervadura", "contratrabe"].includes(tipo))
              initialLongitud = l > 0 ? l.toString() : "";
            else if (
              [
                "zapata aislada",
                "zapata corrida",
                "losa",
                "losa nervada",
                "escalera papelillo",
                "rampa de escalera",
              ].includes(tipo)
            ) {
              initialAnchoCalc = an > 0 ? an.toString() : "";
              initialLongitud = l > 0 ? l.toString() : "";
            }
          } else if (tAcero === "principal") {
            initialLongitud = l > 0 ? l.toString() : "";
            initialAnchoCalc = an > 0 ? an.toString() : "";
          } else if (
            tAcero === "vertical" &&
            ["muro", "muro curvo"].includes(tipo)
          ) {
            initialAnchoCalc = l > 0 ? l.toString() : "";
            initialLongitud = al > 0 ? al.toString() : "";
          }

          const initialTraslapes =
            initialLongitud && parseFloat(initialLongitud) >= 12
              ? Math.floor(parseFloat(initialLongitud) / 12).toString()
              : "0";

          return {
            ...s,
            aceros: [
              ...(s.aceros || []),
              {
                id: `AC-${Date.now()}`,
                tipo: initialTipo,
                numVarilla: "4",
                anchoCalc: initialAnchoCalc,
                separacion: "",
                piezas: "1",
                longitud: initialLongitud,
                ganchos: "",
                anclaje: ANCLAJE_DEFAULT["4"],
                traslapes: initialTraslapes,
              },
            ],
          };
        }),
      );
    },
    [updateActiveEstructuras],
  );
  const removeSteelItem = useCallback(
    (structId, itemId) =>
      updateActiveEstructuras((prev) =>
        prev.map((s) =>
          s.id === structId
            ? { ...s, aceros: (s.aceros || []).filter((a) => a.id !== itemId) }
            : s,
        ),
      ),
    [updateActiveEstructuras],
  );

  const renderSteelSubmodal = () => {
    if (!activeSteelSubmodal) return null;
    const structure = estructuras.find((e) => e.id === activeSteelSubmodal);
    if (!structure) return null;
    let aceros = structure.aceros || [];
    const tiposAceroOpciones = getSteelTypesForElement(structure.tipo);
    const validTypes = tiposAceroOpciones.map((t) => t.toLowerCase());
    if (validTypes.length > 0) {
      aceros = aceros.map((a) => {
        if (!validTypes.includes((a.tipo || "").toLowerCase())) {
          return { ...a, tipo: tiposAceroOpciones[0] };
        }
        return a;
      });
    }
    const ePiezas = parseFloat(structure.piezas) || 1;
    const resumenAcero = aceros.reduce((acc, a) => {
      const calc = calcAceroItem(a),
        num = a.numVarilla || "-";
      if (!acc[num]) acc[num] = { ml: 0, kg: 0 };
      acc[num].ml += calc.ml * ePiezas;
      acc[num].kg += Math.round(calc.kg * ePiezas * 100) / 100;
      return acc;
    }, {});

    return (
      <div className="fixed inset-0 bg-slate-900/80 backdrop-blur-sm z-[80] flex items-center justify-center p-4 sm:p-8 animate-in fade-in zoom-in-95 duration-200">
        <div className="bg-white rounded-3xl shadow-2xl w-full max-w-[1400px] overflow-hidden flex flex-col h-[90vh]">
          <div className="bg-slate-800 text-white p-6 flex justify-between items-center shrink-0">
            <div className="flex items-center gap-4">
              <div className="p-3 bg-slate-700 rounded-xl">
                <Wrench size={24} className="text-slate-300" />
              </div>
              <div>
                <h3 className="font-black uppercase tracking-widest text-lg flex items-center gap-3">
                  Despiece de Acero de Refuerzo
                </h3>
                <p className="text-xs text-slate-400 font-bold uppercase tracking-wider mt-1">
                  {structure.tipo || "Elemento"} • EJE: {structure.eje || "-"} •
                  CLAVE: {structure.clave || "-"}
                </p>
              </div>
            </div>
            <button
              onClick={() => setActiveSteelSubmodal(null)}
              className="hover:bg-white/20 p-2 rounded-full transition-colors"
            >
              <X size={24} />
            </button>
          </div>
          <div className="flex flex-col lg:flex-row flex-1 overflow-hidden">
            <div className="flex-1 p-4 sm:p-6 bg-slate-50 overflow-y-auto border-r border-slate-200">
              <div className="overflow-hidden rounded-xl border border-slate-200 shadow-sm bg-white pb-2">
                <table className="text-left border-collapse text-[10px] md:text-xs table-fixed w-full">
                  <thead className="bg-slate-100 text-slate-600 font-black uppercase tracking-tighter md:tracking-wider border-b-2 border-slate-200 leading-tight">
                    <tr>
                      <th
                        className="p-1 md:p-2 text-center border-r select-none"
                        style={{ width: steelColWidths.no }}
                      >
                        No.
                      </th>
                      <th
                        className="p-1 md:p-2 border-r select-none"
                        style={{ width: steelColWidths.tipo }}
                      >
                        Tipo Ref.
                      </th>
                      <th
                        className="p-1 md:p-2 border-r text-center select-none"
                        style={{ width: steelColWidths.varilla }}
                      >
                        # Var.
                      </th>
                      <th
                        className="p-1 md:p-2 border-r text-center bg-indigo-50/50 select-none"
                        style={{ width: steelColWidths.tramo }}
                      >
                        Tramo
                      </th>
                      <th
                        className="p-1 md:p-2 border-r text-center bg-indigo-50/50 select-none"
                        style={{ width: steelColWidths.separacion }}
                      >
                        Sep.
                      </th>
                      <th
                        className="p-1 md:p-2 border-r text-center bg-blue-50/50 select-none"
                        style={{ width: steelColWidths.piezas }}
                      >
                        Pzas
                      </th>
                      <th
                        className="p-1 md:p-2 border-r text-center bg-blue-50/50 select-none"
                        style={{ width: steelColWidths.longitud }}
                      >
                        Long.
                      </th>
                      <th
                        className="p-1 md:p-2 border-r text-center bg-emerald-50/50 select-none"
                        style={{ width: steelColWidths.ganchos }}
                      >
                        Gan.
                      </th>
                      <th
                        className="p-1 md:p-2 border-r text-center bg-emerald-50/50 select-none"
                        style={{ width: steelColWidths.anclaje }}
                      >
                        Ancl.
                      </th>
                      <th
                        className="p-1 md:p-2 border-r text-center bg-emerald-50/50 select-none"
                        style={{ width: steelColWidths.traslapes }}
                      >
                        Tras.
                      </th>
                      <th
                        className="p-1 md:p-2 border-r text-center bg-slate-100 select-none"
                        style={{ width: steelColWidths.mlPza }}
                      >
                        ML/Pza
                      </th>
                      <th
                        className="p-1 md:p-2 border-r text-right bg-slate-200 select-none"
                        style={{ width: steelColWidths.totalMl }}
                      >
                        Tot.(ml)
                      </th>
                      <th
                        className="p-1 md:p-2 border-r text-right bg-slate-300 select-none"
                        style={{ width: steelColWidths.totalKg }}
                      >
                        Tot.(kg)
                      </th>
                      <th
                        className="p-1 md:p-2 text-center select-none"
                        style={{ width: steelColWidths.action }}
                      ></th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {aceros.length === 0 && (
                      <tr>
                        <td
                          colSpan={13}
                          className="p-6 text-center text-slate-400 font-bold uppercase"
                        >
                          Sin despiece
                        </td>
                      </tr>
                    )}
                    {aceros.map((item, i) => {
                      const calc = calcAceroItem(item),
                        tipoMinuscula = (item.tipo || "").toLowerCase(),
                        isVinculadoLargo =
                          tipoMinuscula === "bastones" ||
                          tipoMinuscula === "refuerzo adicional",
                        isLongitudinalManual =
                          tipoMinuscula === "longitudinal" &&
                          [
                            "columna",
                            "columna circular",
                            "pilas",
                            "pila",
                            "dado",
                            "trabe",
                            "nervadura",
                            "contratrabe",
                          ].includes((structure.tipo || "").toLowerCase()),
                        isPorPiezas = isVinculadoLargo || isLongitudinalManual,
                        isEstriboGroup =
                          tipoMinuscula === "estribos" ||
                          tipoMinuscula === "grapas" ||
                          tipoMinuscula === "zuncho",
                        isTrabeRel = [
                          "trabe",
                          "nervadura",
                          "contratrabe",
                        ].includes((structure.tipo || "").toLowerCase()),
                        isLongOrBast =
                          tipoMinuscula === "longitudinal" ||
                          tipoMinuscula === "bastones",
                        showNAAnchoSep = isTrabeRel && isLongOrBast;
                      return (
                        <tr
                          key={item.id}
                          className="hover:bg-slate-50 transition-colors"
                        >
                          <td className="p-1 border-r text-center font-black text-slate-400">
                            {i + 1}
                          </td>
                          <td className="p-0 border-r">
                            <select
                              className="w-full p-1 bg-transparent outline-none font-bold text-slate-700 uppercase cursor-pointer"
                              value={item.tipo || ""}
                              onChange={(ev) =>
                                updateSteelItem(
                                  structure.id,
                                  item.id,
                                  "tipo",
                                  ev.target.value,
                                )
                              }
                            >
                              <option
                                value=""
                                disabled
                                className="text-slate-300"
                              >
                                Sel...
                              </option>
                              {tiposAceroOpciones.map((t) => (
                                <option key={t} value={t}>
                                  {t}
                                </option>
                              ))}
                            </select>
                          </td>
                          <td className="p-0 border-r">
                            <select
                              className="w-full p-1 bg-transparent outline-none font-black text-slate-800 text-center cursor-pointer"
                              value={item.numVarilla || "4"}
                              onChange={(ev) =>
                                updateSteelItem(
                                  structure.id,
                                  item.id,
                                  "numVarilla",
                                  ev.target.value,
                                )
                              }
                            >
                              {Object.keys(PESOS_VARILLA).map((n) => (
                                <option key={n} value={n}>
                                  #{n}
                                </option>
                              ))}
                            </select>
                          </td>
                          <td className="p-0 border-r bg-indigo-50/30">
                            <DebouncedCell
                              type="number"
                              step="0.01"
                              value={
                                showNAAnchoSep || isLongitudinalManual
                                  ? ""
                                  : item.anchoCalc
                              }
                              disabled={showNAAnchoSep || isLongitudinalManual}
                              onChange={(v) =>
                                updateSteelItem(
                                  structure.id,
                                  item.id,
                                  "anchoCalc",
                                  v,
                                )
                              }
                              onDoubleClick={() =>
                                resetSteelFieldToDefault(
                                  structure.id,
                                  item.id,
                                  "anchoCalc",
                                )
                              }
                              className={`w-full p-1 bg-transparent outline-none text-center font-black ${isLongitudinalManual || showNAAnchoSep ? "text-indigo-400 italic" : "text-indigo-900"}`}
                              placeholder={
                                isLongitudinalManual || showNAAnchoSep
                                  ? "N/A"
                                  : "0.00"
                              }
                            />
                          </td>
                          <td className="p-0 border-r bg-indigo-50/30">
                            <DebouncedCell
                              type="number"
                              step="0.01"
                              value={
                                showNAAnchoSep || isPorPiezas
                                  ? ""
                                  : item.separacion
                              }
                              disabled={showNAAnchoSep || isPorPiezas}
                              onChange={(v) =>
                                updateSteelItem(
                                  structure.id,
                                  item.id,
                                  "separacion",
                                  v,
                                )
                              }
                              onDoubleClick={() =>
                                resetSteelFieldToDefault(
                                  structure.id,
                                  item.id,
                                  "separacion",
                                )
                              }
                              className={`w-full p-1 bg-transparent outline-none text-center font-black ${isPorPiezas || showNAAnchoSep ? "text-slate-400 italic" : "text-indigo-900"}`}
                              placeholder={
                                isPorPiezas || showNAAnchoSep ? "N/A" : "-"
                              }
                            />
                          </td>
                          <td className="p-0 border-r bg-blue-50/30">
                            <DebouncedCell
                              type="number"
                              min="1"
                              value={item.piezas}
                              onChange={(v) =>
                                updateSteelItem(
                                  structure.id,
                                  item.id,
                                  "piezas",
                                  v,
                                )
                              }
                              onDoubleClick={() =>
                                resetSteelFieldToDefault(
                                  structure.id,
                                  item.id,
                                  "piezas",
                                )
                              }
                              className="w-full p-1 bg-transparent outline-none text-center font-black text-blue-900"
                              placeholder="1"
                            />
                          </td>
                          <td className="p-0 border-r bg-blue-50/30">
                            <DebouncedCell
                              type="number"
                              step="0.01"
                              value={item.longitud}
                              onChange={(v) =>
                                updateSteelItem(
                                  structure.id,
                                  item.id,
                                  "longitud",
                                  v,
                                )
                              }
                              onDoubleClick={() =>
                                resetSteelFieldToDefault(
                                  structure.id,
                                  item.id,
                                  "longitud",
                                )
                              }
                              className={`w-full p-1 bg-transparent outline-none text-center font-black ${isLongitudinalManual || isVinculadoLargo ? "text-blue-600 italic" : "text-blue-800"}`}
                              placeholder="0.00"
                            />
                          </td>
                          <td
                            className={`p-0 border-r ${isEstriboGroup ? "bg-slate-100 cursor-not-allowed" : "bg-emerald-50/30"}`}
                          >
                            <DebouncedCell
                              type="number"
                              step="1"
                              value={isEstriboGroup ? "" : item.ganchos}
                              disabled={isEstriboGroup}
                              onChange={(v) =>
                                updateSteelItem(
                                  structure.id,
                                  item.id,
                                  "ganchos",
                                  v,
                                )
                              }
                              onDoubleClick={() =>
                                resetSteelFieldToDefault(
                                  structure.id,
                                  item.id,
                                  "ganchos",
                                )
                              }
                              className={`w-full p-1 bg-transparent outline-none text-center font-black ${isEstriboGroup ? "text-slate-300 italic" : "text-emerald-800"}`}
                              placeholder={isEstriboGroup ? "N/A" : "0"}
                            />
                          </td>
                          <td
                            className={`p-0 border-r ${isEstriboGroup ? "bg-slate-100 cursor-not-allowed" : "bg-emerald-50/30"}`}
                          >
                            <DebouncedCell
                              type="number"
                              step="0.01"
                              value={isEstriboGroup ? "" : item.anclaje}
                              disabled={isEstriboGroup}
                              onChange={(v) =>
                                updateSteelItem(
                                  structure.id,
                                  item.id,
                                  "anclaje",
                                  v,
                                )
                              }
                              onDoubleClick={() =>
                                resetSteelFieldToDefault(
                                  structure.id,
                                  item.id,
                                  "anclaje",
                                )
                              }
                              className={`w-full p-1 bg-transparent outline-none text-center font-black ${isEstriboGroup ? "text-slate-300 italic" : "text-emerald-800"}`}
                              placeholder={isEstriboGroup ? "N/A" : "0.00"}
                            />
                          </td>
                          <td
                            className={`p-0 border-r ${isEstriboGroup ? "bg-slate-100 cursor-not-allowed" : "bg-emerald-50/30"}`}
                          >
                            <DebouncedCell
                              type="number"
                              step="0.01"
                              value={isEstriboGroup ? "" : item.traslapes}
                              disabled={isEstriboGroup}
                              onChange={(v) =>
                                updateSteelItem(
                                  structure.id,
                                  item.id,
                                  "traslapes",
                                  v,
                                )
                              }
                              onDoubleClick={() =>
                                resetSteelFieldToDefault(
                                  structure.id,
                                  item.id,
                                  "traslapes",
                                )
                              }
                              className={`w-full p-1 bg-transparent outline-none text-center font-black ${isEstriboGroup ? "text-slate-300 italic" : "text-emerald-800"}`}
                              placeholder={isEstriboGroup ? "N/A" : "0.00"}
                            />
                          </td>
                          <td className="p-1 border-r text-center font-black text-slate-500 bg-slate-50">
                            {calc.mlPorPieza > 0
                              ? calc.mlPorPieza.toFixed(2)
                              : "-"}
                          </td>
                          <td className="p-1 border-r text-right font-black text-slate-600 bg-slate-100/50">
                            {calc.ml > 0 ? calc.ml.toFixed(2) : "-"}
                          </td>
                          <td className="p-1 border-r text-right font-black text-blue-700 bg-slate-200/50">
                            {calc.kg > 0 ? calc.kg.toFixed(2) : "-"}
                          </td>
                          <td className="p-1 text-center">
                            <button
                              onClick={() =>
                                removeSteelItem(structure.id, item.id)
                              }
                              className="text-slate-300 hover:text-red-500 transition-colors p-1"
                            >
                              <Trash2 size={16} />
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                <button
                  onClick={() => addSteelItem(structure.id)}
                  className="m-4 px-6 py-2 border-2 border-dashed border-slate-300 text-slate-500 font-black rounded-xl hover:bg-slate-200 text-xs"
                >
                  + Agregar Refuerzo
                </button>
              </div>
              <div className="p-6 bg-slate-100 border-t border-slate-200 flex flex-col gap-4">
                {Object.keys(resumenAcero).length > 0 && (
                  <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                    {Object.entries(resumenAcero).map(([num, data]) => (
                      <div
                        key={num}
                        className="bg-white border border-slate-200 rounded-xl p-4 shadow-sm"
                      >
                        <div className="flex justify-between items-end mb-2 border-b border-slate-100 pb-2">
                          <span className="font-black text-lg text-slate-800">
                            #{num}
                          </span>
                          <div className="flex flex-col items-end">
                            <span className="text-[10px] font-bold text-slate-400 uppercase">
                              ({(PESOS_VARILLA[num] || 0).toFixed(3)} kg/m)
                            </span>
                            <span className="text-[8px] font-black text-indigo-500 bg-indigo-50 px-1.5 py-0.5 rounded mt-0.5">
                              X {ePiezas} PZAS
                            </span>
                          </div>
                        </div>
                        <div className="flex justify-between items-center text-sm mb-1">
                          <span className="font-bold text-slate-500 uppercase text-[10px]">
                            Total ML:
                          </span>
                          <span className="font-black text-slate-700">
                            {data.ml.toLocaleString(undefined, {
                              minimumFractionDigits: 2,
                            })}{" "}
                            ml
                          </span>
                        </div>
                        <div className="flex justify-between items-center">
                          <span className="font-bold text-slate-500 uppercase text-[10px]">
                            Total KG:
                          </span>
                          <span className="font-black text-blue-700 text-lg">
                            {data.kg.toLocaleString(undefined, {
                              minimumFractionDigits: 2,
                            })}{" "}
                            kg
                          </span>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
                <div className="flex flex-col md:flex-row justify-between items-stretch md:items-center mt-4 gap-4 bg-slate-200/50 p-2 rounded-3xl border border-slate-200">
                  <div className="flex items-center gap-2 overflow-x-auto pb-2 md:pb-0">
                    <div className="bg-white p-3 rounded-2xl shadow-sm border border-slate-200 inline-flex flex-col items-start min-w-[140px]">
                      <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest mb-1">
                        Peso (1 Pza)
                      </span>
                      <div className="text-xl font-black text-slate-800 flex items-baseline gap-1">
                        {calcAceroTotalKg(aceros, 1).toLocaleString(undefined, {
                          minimumFractionDigits: 2,
                          maximumFractionDigits: 2,
                        })}{" "}
                        <span className="text-[10px] text-slate-400">kg</span>
                      </div>
                    </div>
                    <div className="text-slate-400 font-black text-lg px-1">
                      ×
                    </div>
                    <div className="bg-indigo-50 p-3 rounded-2xl border border-indigo-100 shadow-sm inline-flex flex-col items-start min-w-[120px]">
                      <span className="text-[10px] font-black text-indigo-600 uppercase tracking-widest mb-1">
                        Pzas (Est.)
                      </span>
                      <div className="text-xl font-black text-indigo-900">
                        {ePiezas}
                      </div>
                    </div>
                    <div className="text-slate-400 font-black text-lg px-1">
                      =
                    </div>
                    <div className="bg-slate-800 p-3 rounded-2xl text-white shadow-lg inline-flex flex-col items-start min-w-[160px]">
                      <span className="text-[10px] font-black text-blue-300 uppercase tracking-widest mb-1">
                        Gran Total
                      </span>
                      <div className="text-2xl font-black flex items-baseline gap-1">
                        {calcAceroTotalKg(aceros, ePiezas).toLocaleString(
                          undefined,
                          {
                            minimumFractionDigits: 2,
                            maximumFractionDigits: 2,
                          },
                        )}{" "}
                        <span className="text-xs text-slate-400">kg</span>
                      </div>
                    </div>
                  </div>
                  <button
                    onClick={() => setActiveSteelSubmodal(null)}
                    className="px-8 py-4 bg-blue-600 text-white rounded-2xl font-black text-xs uppercase shadow-md hover:bg-blue-500 transition-transform active:scale-95 shrink-0 h-full"
                  >
                    Cerrar Despiece
                  </button>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  };

  const renderWallSubmodal = () => {
    if (!activeWallSubmodal) return null;
    const wall = muros.find((m) => m.id === activeWallSubmodal.wallId);
    if (!wall) return null;
    const isHueco = activeWallSubmodal.type === "huecos";
    const items = wall[activeWallSubmodal.type] || [];
    return (
      <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[70] flex items-center justify-center p-4 animate-in fade-in zoom-in-95 duration-200">
        <div className="bg-white rounded-2xl shadow-2xl w-full max-w-3xl overflow-hidden flex flex-col">
          <div
            className={`p-4 flex justify-between items-center text-white ${isHueco ? "bg-red-800" : "bg-orange-800"}`}
          >
            <h3 className="font-black uppercase tracking-widest text-sm flex items-center gap-3">
              {isHueco ? "Deducción por Huecos" : "Deducción por Columnas"}{" "}
              <span className="bg-black/20 px-2 py-1 rounded">
                Eje {wall.eje || "-"}
              </span>
            </h3>
            <button
              onClick={() => setActiveWallSubmodal(null)}
              className="hover:bg-white/20 p-1 rounded transition-colors"
            >
              <X size={20} />
            </button>
          </div>
          <div className="p-6 bg-slate-50 flex-1 overflow-auto max-h-[60vh]">
            <table className="w-full text-left border-collapse text-xs bg-white rounded-lg overflow-hidden shadow-sm">
              <thead className="bg-slate-200 text-slate-600 font-bold uppercase">
                <tr>
                  <th className="p-3 border-b text-center w-12">No.</th>
                  {isHueco && <th className="p-3 border-b">Tipo</th>}
                  <th className="p-3 border-b text-center">
                    {isHueco ? "Largo (m)" : "Ancho (m)"}
                  </th>
                  <th className="p-3 border-b text-center">Alto (m)</th>
                  <th className="p-3 border-b text-center">Pzas</th>
                  <th className="p-3 border-b text-right bg-slate-300">
                    Deducción (m2)
                  </th>
                  <th className="p-3 border-b text-center w-12"></th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {items.length === 0 && (
                  <tr>
                    <td
                      colSpan={7}
                      className="p-8 text-center text-slate-400 font-bold uppercase tracking-widest"
                    >
                      Sin capturas en este rubro
                    </td>
                  </tr>
                )}
                {items.map((item, i) => {
                  const ded = isHueco
                    ? (parseFloat(item.largo) || 0) *
                      (parseFloat(item.alto) || 0) *
                      (parseFloat(item.piezas) || 1)
                    : (parseFloat(item.ancho) || 0) *
                      (parseFloat(wall.alto) || 0) *
                      (parseFloat(item.piezas) || 1);
                  return (
                    <tr
                      key={item.id}
                      className="hover:bg-blue-50 transition-colors"
                    >
                      <td className="p-3 border-r text-center font-black text-slate-400">
                        {i + 1}
                      </td>
                      {isHueco && (
                        <td className="p-1 border-r">
                          <DebouncedCell
                            value={item.tipo}
                            onChange={(v) =>
                              updateWallSubItem(
                                wall.id,
                                activeWallSubmodal.type,
                                item.id,
                                "tipo",
                                v,
                              )
                            }
                            className="w-full p-2 bg-transparent outline-none font-bold text-slate-700 uppercase"
                            placeholder="Ej. Puerta Principal"
                          />
                        </td>
                      )}
                      <td className="p-1 border-r">
                        <DebouncedCell
                          value={isHueco ? item.largo : item.ancho}
                          onChange={(v) =>
                            updateWallSubItem(
                              wall.id,
                              activeWallSubmodal.type,
                              item.id,
                              isHueco ? "largo" : "ancho",
                              v,
                            )
                          }
                          className="w-full p-2 bg-transparent outline-none text-center font-bold text-blue-800"
                          placeholder="0.00"
                        />
                      </td>
                      <td className="p-3 border-r text-center bg-slate-50">
                        {isHueco ? (
                          <DebouncedCell
                            value={item.alto}
                            onChange={(v) =>
                              updateWallSubItem(
                                wall.id,
                                activeWallSubmodal.type,
                                item.id,
                                "alto",
                                v,
                              )
                            }
                            className="w-full bg-transparent outline-none text-center font-bold text-blue-800"
                            placeholder="0.00"
                          />
                        ) : (
                          <span
                            className="font-bold text-slate-500 bg-slate-200 px-3 py-1 rounded cursor-not-allowed"
                            title="Heredado del Alto de Muro"
                          >
                            {parseFloat(wall.alto || 0).toFixed(2)}
                          </span>
                        )}
                      </td>
                      <td className="p-1 border-r">
                        <DebouncedCell
                          value={item.piezas}
                          onChange={(v) =>
                            updateWallSubItem(
                              wall.id,
                              activeWallSubmodal.type,
                              item.id,
                              "piezas",
                              v,
                            )
                          }
                          className="w-full p-2 bg-transparent outline-none text-center font-black text-lg text-slate-700"
                          placeholder="1"
                        />
                      </td>
                      <td className="p-3 border-r text-right font-black text-red-600 bg-red-50/50">
                        {ded.toFixed(2)}
                      </td>
                      <td className="p-3 text-center">
                        <button
                          onClick={() =>
                            removeWallSubItem(
                              wall.id,
                              activeWallSubmodal.type,
                              item.id,
                            )
                          }
                          className="text-slate-300 hover:text-red-500 transition-colors"
                        >
                          <Trash2 size={18} />
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            <button
              onClick={() => addWallSubItem(wall.id, activeWallSubmodal.type)}
              className="mt-4 px-6 py-3 border-2 border-dashed border-slate-300 text-slate-500 font-black rounded-xl hover:bg-slate-200 text-[10px] uppercase tracking-widest w-full"
            >
              + Añadir Fila de Deducción
            </button>
          </div>
          <div className="bg-slate-100 p-5 flex justify-end border-t border-slate-200">
            <button
              onClick={() => setActiveWallSubmodal(null)}
              className="px-8 py-3 bg-slate-800 text-white rounded-xl font-black text-[11px] uppercase tracking-[0.2em] shadow-lg hover:bg-slate-700 hover:scale-95 transition-transform"
            >
              Guardar y Cerrar
            </button>
          </div>
        </div>
      </div>
    );
  };

  const renderCasetonSubmodal = () => {
    if (!activeCasetonSubmodal) return null;
    const structure = estructuras.find((e) => e.id === activeCasetonSubmodal);
    if (!structure) return null;
    const items = structure.casetones || [];
    const ePiezas = parseFloat(structure.piezas) || 1;
    return (
      <div className="fixed inset-0 bg-slate-900/80 backdrop-blur-sm z-[80] flex items-center justify-center p-4 sm:p-8 animate-in fade-in zoom-in-95 duration-200">
        <div className="bg-white rounded-3xl shadow-2xl w-full max-w-4xl overflow-hidden flex flex-col h-[80vh]">
          <div className="bg-sky-800 text-white p-6 flex justify-between items-center shrink-0">
            <div className="flex items-center gap-4">
              <div className="p-3 bg-sky-700 rounded-xl">
                <Box size={24} className="text-sky-200" />
              </div>
              <div>
                <h3 className="font-black uppercase tracking-widest text-lg flex items-center gap-3">
                  Desglose de Casetones
                </h3>
                <p className="text-xs text-sky-300 font-bold uppercase tracking-wider mt-1">
                  {structure.tipo || "Elemento"} • EJE: {structure.eje || "-"} •
                  CLAVE: {structure.clave || "-"}
                </p>
              </div>
            </div>
            <button
              onClick={() => setActiveCasetonSubmodal(null)}
              className="hover:bg-white/20 p-2 rounded-full transition-colors"
            >
              <X size={24} />
            </button>
          </div>
          <div className="p-6 bg-slate-50 flex-1 overflow-auto">
            <table className="w-full text-left border-collapse text-xs bg-white rounded-xl overflow-hidden shadow-sm">
              <thead className="bg-slate-200 text-slate-600 font-bold uppercase tracking-wider">
                <tr>
                  <th className="p-3 border-b text-center w-12">No.</th>
                  <th className="p-3 border-b">Clave</th>
                  <th className="p-3 border-b text-center">Largo (m)</th>
                  <th className="p-3 border-b text-center">Ancho (m)</th>
                  <th className="p-3 border-b text-center">Alto (m)</th>
                  <th className="p-3 border-b text-center">Pzas</th>
                  <th className="p-3 border-b text-right bg-slate-300">
                    Vol. (m3)
                  </th>
                  <th className="p-3 border-b text-center w-12"></th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {items.length === 0 && (
                  <tr>
                    <td
                      colSpan={8}
                      className="p-8 text-center text-slate-400 font-bold uppercase tracking-widest"
                    >
                      Sin casetones capturados
                    </td>
                  </tr>
                )}
                {items.map((item, i) => {
                  const vol =
                    (parseFloat(item.largo) || 0) *
                    (parseFloat(item.ancho) || 0) *
                    (parseFloat(item.alto) || 0) *
                    (parseFloat(item.piezas) || 1);
                  return (
                    <tr
                      key={item.id}
                      className="hover:bg-sky-50 transition-colors"
                    >
                      <td className="p-3 border-r text-center font-black text-slate-400">
                        {i + 1}
                      </td>
                      <td className="p-1 border-r">
                        <DebouncedCell
                          value={item.clave}
                          onChange={(v) =>
                            updateCasetonItem(structure.id, item.id, "clave", v)
                          }
                          className="w-full p-2 bg-transparent outline-none font-bold text-slate-700 uppercase"
                          placeholder="Ej. CAS-01"
                        />
                      </td>
                      <td className="p-1 border-r">
                        <DebouncedCell
                          type="number"
                          step="0.01"
                          value={item.largo}
                          onChange={(v) =>
                            updateCasetonItem(structure.id, item.id, "largo", v)
                          }
                          className="w-full p-2 bg-transparent outline-none text-center font-bold text-sky-800"
                          placeholder="0.00"
                        />
                      </td>
                      <td className="p-1 border-r">
                        <DebouncedCell
                          type="number"
                          step="0.01"
                          value={item.ancho}
                          onChange={(v) =>
                            updateCasetonItem(structure.id, item.id, "ancho", v)
                          }
                          className="w-full p-2 bg-transparent outline-none text-center font-bold text-sky-800"
                          placeholder="0.00"
                        />
                      </td>
                      <td className="p-1 border-r">
                        <DebouncedCell
                          type="number"
                          step="0.01"
                          value={item.alto}
                          onChange={(v) =>
                            updateCasetonItem(structure.id, item.id, "alto", v)
                          }
                          className="w-full p-2 bg-transparent outline-none text-center font-bold text-sky-800"
                          placeholder="0.00"
                        />
                      </td>
                      <td className="p-1 border-r">
                        <DebouncedCell
                          type="number"
                          value={item.piezas}
                          onChange={(v) =>
                            updateCasetonItem(
                              structure.id,
                              item.id,
                              "piezas",
                              v,
                            )
                          }
                          className="w-full p-2 bg-transparent outline-none text-center font-black text-lg text-slate-700"
                          placeholder="1"
                        />
                      </td>
                      <td className="p-3 border-r text-right font-black text-sky-700 bg-sky-50/50">
                        {(vol * ePiezas).toFixed(2)}
                      </td>
                      <td className="p-3 text-center">
                        <button
                          onClick={() =>
                            removeCasetonItem(structure.id, item.id)
                          }
                          className="text-slate-300 hover:text-red-500 transition-colors"
                        >
                          <Trash2 size={18} />
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            <button
              onClick={() => addCasetonItem(structure.id)}
              className="mt-4 px-6 py-3 border-2 border-dashed border-sky-300 text-sky-600 font-black rounded-xl hover:bg-sky-50 text-[10px] uppercase tracking-widest w-full"
            >
              + Añadir Casetón
            </button>
          </div>
          <div className="bg-slate-100 p-5 flex justify-between items-center border-t border-slate-200">
            <div className="flex items-baseline gap-2">
              <span className="text-[10px] font-black uppercase text-slate-500 tracking-widest">
                Volumen Total Casetones:
              </span>{" "}
              <span className="text-2xl font-black text-sky-800">
                {getCasetonesTotalVol(items, ePiezas).toFixed(2)}{" "}
                <span className="text-xs text-sky-600">m3</span>
              </span>
            </div>
            <button
              onClick={() => setActiveCasetonSubmodal(null)}
              className="px-8 py-3 bg-sky-800 text-white rounded-xl font-black text-[11px] uppercase tracking-[0.2em] shadow-lg hover:bg-sky-700 hover:scale-95 transition-transform"
            >
              Guardar y Cerrar
            </button>
          </div>
        </div>
      </div>
    );
  };

  const renderConceptSelectorModal = () => {
    if (!conceptSelectorFor) return null;
    const filtered = catalogoConceptos.filter((c) =>
      c.partidaId !== null &&
      c.descripcion && 
      c.descripcion.trim() !== "" &&
      (c.id + " " + (c.descripcion || ""))
        .toLowerCase()
        .includes(conceptSearchTerm.toLowerCase()),
    );

    return (
      <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[300] flex items-center justify-center p-4 md:p-8 animate-in fade-in zoom-in-95 duration-200">
        <div className="bg-white rounded-3xl shadow-2xl w-full max-w-4xl flex flex-col h-[70vh] overflow-hidden">
          <div className="p-6 bg-blue-800 text-white flex justify-between items-center shrink-0">
            <div>
              <h3 className="font-black uppercase tracking-widest text-lg">
                Buscador del Catálogo
              </h3>
              <p className="text-xs text-blue-200 mt-1">
                Selecciona un concepto para auto-rellenar la fila
              </p>
            </div>
            <button
              onClick={() => setConceptSelectorFor(null)}
              className="p-2 hover:bg-white/20 rounded-full transition-colors"
            >
              <X size={24} />
            </button>
          </div>
          <div className="p-4 bg-slate-50 border-b border-slate-200 shrink-0">
            <div className="relative">
              <Search
                size={18}
                className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400"
              />
              <input
                type="text"
                placeholder="Buscar por código o descripción..."
                className="w-full p-3 pl-12 border border-slate-300 rounded-xl text-sm font-bold shadow-sm outline-none focus:ring-2 focus:ring-blue-500"
                value={conceptSearchTerm}
                onChange={(e) => setConceptSearchTerm(e.target.value)}
                autoFocus
              />
            </div>
          </div>
          <div className="flex-1 overflow-auto p-4 bg-slate-50">
            <table className="w-full text-xs text-left border-collapse bg-white rounded-xl shadow-sm overflow-hidden">
              <thead className="bg-slate-200 text-slate-600 uppercase font-black">
                <tr>
                  <th className="p-3 border-b">Código</th>
                  <th className="p-3 border-b">Descripción</th>
                  <th className="p-3 border-b text-center">Unidad</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filtered.length === 0 && (
                  <tr>
                    <td
                      colSpan={3}
                      className="p-8 text-center text-slate-400 font-bold uppercase"
                    >
                      No se encontraron conceptos
                    </td>
                  </tr>
                )}
                {filtered.map((cat) => (
                  <tr
                    key={cat.id}
                    onClick={() => {
                      updateConceptoId(conceptSelectorFor, cat.id);
                      setConceptSelectorFor(null);
                    }}
                    className="hover:bg-blue-50 cursor-pointer transition-colors group"
                  >
                    <td className="p-3 font-black text-blue-700 w-[20%]">
                      {cat.id}
                    </td>
                    <td className="p-3 text-slate-600 group-hover:text-blue-900 w-[60%] font-bold uppercase">
                      {cat.descripcion || "-"}
                    </td>
                    <td className="p-3 text-center font-bold text-slate-500 w-[20%] uppercase">
                      {cat.unidad || "-"}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
  };

  const renderGeneratorModal = () => {
    if (!editingModal) return null;
    const { concepto, nivel } = editingModal;
    const currentLevel = generadores[concepto.id]?.[nivel.id] || {
      rows: [],
      image: null,
    };
    const rows = currentLevel.rows || [];

    return (
      <div className="fixed inset-0 bg-slate-900/80 backdrop-blur-sm z-[100] flex items-center justify-center p-4 sm:p-8 animate-in fade-in zoom-in-95 duration-200">
        <div className="bg-white rounded-3xl shadow-2xl w-full max-w-[1400px] flex flex-col h-[90vh] overflow-hidden">
          <div className="bg-slate-800 text-white p-6 flex justify-between items-center shrink-0">
            <div>
              <h3 className="font-black uppercase tracking-widest text-lg">
                {concepto.clave} - {concepto.descripcion}
              </h3>
              <p className="text-xs text-slate-400 font-bold uppercase mt-1">
                Generador de volúmenes • Nivel: {nivel.nombre}
              </p>
            </div>
            <button
              onClick={() => {
                setEditingModal(null);
                setSelectedGeneratorRows([]);
              }}
              className="p-2 hover:bg-white/20 rounded-full transition-colors"
            >
              <X size={24} />
            </button>
          </div>
          <div className="flex flex-col lg:flex-row flex-1 overflow-hidden bg-slate-50">
            <div className="flex-1 p-4 overflow-y-auto flex flex-col border-r border-slate-200">
              <div className="flex flex-wrap gap-2 mb-4">
                <button
                  onClick={() => {
                    updateActiveGeneradores((prev) => {
                      const c = prev[concepto.id] || {};
                      const n = c[nivel.id] || { rows: [] };
                      return {
                        ...prev,
                        [concepto.id]: {
                          ...c,
                          [nivel.id]: {
                            ...n,
                            rows: [
                              ...n.rows,
                              {
                                id: `R-${Date.now()}`,
                                eje: "",
                                claveLoc: "",
                                largo: "",
                                ancho: "",
                                alto: "",
                                piezas: "1",
                              },
                            ],
                          },
                        },
                      };
                    });
                  }}
                  className="px-4 py-2 bg-blue-600 text-white font-black text-[10px] rounded-lg shadow-sm hover:bg-blue-500 uppercase flex items-center gap-2"
                >
                  <Plus size={14} /> Fila
                </button>
                <button
                  onClick={handleCopyRows}
                  className="px-4 py-2 bg-slate-200 text-slate-700 font-black text-[10px] rounded-lg hover:bg-slate-300 uppercase flex items-center gap-2"
                >
                  <Copy size={14} /> Copiar{" "}
                  {selectedGeneratorRows.length > 0
                    ? `(${selectedGeneratorRows.length})`
                    : "Todas"}
                </button>
                <button
                  onClick={handlePasteRows}
                  disabled={!clipboardRows || clipboardRows.length === 0}
                  className="px-4 py-2 bg-emerald-100 text-emerald-700 font-black text-[10px] rounded-lg hover:bg-emerald-200 uppercase flex items-center gap-2 disabled:opacity-50"
                >
                  <ClipboardPaste size={14} /> Pegar
                </button>
                <input
                  type="text"
                  placeholder="Notas internas..."
                  value={generadores[concepto.id]?.[nivel.id]?.comentario || ""}
                  onChange={(e) => {
                    updateActiveGeneradores((prev) => {
                      const c = prev[concepto.id] || {};
                      const n = c[nivel.id] || { rows: [] };
                      return {
                        ...prev,
                        [concepto.id]: {
                          ...c,
                          [nivel.id]: {
                            ...n,
                            comentario: e.target.value,
                          },
                        },
                      };
                    });
                  }}
                  className="flex-1 min-w-[200px] max-w-sm px-3 py-1.5 text-xs border border-slate-300 rounded focus:border-blue-500 outline-none"
                />
                {selectedGeneratorRows.length > 0 && (
                  <button
                    onClick={() => {
                      updateActiveGeneradores((prev) => {
                        const c = prev[concepto.id] || {};
                        const n = c[nivel.id];
                        if (!n) return prev;
                        return {
                          ...prev,
                          [concepto.id]: {
                            ...c,
                            [nivel.id]: {
                              ...n,
                              rows: n.rows.filter(
                                (r) => !selectedGeneratorRows.includes(r.id),
                              ),
                            },
                          },
                        };
                      });
                      setSelectedGeneratorRows([]);
                    }}
                    className="px-4 py-2 bg-red-100 text-red-700 font-black text-[10px] rounded-lg hover:bg-red-200 uppercase flex items-center gap-2"
                  >
                    <Trash2 size={14} /> Borrar ({selectedGeneratorRows.length})
                  </button>
                )}
              </div>
              <div className="overflow-x-auto bg-white rounded-xl shadow-sm border border-slate-200">
                <table className="w-full text-left border-collapse text-[10px] md:text-xs">
                  <thead className="bg-slate-100 text-slate-500 font-black uppercase border-b-2 border-slate-200 tracking-wider">
                    <tr>
                      <th
                        className="p-2 text-center border-r"
                        style={{ width: genColWidths.select }}
                      >
                        <button onClick={toggleAllGeneratorRows}>
                          {rows.length > 0 &&
                          selectedGeneratorRows.length === rows.length ? (
                            <CheckSquare size={14} />
                          ) : (
                            <Square size={14} />
                          )}
                        </button>
                      </th>
                      <th
                        className="p-2 border-r"
                        style={{ width: genColWidths.localizacion }}
                      >
                        Localización (Eje / Clave)
                      </th>
                      <th
                        className="p-2 text-center border-r text-blue-800 bg-blue-50"
                        style={{ width: genColWidths.largo }}
                      >
                        Largo
                      </th>
                      <th
                        className="p-2 text-center border-r text-blue-800 bg-blue-50"
                        style={{ width: genColWidths.ancho }}
                      >
                        Ancho
                      </th>
                      <th
                        className="p-2 text-center border-r text-blue-800 bg-blue-50"
                        style={{ width: genColWidths.kg_ml }}
                      >
                        kg/ml
                      </th>
                      <th
                        className="p-2 text-center border-r text-blue-800 bg-blue-50"
                        style={{ width: genColWidths.alto }}
                      >
                        Alto
                      </th>
                      <th
                        className="p-2 text-center border-r text-blue-800 bg-blue-50"
                        style={{ width: genColWidths.volPza }}
                      >
                        Vol/Pza
                      </th>
                      <th
                        className="p-2 text-center border-r text-blue-800 bg-blue-50"
                        style={{ width: genColWidths.piezas }}
                      >
                        Pzas
                      </th>
                      <th
                        className="p-2 text-right border-r bg-slate-200"
                        style={{ width: genColWidths.volumen }}
                      >
                        Total
                      </th>
                      <th
                        className="p-2 text-center"
                        style={{ width: genColWidths.action }}
                      >
                        <div className="flex justify-center items-center gap-1">
                          <button
                            onClick={handleCopyRows}
                            className={`p-1 rounded transition-colors ${selectedGeneratorRows.length > 0 ? "text-blue-600 hover:bg-blue-100" : "text-slate-300 cursor-not-allowed"}`}
                            disabled={selectedGeneratorRows.length === 0}
                            title="Copiar Seleccionados"
                          >
                            <Copy size={14} />
                          </button>
                          <button
                            onClick={() => {
                              if (selectedGeneratorRows.length > 0) {
                                updateActiveGeneradores((prev) => {
                                  const currentConcept =
                                    prev[editingModal.concepto.id] || {};
                                  const currentLevel =
                                    currentConcept[editingModal.nivel.id];
                                  if (!currentLevel) return prev;

                                  return {
                                    ...prev,
                                    [editingModal.concepto.id]: {
                                      ...currentConcept,
                                      [editingModal.nivel.id]: {
                                        ...currentLevel,
                                        rows: currentLevel.rows.filter(
                                          (r) =>
                                            !selectedGeneratorRows.includes(
                                              r.id,
                                            ),
                                        ),
                                      },
                                    },
                                  };
                                });
                                setSelectedGeneratorRows([]);
                              }
                            }}
                            className={`p-1 rounded transition-colors ${selectedGeneratorRows.length > 0 ? "text-red-500 hover:bg-red-50" : "text-slate-300 cursor-not-allowed"}`}
                            disabled={selectedGeneratorRows.length === 0}
                            title="Borrar Seleccionados"
                          >
                            <Trash2 size={14} />
                          </button>
                        </div>
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {rows.length === 0 && (
                      <tr>
                        <td
                          colSpan={8}
                          className="p-8 text-center text-slate-400 font-bold uppercase"
                        >
                          Sin filas. Presiona "+ Fila" o pega datos.
                        </td>
                      </tr>
                    )}
                    {rows.map((row, i) => (
                      <GeneratorRow
                        key={row.id}
                        row={row}
                        index={i}
                        updateRow={updateRowGenerador}
                        deleteRow={deleteRowGenerador}
                        calculateVolume={calculateVolumeRow}
                        isSelected={selectedGeneratorRows.includes(row.id)}
                        onToggleSelect={(id) =>
                          setSelectedGeneratorRows((prev) =>
                            prev.includes(id)
                              ? prev.filter((x) => x !== id)
                              : [...prev, id],
                          )
                        }
                      />
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
            <div
              className="w-full lg:w-[350px] bg-slate-100 p-6 flex flex-col border-l border-slate-200"
              onPaste={handlePasteImage}
            >
              <h4 className="font-black text-slate-600 uppercase text-xs mb-4 flex items-center gap-2 shrink-0">
                <FileDown size={16} /> Croquis / Referencia
              </h4>
              <div className="flex-1 overflow-y-auto space-y-4 pr-1 custom-scrollbar">
                {[0, 1, 2].map((i) => {
                  const currentImages = currentLevel.images || (currentLevel.image ? [currentLevel.image] : []);
                  const img = currentImages[i];
                  return (
                    <div key={i} className="w-full bg-white border-2 border-dashed border-slate-300 rounded-xl h-[250px] shrink-0 flex items-center justify-center relative overflow-hidden group">
                      {img ? (
                        <>
                          <img
                            src={img}
                            alt={`Referencia ${i + 1}`}
                            className="w-full h-full object-contain p-2"
                          />
                          <div className="absolute inset-0 bg-black/50 opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-center gap-4">
                            <button
                              onClick={() =>
                                updateActiveGeneradores((prev) => {
                                  const c = prev[concepto.id];
                                  const cLevel = c[nivel.id] || {};
                                  const imgs = cLevel.images || (cLevel.image ? [cLevel.image] : []);
                                  const newImgs = [...imgs];
                                  newImgs.splice(i, 1);
                                  return {
                                    ...prev,
                                    [concepto.id]: {
                                      ...c,
                                      [nivel.id]: { ...cLevel, images: newImgs, image: null },
                                    },
                                  };
                                })
                              }
                              className="p-3 bg-red-600 text-white rounded-full hover:bg-red-500"
                            >
                              <Trash2 size={20} />
                            </button>
                          </div>
                        </>
                      ) : (
                        <div className="text-center p-6">
                          <ClipboardPaste
                            size={24}
                            className="mx-auto text-slate-300 mb-2"
                          />
                          <p className="text-[10px] font-bold text-slate-400 uppercase">
                            Pegar Imagen {i + 1}
                          </p>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
              <div className="mt-6 bg-slate-800 text-white p-5 rounded-2xl shadow-lg">
                <p className="text-[10px] text-blue-300 font-black uppercase mb-1 tracking-widest">
                  Total Nivel {nivel.nombre}
                </p>
                <p className="text-3xl font-black">
                  {getVolumenNivel(concepto.id, nivel.id).toLocaleString(
                    undefined,
                    { minimumFractionDigits: 2, maximumFractionDigits: 2 },
                  )}{" "}
                  <span className="text-sm text-slate-400">
                    {concepto.unidad}
                  </span>
                </p>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  };

  const renderLevelReport = () => {
    if (!activeLevelReport) return null;
    const n = activeLevelReport;
    const conceptosConDatos = conceptos.filter(
      (c) =>
        getVolumenNivel(c.id, n.id) > 0 ||
        (generadores[c.id]?.[n.id]?.rows &&
          generadores[c.id][n.id].rows.length > 0),
    );

    return (
      <div className="fixed inset-0 bg-slate-900/80 backdrop-blur-sm z-[100] flex items-center justify-center p-4 sm:p-8 animate-in fade-in zoom-in-95 duration-200">
        <div className="bg-white rounded-3xl shadow-2xl w-full max-w-5xl flex flex-col h-[90vh] overflow-hidden">
          <div className="bg-blue-800 text-white p-6 flex justify-between items-center shrink-0">
            <div className="flex items-center gap-4">
              <div className="p-3 bg-blue-700 rounded-xl">
                <FileSpreadsheet size={24} className="text-blue-200" />
              </div>
              <div>
                <h3 className="font-black uppercase tracking-widest text-lg">
                  Cédula de Cuantificación (Generador)
                </h3>
                <p className="text-xs text-blue-300 font-bold uppercase mt-1">
                  Reporte Detallado de Nivel: {n.nombre}
                </p>
              </div>
            </div>
            <button
              onClick={() => setActiveLevelReport(null)}
              className="p-2 hover:bg-white/20 rounded-full transition-colors"
            >
              <X size={24} />
            </button>
          </div>
          <div className="p-6 bg-slate-100 flex-1 overflow-auto">
            {conceptosConDatos.length === 0 && (
              <div className="p-8 text-center text-slate-400 font-bold uppercase">
                No hay conceptos con generadores en este nivel
              </div>
            )}
            {conceptosConDatos.map((c) => {
              const dataNivel = generadores[c.id]?.[n.id] || { rows: [] };
              const rows = dataNivel.rows || [];
              const refImages = dataNivel.images || (dataNivel.image ? [dataNivel.image] : []);
              const totalVol = getVolumenNivel(c.id, n.id);

              return (
                <div
                  key={c.id}
                  className="mb-8 bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden break-inside-avoid"
                >
                  <div className="bg-slate-800 text-white p-6 relative rounded-t-xl overflow-hidden shadow-inner">
                    <div className="flex flex-col md:flex-row items-center justify-between gap-6 relative z-10">
                      <div className="flex flex-col items-center md:items-start bg-slate-700/30 p-3 rounded-2xl border border-slate-700/50 min-w-[120px]">
                        <span className="text-[9px] font-black text-slate-400 uppercase tracking-[0.25em] mb-1">
                          Concepto
                        </span>
                        <h4 className="font-black text-sm tracking-wider text-blue-400">
                          {c.clave || c.id}
                        </h4>
                      </div>
                      <div className="flex-1 flex items-center justify-center py-2 px-4">
                        <p className="text-[11px] md:text-[12px] leading-tight text-white font-black uppercase text-center max-w-2xl drop-shadow-sm">
                          {c.descripcion}
                        </p>
                      </div>
                      <div className="flex flex-col items-center md:items-end bg-slate-700/30 p-3 rounded-2xl border border-slate-700/50 min-w-[120px]">
                        <span className="text-[9px] font-black text-slate-400 uppercase tracking-[0.25em] mb-1">
                          Unidad
                        </span>
                        <span className="font-black text-lg text-blue-400 leading-none">
                          {c.unidad}
                        </span>
                      </div>
                    </div>
                  </div>
                  <div className="flex flex-col border border-slate-200 rounded-b-xl border-t-0">
                    <div className="w-full p-0">
                      <table className="w-full text-left border-collapse text-[12px]">
                        <thead className="bg-slate-50 text-slate-500 uppercase font-black text-[12px]">
                          <tr>
                            <th className="p-2 border-b border-r border-slate-200">
                              Localización
                            </th>
                            <th className="p-2 border-b border-r border-slate-200 text-center">
                              Largo
                            </th>
                            <th className="p-2 border-b border-r border-slate-200 text-center">
                              Ancho
                            </th>
                            <th className="p-2 border-b border-r border-slate-200 text-center">
                              Kg/ml
                            </th>
                            <th className="p-2 border-b border-r border-slate-200 text-center">
                              Alto
                            </th>
                            <th className="p-2 border-b border-r border-slate-200 text-center">
                              Vol/Pza
                            </th>
                            <th className="p-2 border-b border-r border-slate-200 text-center">
                              Pzas
                            </th>
                            <th className="p-2 border-b border-slate-200 text-right text-blue-700">
                              VolumetTot
                            </th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100 text-[12px]">
                          {rows.length === 0 && (
                            <tr>
                              <td
                                colSpan={8}
                                className="p-3 text-center text-slate-400"
                              >
                                Sin filas
                              </td>
                            </tr>
                          )}
                          {rows.map((r) => {
                            const l = parseFloat(r.largo);
                            const lAbs = isNaN(l) || l === 0 ? 0 : Math.abs(l);
                            const a = parseFloat(r.ancho);
                            const aAbs = isNaN(a) || a === 0 ? 0 : Math.abs(a);
                            const h = parseFloat(r.alto);
                            const hAbs = isNaN(h) || h === 0 ? 0 : Math.abs(h);
                            const k = parseFloat(r.kg_ml);
                            const kAbs = isNaN(k) || k === 0 ? 0 : Math.abs(k);
                            const isActuallyZero =
                              lAbs === 0 &&
                              aAbs === 0 &&
                              hAbs === 0 &&
                              kAbs === 0;
                            let calculatedVolPza = isActuallyZero
                              ? 0
                              : (lAbs > 0 ? lAbs : 1) *
                                (aAbs > 0 ? aAbs : 1) *
                                (hAbs > 0 ? hAbs : 1) *
                                (kAbs > 0 ? kAbs : 1);
                            if (r.isPasos)
                              calculatedVolPza = Math.floor(lAbs / 4);
                            const volPza =
                              typeof r.overrideVolPza === "number"
                                ? r.overrideVolPza
                                : calculatedVolPza;
                            return (
                              <tr key={r.id} className="hover:bg-slate-50">
                                <td className="p-2 border-r border-slate-100 font-bold uppercase">
                                  {r.eje} {r.claveLoc}
                                </td>
                                <td className="p-2 border-r border-slate-100 text-center">
                                  {r.largo
                                    ? parseFloat(r.largo).toFixed(2)
                                    : "0.00"}
                                </td>
                                <td className="p-2 border-r border-slate-100 text-center">
                                  {r.ancho
                                    ? parseFloat(r.ancho).toFixed(2)
                                    : "0.00"}
                                </td>
                                <td className="p-2 border-r border-slate-100 text-center">
                                  {r.kg_ml
                                    ? parseFloat(r.kg_ml).toFixed(3)
                                    : "0.000"}
                                </td>
                                <td className="p-2 border-r border-slate-100 text-center">
                                  {r.alto
                                    ? parseFloat(r.alto).toFixed(2)
                                    : "0.00"}
                                </td>
                                <td className="p-2 border-r border-slate-100 text-center">
                                  {volPza.toFixed(2)}
                                </td>
                                <td className="p-2 border-r border-slate-100 text-center font-bold">
                                  {r.piezas
                                    ? parseFloat(r.piezas).toFixed(2)
                                    : "1.00"}
                                </td>
                                <td className="p-2 text-right font-black text-slate-700">
                                  {calculateVolumeRow(r).toFixed(2)}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                        <tfoot className="bg-blue-50/50">
                          <tr>
                            <td
                              colSpan={7}
                              className="p-3 text-right font-black uppercase tracking-widest text-slate-600 border-t border-blue-100 text-[13.5px]"
                            >
                              Total Concepto:
                            </td>
                            <td className="p-3 text-right font-black text-blue-800 text-[15px] border-t border-blue-100">
                              {totalVol.toFixed(2)}
                            </td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                    {refImages.length > 0 && (
                      <div className="w-full bg-slate-50 border-t border-slate-200 p-4 flex flex-col gap-4">
                        <span className="text-[12px] font-black uppercase tracking-widest text-slate-400 text-center border-b border-slate-200 pb-1">
                          Croquis de Referencia
                        </span>
                        {refImages.map((croquis, i) => (
                           <div key={i} className="w-full flex items-center justify-center bg-white border-2 border-dashed border-slate-200 rounded-xl overflow-hidden p-2">
                             <img
                               src={croquis}
                               alt={`Croquis ${i + 1}`}
                               className="w-full h-auto object-contain max-h-[300px]"
                             />
                           </div>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
          <div className="bg-slate-100 p-5 flex justify-end border-t border-slate-200 shadow-inner z-10">
            <button
              onClick={() => {
                let html = ``;
                let isFirstTable = true;
                conceptosConDatos.forEach((c) => {
                  const rows = generadores[c.id]?.[n.id]?.rows || [];
                  const total = getVolumenNivel(c.id, n.id);
                  const cLabel = c.clave || c.id;
                  const desc = (c.descripcion || "").replace(/\\n/g, " ");

                  const renderExVal = (v) => {
                    const num = parseFloat(v);
                    if (isNaN(num) || num === 0) return "-  ";
                    return num.toFixed(2);
                  };

                  html += `
                   <table style="border-collapse: collapse; margin-bottom: 30px; font-family: Arial, sans-serif; font-size: 10pt;">
                      <colgroup>
                         <col width="20" />
                         <col width="145" />
                         <col width="145" />
                         <col width="145" />
                         <col width="145" />
                         <col width="145" />
                         <col width="145" />
                         <col width="145" />
                         <col width="145" />
                      </colgroup>
                      <thead>`;

                  if (isFirstTable) {
                    html += `
                         <tr><td colspan="9" style="height: 10px; border: none;"></td></tr>
                         <tr>
                            <td style="border: none;"></td>
                            <td colspan="8" style="background-color: #203764; color: #ffffff; font-size: 14pt; font-weight: bold; text-align: center; padding: 6px; border: none;">
                               CÉDULA DE CUANTIFICACIÓN (GENERADOR) - NIVEL: ${n.nombre.toUpperCase()}
                            </td>
                         </tr>`;
                    isFirstTable = false;
                  }

                  html += `
                         <tr><td colspan="9" style="height: 10px; border: none;"></td></tr>
                         <tr>
                            <td style="border: none;"></td>
                            <td style="background-color: #000000; color: #ffffff; padding: 4px 8px; text-align: left; font-size: 10pt; font-weight: bold; border-top: 2px solid #000000; border-left: 2px solid #000000; border-bottom: 1px solid #000000; border-right: 1px solid #ffffff;">
                               CÓD: ${cLabel}
                            </td>
                            <td colspan="5" style="background-color: #000000; color: #ffffff; padding: 10px; text-align: center; font-size: 10pt; font-weight: bold; border-top: 2px solid #000000; border-bottom: 1px solid #000000; border-right: 1px solid #ffffff; white-space: normal; word-wrap: break-word; mso-height-source: auto;">
                               ${desc}
                            </td>
                            <td colspan="2" style="background-color: #000000; color: #5B9BD5; padding: 4px 8px; text-align: center; font-size: 11pt; font-weight: bold; border-top: 2px solid #000000; border-right: 2px solid #000000; border-bottom: 1px solid #000000;">
                               ${c.unidad}
                            </td>
                         </tr>
                         <tr>
                            <td style="border: none;"></td>
                            <td style="color: #203764; font-weight: bold; text-align: left; padding: 4px; border: 1px solid #D9D9D9; border-left: 2px solid #000000; font-size: 9pt; text-transform: uppercase;">LOCALIZACIÓN</td>
                            <td style="color: #203764; font-weight: bold; text-align: center; padding: 4px; border: 1px solid #D9D9D9; font-size: 9pt; text-transform: uppercase;">LARGO</td>
                            <td style="color: #203764; font-weight: bold; text-align: center; padding: 4px; border: 1px solid #D9D9D9; font-size: 9pt; text-transform: uppercase;">ANCHO</td>
                            <td style="color: #203764; font-weight: bold; text-align: center; padding: 4px; border: 1px solid #D9D9D9; font-size: 9pt; text-transform: uppercase;">KG/ML</td>
                            <td style="color: #203764; font-weight: bold; text-align: center; padding: 4px; border: 1px solid #D9D9D9; font-size: 9pt; text-transform: uppercase;">ALTO</td>
                            <td style="color: #203764; font-weight: bold; text-align: center; padding: 4px; border: 1px solid #D9D9D9; font-size: 9pt; text-transform: uppercase;">VOL/PZA</td>
                            <td style="color: #203764; font-weight: bold; text-align: center; padding: 4px; border: 1px solid #D9D9D9; font-size: 9pt; text-transform: uppercase;">PZAS</td>
                            <td style="color: #203764; font-weight: bold; text-align: center; padding: 4px; border: 1px solid #D9D9D9; border-right: 2px solid #000000; font-size: 9pt; text-transform: uppercase;">VOLUMEN TOT</td>
                         </tr>
                      </thead>
                      <tbody>`;

                  if (rows.length === 0) {
                    html += `<tr><td style="border: none;"></td><td colspan="8" style="padding: 10px; text-align: center; color: #808080; border-left: 2px solid #000000; border-right: 2px solid #000000; border-bottom: 1px solid #D9D9D9;">Sin filas</td></tr>`;
                  } else {
                    rows.forEach((r, i) => {
                      const l = parseFloat(r.largo);
                      const lAbs = isNaN(l) || l === 0 ? 0 : Math.abs(l);
                      const a = parseFloat(r.ancho);
                      const aAbs = isNaN(a) || a === 0 ? 0 : Math.abs(a);
                      const h = parseFloat(r.alto);
                      const hAbs = isNaN(h) || h === 0 ? 0 : Math.abs(h);
                      const k = parseFloat(r.kg_ml);
                      const kAbs = isNaN(k) || k === 0 ? 0 : Math.abs(k);
                      const isActuallyZero =
                        lAbs === 0 && aAbs === 0 && hAbs === 0 && kAbs === 0;
                      let calculatedVolPza = isActuallyZero
                        ? 0
                        : (lAbs > 0 ? lAbs : 1) *
                          (aAbs > 0 ? aAbs : 1) *
                          (hAbs > 0 ? hAbs : 1) *
                          (kAbs > 0 ? kAbs : 1);
                      if (r.isPasos) calculatedVolPza = Math.floor(lAbs / 4);
                      const volPza =
                        typeof r.overrideVolPza === "number"
                          ? r.overrideVolPza
                          : calculatedVolPza;
                      html += `<tr>
                            <td style="border: none;"></td>
                            <td style="padding: 4px; border: 1px solid #D9D9D9; border-left: 2px solid #000000; text-transform: uppercase; text-align: left; color: #000000; font-size: 10pt;">${r.eje} ${r.claveLoc}</td>
                            <td style="padding: 4px; border: 1px solid #D9D9D9; text-align: center; color: #000000; font-size: 10pt; mso-number-format:'0\\.00';">${renderExVal(r.largo)}</td>
                            <td style="padding: 4px; border: 1px solid #D9D9D9; text-align: center; color: #000000; font-size: 10pt; mso-number-format:'0\\.00';">${renderExVal(r.ancho)}</td>
                            <td style="padding: 4px; border: 1px solid #D9D9D9; text-align: center; color: #000000; font-size: 10pt; mso-number-format:'0\\.000';">${r.kg_ml ? parseFloat(r.kg_ml).toFixed(3) : "-  "}</td>
                            <td style="padding: 4px; border: 1px solid #D9D9D9; text-align: center; color: #000000; font-size: 10pt; mso-number-format:'0\\.00';">${renderExVal(r.alto)}</td>
                            <td style="padding: 4px; border: 1px solid #D9D9D9; text-align: center; color: #000000; font-size: 10pt; mso-number-format:'0\\.00';">${volPza.toFixed(2)}</td>
                            <td style="padding: 4px; border: 1px solid #D9D9D9; text-align: center; font-weight: bold; color: #000000; font-size: 10pt; mso-number-format:'0\\.00';">${r.piezas ? parseFloat(r.piezas).toFixed(2) : "1.00"}</td>
                            <td style="padding: 4px; border: 1px solid #D9D9D9; border-right: 2px solid #000000; text-align: center; font-weight: bold; color: #000000; font-size: 10pt; mso-number-format:'0\\.00';">${calculateVolumeRow(r).toFixed(2)}</td>
                         </tr>`;
                    });
                  }

                  html += `</tbody>
                      <tfoot>
                         <tr>
                            <td style="border: none;"></td>
                            <td colspan="6" style="border-top: 2px solid #000000; border-bottom: 2px solid #000000; border-left: 2px solid #000000; border-right: none;"></td>
                            <td style="padding: 6px 10px; border-top: 2px solid #000000; border-bottom: 2px solid #000000; border-left: none; text-align: right; color: #595959; font-size: 10pt; font-weight: bold;">TOTAL:</td>
                            <td style="padding: 6px 10px; border-top: 2px solid #000000; border-bottom: 2px solid #000000; border-right: 2px solid #000000; text-align: right; font-weight: bold; color: #0070C0; font-size: 11pt; mso-number-format:'0\\.00';">${total.toFixed(2)}</td>
                         </tr>`;

                  if (refImages.length === 0) {
                      html += `
                         <tr><td colspan="9" style="height: 10px; border: none;"></td></tr>
                         <tr>
                            <td style="border: none;"></td>
                            <td colspan="8" style="height: 250px; border: 1.5px dashed #A6A6A6; text-align: center; vertical-align: middle; color: #BFBFBF; font-size: 10pt; text-transform: uppercase;">
                               [ ESPACIO PARA INSERTAR CROQUIS / IMAGEN DE REFERENCIA ]
                            </td>
                         </tr>`;
                  } else {
                      refImages.forEach((imgSrc) => {
                          html += `
                          <tr><td colspan="9" style="height: 10px; border: none;"></td></tr>
                          <tr>
                             <td style="border: none;"></td>
                             <td colspan="8" style="text-align: center; vertical-align: middle; border: 1.5px solid #D9D9D9; padding: 10px;">
                                <img src="${imgSrc}" style="max-height: 400px; max-width: 100%;" />
                             </td>
                          </tr>`;
                      });
                  }

                  html += `
                      </tfoot>
                   </table>`;
                });
                if (html === "") {
                  alert("No hay datos en este nivel.");
                  return;
                }
                exportFormattedExcel(html, `Cedula_Cuantificacion_${n.nombre}`);
              }}
              className="px-6 py-3 bg-emerald-600 text-white rounded-xl font-black text-[11px] uppercase tracking-wider shadow-md hover:bg-emerald-500 flex gap-2 items-center"
            >
              <FileDown size={16} /> Exportar Generadores Excel
            </button>
          </div>
        </div>
      </div>
    );
  };

  const renderGlobalSteelReport = () => {
    if (!showGlobalSteelReport) return null;

    const filteredEstructuras =
      globalReportLevel === "todos"
        ? todasLasEstructuras
        : todasLasEstructuras.filter((e) => e._nivelId === globalReportLevel);
    const estructurasConAcero = filteredEstructuras.filter(
      (e) => e.aceros && e.aceros.length > 0,
    );

    let totalGlobalAcero = 0;
    let resumenCimentacion = {};
    let resumenEstructura = {};
    let resumenMuros = {};

    estructurasConAcero.forEach((e) => {
      const ePiezas = parseFloat(e.piezas) || 1;
      const tipo = (e.tipo || "").toLowerCase().trim();
      const isCimentacion = [
        "zapata aislada",
        "zapata corrida",
        "contratrabe",
        "dado",
        "pilas",
        "pila",
      ].includes(tipo);
      const isMuro =
        tipo === "muro" ||
        tipo === "muro curvo" ||
        tipo === "muros" ||
        tipo === "muros curvos";

      e.aceros.forEach((a) => {
        const calc = calcAceroItem(a);
        const kgTotalItem = Math.round(calc.kg * ePiezas * 100) / 100;
        totalGlobalAcero += kgTotalItem;
        const numVar = a.numVarilla || "-";

        if (isCimentacion) {
          resumenCimentacion[numVar] =
            (resumenCimentacion[numVar] || 0) + kgTotalItem;
        } else if (isMuro) {
          resumenMuros[numVar] = (resumenMuros[numVar] || 0) + kgTotalItem;
        } else {
          resumenEstructura[numVar] =
            (resumenEstructura[numVar] || 0) + kgTotalItem;
        }
      });
    });

    return (
      <div className="fixed inset-0 bg-slate-900/80 backdrop-blur-sm z-[100] flex items-center justify-center p-4 sm:p-8 animate-in fade-in zoom-in-95 duration-200">
        <div className="bg-white rounded-3xl shadow-2xl w-full max-w-[1400px] flex flex-col h-[90vh] overflow-hidden">
          <div className="bg-slate-800 text-white p-6 flex justify-between items-center shrink-0">
            <div className="flex items-center gap-4">
              <div className="p-3 bg-slate-700 rounded-xl">
                <Printer size={24} className="text-slate-200" />
              </div>
              <div>
                <h3 className="font-black uppercase tracking-widest text-lg">
                  Reporte Global de Acero (Desglose por Elemento)
                </h3>
                <p className="text-xs text-slate-400 font-bold uppercase mt-1">
                  Análisis detallado de varilla y refuerzos
                </p>
              </div>
            </div>
            <div className="flex items-center gap-4">
              <select
                value={globalReportLevel}
                onChange={(e) => setGlobalReportLevel(e.target.value)}
                className="bg-slate-700 text-white text-xs font-bold uppercase p-2 rounded-lg outline-none border border-slate-600"
              >
                <option value="todos">Todos los Niveles</option>
                {niveles.map((n) => (
                  <option key={n.id} value={n.id}>
                    {n.nombre}
                  </option>
                ))}
              </select>
              <button
                onClick={() => {
                  let titleLevelName =
                    globalReportLevel === "todos"
                      ? "TODOS LOS NIVELES"
                      : niveles.find((n) => n.id === globalReportLevel)
                          ?.nombre || "";

                  const renderExVal = (v) => {
                    const num = parseFloat(v);
                    if (isNaN(num)) return "0.00";
                    return num.toFixed(2);
                  };

                  let html = `<table style="table-layout: fixed; border-collapse: collapse; font-family: Arial, sans-serif;">
                      <colgroup>
                         <col width="30" style="width: 30px;" />
                         <col width="100" style="width: 100px;" />
                         <col width="160" style="width: 160px;" />
                         <col width="140" style="width: 140px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="130" style="width: 130px;" />
                      </colgroup>
                      <thead>
                      <tr style="height: 20pt;">
                         <td style="border: none; background-color: #ffffff;"></td>
                         <th colspan="17" style="background-color: #0b1a30; color: #ffffff; font-size: 11pt; font-family: Arial, sans-serif; font-weight: bold; text-align: center; vertical-align: middle; border: 1.5pt solid #000000; text-transform: uppercase;">REPORTE GLOBAL DE ACERO (DESGLOSE POR ELEMENTO) - NIVEL: ${titleLevelName}</th>
                      </tr>
                      <tr style="height: 10pt;">
                         <td colspan="18" style="border: none; background-color: #ffffff;"></td>
                      </tr>
                      <tr style="background-color: #ffffff; color: #000000; font-size: 8.5pt; font-family: Arial, sans-serif; font-weight: bold; text-align: center; height: 18pt;">
                         <td style="border: none; background-color: #ffffff;"></td>
                         <th width="100" style="width: 100px; border: 1pt solid #000000; border-top: 1.5pt solid #000000; border-left: 1.5pt solid #000000; vertical-align: middle; background-color: #ffffff;">NIVEL</th>
                         <th width="160" style="width: 160px; border: 1pt solid #000000; border-top: 1.5pt solid #000000; vertical-align: middle; background-color: #ffffff;">UBICACIÓN (EJE / CLAVE)</th>
                         <th width="140" style="width: 140px; border: 1pt solid #000000; border-top: 1.5pt solid #000000; vertical-align: middle; background-color: #ffffff;">ELEMENTO</th>
                         <th width="115" style="width: 115px; border: 1pt solid #000000; border-top: 1.5pt solid #000000; vertical-align: middle; background-color: #ffffff;">PZAS</th>
                         <th width="115" style="width: 115px; border: 1pt solid #000000; border-top: 1.5pt solid #000000; vertical-align: middle; background-color: #ffffff;">TIPO REF.</th>
                         <th width="115" style="width: 115px; border: 1pt solid #000000; border-top: 1.5pt solid #000000; vertical-align: middle; background-color: #ffffff;"># VAR</th>
                         <th width="115" style="width: 115px; border: 1pt solid #000000; border-top: 1.5pt solid #000000; vertical-align: middle; background-color: #ffffff;">TRAMO</th>
                         <th width="115" style="width: 115px; border: 1pt solid #000000; border-top: 1.5pt solid #000000; vertical-align: middle; background-color: #ffffff;">SEP.</th>
                         <th width="115" style="width: 115px; border: 1pt solid #000000; border-top: 1.5pt solid #000000; vertical-align: middle; background-color: #ffffff;">PZAS REF.</th>
                         <th width="115" style="width: 115px; border: 1pt solid #000000; border-top: 1.5pt solid #000000; vertical-align: middle; background-color: #ffffff;">LONG.</th>
                         <th width="115" style="width: 115px; border: 1pt solid #000000; border-top: 1.5pt solid #000000; vertical-align: middle; background-color: #ffffff;">GAN.</th>
                         <th width="115" style="width: 115px; border: 1pt solid #000000; border-top: 1.5pt solid #000000; vertical-align: middle; background-color: #ffffff;">ANCL.</th>
                         <th width="115" style="width: 115px; border: 1pt solid #000000; border-top: 1.5pt solid #000000; vertical-align: middle; background-color: #ffffff;">TRAS.</th>
                         <th width="115" style="width: 115px; border: 1pt solid #000000; border-top: 1.5pt solid #000000; vertical-align: middle; background-color: #ffffff;">ML/PZA</th>
                         <th width="115" style="width: 115px; border: 1pt solid #000000; border-top: 1.5pt solid #000000; color: #0054ff; vertical-align: middle; background-color: #ffffff;">TOT.(ML)</th>
                         <th width="115" style="width: 115px; border: 1pt solid #000000; border-top: 1.5pt solid #000000; color: #0054ff; vertical-align: middle; background-color: #ffffff;">TOT.(KG)</th>
                         <th width="130" style="width: 130px; border: 1pt solid #000000; border-top: 1.5pt solid #000000; border-right: 1.5pt solid #000000; color: #0054ff; vertical-align: middle; background-color: #ffffff;">TOTAL ELEM.(KG)</th>
                      </tr>
                   </thead><tbody>`;

                  estructurasConAcero.forEach((e) => {
                    const ePiezas = parseFloat(e.piezas) || 1;
                    const elemTotalKg = calcAceroTotalKg(e.aceros, e.piezas);
                    const tipoRowSpan =
                      e.aceros.length > 0 ? e.aceros.length : 1;

                    if (e.aceros.length === 0) {
                      html += `<tr style="background-color: #ffffff; text-align: center; height: 18pt;">
                            <td style="border: none; background-color: #ffffff;"></td>
                            <td style="border: 1pt solid #000000; border-left: 1.5pt solid #000000; border-bottom: 1.5pt solid #000000; font-size: 8.5pt; font-weight: bold; color: #000000; vertical-align: middle; text-align: left; padding-left: 5px;">${e._nivelInfo}</td>
                            <td style="border: 1pt solid #000000; border-bottom: 1.5pt solid #000000; font-size: 8.5pt; font-weight: normal; color: #000000; vertical-align: middle; text-align: center;">${e.eje} ${e.clave}</td>
                            <td style="border: 1pt solid #000000; border-bottom: 1.5pt solid #000000; font-size: 11pt; font-weight: bold; color: #000000; vertical-align: middle; text-align: center;">${e.tipo}</td>
                            <td style="border: 1pt solid #000000; border-bottom: 1.5pt solid #000000; font-size: 10pt; font-weight: normal; color: #0054ff; vertical-align: middle; text-align: center; mso-number-format:'0\\.00';">${ePiezas.toFixed(2)}</td>
                            <td colspan="12" style="border: 1pt solid #000000; border-bottom: 1.5pt solid #000000; font-size: 8.5pt; font-weight: normal; color: #000000; vertical-align: middle; text-align: center;">Sin acero</td>
                            <td style="border: 1pt solid #000000; border-right: 1.5pt solid #000000; border-bottom: 1.5pt solid #000000; font-size: 11pt; font-weight: bold; color: #000000; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${elemTotalKg.toFixed(2)}</td>
                         </tr>`;
                    } else {
                      e.aceros.forEach((a, idx) => {
                        const calc = calcAceroItem(a);
                        const kgTotal =
                          Math.round(calc.kg * ePiezas * 100) / 100;
                        const isLastAcero = idx === e.aceros.length - 1;

                        const bottomBorderStyle = isLastAcero
                          ? "1.5pt solid #000000"
                          : "1pt solid #000000";
                        const rowStyle = `border: 1pt solid #000000; border-bottom: ${bottomBorderStyle};`;
                        const rowspanStyle = `border: 1pt solid #000000; border-bottom: 1.5pt solid #000000;`;

                        html += `<tr style="background-color: #ffffff; text-align: center; height: 18pt;">
                               <td style="border: none; background-color: #ffffff;"></td>`;

                        if (idx === 0) {
                          html += `
                               <td rowspan="${tipoRowSpan}" style="${rowspanStyle} border-left: 1.5pt solid #000000; font-size: 8.5pt; font-weight: bold; color: #000000; vertical-align: middle; text-align: left; padding-left: 5px;">${e._nivelInfo}</td>
                               <td rowspan="${tipoRowSpan}" style="${rowspanStyle} font-size: 8.5pt; font-weight: normal; color: #000000; vertical-align: middle; text-align: center;">${e.eje} ${e.clave}</td>
                               <td rowspan="${tipoRowSpan}" style="${rowspanStyle} font-size: 11pt; font-weight: bold; color: #000000; vertical-align: middle; text-align: center;">${e.tipo}</td>
                               <td rowspan="${tipoRowSpan}" style="${rowspanStyle} font-size: 10pt; font-weight: normal; color: #0054ff; vertical-align: middle; text-align: center; mso-number-format:'0\\.00';">${ePiezas.toFixed(2)}</td>`;
                        }

                        html += `
                               <td style="${rowStyle} font-size: 8.5pt; color: #000000; font-weight: normal; text-align: center; vertical-align: middle;">${a.tipo}</td>
                               <td style="${rowStyle} font-size: 8.5pt; color: #000000; font-weight: normal; text-align: center; vertical-align: middle;">#${parseFloat(a.numVarilla || 0).toFixed(2)}</td>
                               <td style="${rowStyle} font-size: 8.5pt; color: #000000; font-weight: normal; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${renderExVal(a.anchoCalc)}</td>
                               <td style="${rowStyle} font-size: 8.5pt; color: #000000; font-weight: normal; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${renderExVal(a.separacion)}</td>
                               <td style="${rowStyle} font-size: 8.5pt; color: #000000; font-weight: bold; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${parseFloat(a.piezas || 1).toFixed(2)}</td>
                               <td style="${rowStyle} font-size: 8.5pt; color: #000000; font-weight: normal; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${renderExVal(a.longitud)}</td>
                               <td style="${rowStyle} font-size: 8.5pt; color: #000000; font-weight: normal; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${renderExVal(a.ganchos)}</td>
                               <td style="${rowStyle} font-size: 8.5pt; color: #000000; font-weight: normal; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${renderExVal(a.anclaje)}</td>
                               <td style="${rowStyle} font-size: 8.5pt; color: #000000; font-weight: normal; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${renderExVal(a.traslapes)}</td>
                               <td style="${rowStyle} font-size: 8.5pt; color: #000000; font-weight: bold; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${calc.mlPorPieza.toFixed(2)}</td>
                               <td style="${rowStyle} font-size: 8.5pt; color: #000000; font-weight: normal; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${(calc.ml * ePiezas).toFixed(2)}</td>
                               <td style="${rowStyle} font-size: 8.5pt; color: #000000; font-weight: bold; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${kgTotal.toFixed(2)}</td>`;

                        if (idx === 0) {
                          html += `<td rowspan="${tipoRowSpan}" style="${rowspanStyle} border-right: 1.5pt solid #000000; font-size: 11pt; color: #000000; font-weight: bold; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${elemTotalKg.toFixed(2)}</td>`;
                        }
                        html += `</tr>`;
                      });
                    }
                  });
                  html += `</tbody></table><br/><br/>`;

                  const cimArr = Object.entries(resumenCimentacion).sort(
                    (a, b) => Number(a[0]) - Number(b[0]),
                  );
                  const estArr = Object.entries(resumenEstructura).sort(
                    (a, b) => Number(a[0]) - Number(b[0]),
                  );
                  const murArr = Object.entries(resumenMuros).sort(
                    (a, b) => Number(a[0]) - Number(b[0]),
                  );

                  const maxRows = Math.max(
                    cimArr.length,
                    estArr.length,
                    murArr.length,
                    1,
                  );

                  html += `<table border="0" style="table-layout: fixed; border-collapse: collapse; text-align: center; font-family: Arial, sans-serif;">
                      <colgroup>
                         <col width="30" style="width: 30px;" />
                         <col width="100" style="width: 100px;" />
                         <col width="160" style="width: 160px;" />
                         <col width="140" style="width: 140px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="115" style="width: 115px;" />
                         <col width="130" style="width: 130px;" />
                      </colgroup>
                     <thead>
                        <tr style="height: 18pt;">
                           <td style="width: 30px; border: none;"></td>
                           <th colspan="2" style="background-color: #008060; color: #ffffff; font-size: 11pt; border: 1.5pt solid #000000;">CIMENTACIÓN</th>
                           <th colspan="7" style="background-color: #0054ff; color: #ffffff; font-size: 11pt; border: 1.5pt solid #000000;">ESTRUCTURA</th>
                           <th colspan="7" style="background-color: #6600cc; color: #ffffff; font-size: 11pt; border: 1.5pt solid #000000;">MUROS</th>
                           <th colspan="1" style="background-color: #000000; color: #ffffff; font-size: 11pt; border: 1.5pt solid #000000;">TOTAL</th>
                        </tr>
                        <tr style="height: 16pt;">
                           <td style="border: none;"></td>
                           <th colspan="1" style="border: 1pt solid #000000; border-left: 1.5pt solid #000000; background-color: #ccffeb; color: #008060; padding: 2px;">CALIBRE</th>
                           <th colspan="1" style="border: 1pt solid #000000; border-right: 1.5pt solid #000000; background-color: #ccffeb; color: #008060; padding: 2px;">PESO (KG)</th>
                           
                           <th colspan="4" style="border: 1pt solid #000000; background-color: #e6f0ff; color: #0054ff; padding: 2px;">CALIBRE</th>
                           <th colspan="3" style="border: 1pt solid #000000; border-right: 1.5pt solid #000000; background-color: #e6f0ff; color: #0054ff; padding: 2px;">PESO (KG)</th>
                           
                           <th colspan="4" style="border: 1pt solid #000000; background-color: #f2e6ff; color: #6600cc; padding: 2px;">CALIBRE</th>
                           <th colspan="3" style="border: 1pt solid #000000; border-right: 1.5pt solid #000000; background-color: #f2e6ff; color: #6600cc; padding: 2px;">PESO<br/>(KG)</th>
                           
                           <td colspan="1" rowspan="${maxRows + 1}" style="border: 1.5pt solid #000000; background-color: #ffffff; color: #0054ff; font-weight: bold; font-size: 14pt; vertical-align: middle; text-align: center; mso-number-format:'0\\.00';">${totalGlobalAcero.toFixed(2)}</td>
                        </tr>
                     </thead>
                     <tbody>`;

                  for (let i = 0; i < maxRows; i++) {
                    const isLast = i === maxRows - 1;
                    const btmBorder = isLast
                      ? "1.5pt solid #000000"
                      : "1pt solid #000000";

                    const cItem = cimArr[i];
                    const eItem = estArr[i];
                    const mItem = murArr[i];

                    html += `<tr style="height: 16pt; background-color: #ffffff;">
                        <td style="border: none;"></td>`;

                    if (cimArr.length === 0 && i === 0) {
                      html += `<td colspan="1" style="border: 1pt solid #000000; border-bottom: ${btmBorder}; border-left: 1.5pt solid #000000; color: #a6a6a6; text-align: center; vertical-align: middle;">Sin datos</td>
                                 <td colspan="1" style="border: 1pt solid #000000; border-bottom: ${btmBorder}; border-right: 1.5pt solid #000000;"></td>`;
                    } else if (cItem) {
                      html += `<td colspan="1" style="border: 1pt solid #000000; border-bottom: ${btmBorder}; border-left: 1.5pt solid #000000; font-weight: bold; color: #0054ff; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">#${parseFloat(cItem[0]).toFixed(2)}</td>
                                 <td colspan="1" style="border: 1pt solid #000000; border-bottom: ${btmBorder}; border-right: 1.5pt solid #000000; font-weight: bold; color: #000000; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${cItem[1].toFixed(2)}</td>`;
                    } else {
                      html += `<td colspan="1" style="border: 1pt solid #000000; border-bottom: ${btmBorder}; border-left: 1.5pt solid #000000;"></td>
                                 <td colspan="1" style="border: 1pt solid #000000; border-bottom: ${btmBorder}; border-right: 1.5pt solid #000000;"></td>`;
                    }

                    if (estArr.length === 0 && i === 0) {
                      html += `<td colspan="4" style="border: 1pt solid #000000; border-bottom: ${btmBorder}; color: #a6a6a6; text-align: center; vertical-align: middle;">Sin datos</td>
                                 <td colspan="3" style="border: 1pt solid #000000; border-bottom: ${btmBorder}; border-right: 1.5pt solid #000000;"></td>`;
                    } else if (eItem) {
                      html += `<td colspan="4" style="border: 1pt solid #000000; border-bottom: ${btmBorder}; font-weight: bold; color: #0054ff; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">#${parseFloat(eItem[0]).toFixed(2)}</td>
                                 <td colspan="3" style="border: 1pt solid #000000; border-bottom: ${btmBorder}; border-right: 1.5pt solid #000000; font-weight: bold; color: #000000; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${eItem[1].toFixed(2)}</td>`;
                    } else {
                      html += `<td colspan="4" style="border: 1pt solid #000000; border-bottom: ${btmBorder};"></td>
                                 <td colspan="3" style="border: 1pt solid #000000; border-bottom: ${btmBorder}; border-right: 1.5pt solid #000000;"></td>`;
                    }

                    if (murArr.length === 0 && i === 0) {
                      html += `<td colspan="4" style="border: 1pt solid #000000; border-bottom: ${btmBorder}; color: #a6a6a6; text-align: center; vertical-align: middle;">Sin datos</td>
                                 <td colspan="3" style="border: 1pt solid #000000; border-bottom: ${btmBorder}; border-right: 1.5pt solid #000000;"></td>`;
                    } else if (mItem) {
                      html += `<td colspan="4" style="border: 1pt solid #000000; border-bottom: ${btmBorder}; font-weight: bold; color: #0054ff; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">#${parseFloat(mItem[0]).toFixed(2)}</td>
                                 <td colspan="3" style="border: 1pt solid #000000; border-bottom: ${btmBorder}; border-right: 1.5pt solid #000000; font-weight: bold; color: #000000; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${mItem[1].toFixed(2)}</td>`;
                    } else {
                      html += `<td colspan="4" style="border: 1pt solid #000000; border-bottom: ${btmBorder};"></td>
                                 <td colspan="3" style="border: 1pt solid #000000; border-bottom: ${btmBorder}; border-right: 1.5pt solid #000000;"></td>`;
                    }
                    html += `</tr>`;
                  }

                  html += `</tbody></table>`;

                  exportFormattedExcel(
                    html,
                    `Reporte_Global_Acero_${globalReportLevel}`,
                  );
                }}
                className="px-4 py-2 bg-emerald-600 text-white rounded-lg font-black text-[10px] uppercase tracking-wider shadow-sm hover:bg-emerald-500 transition-colors flex items-center gap-2"
              >
                <FileDown size={14} /> Excel
              </button>
              <button
                onClick={() => setShowGlobalSteelReport(false)}
                className="p-2 hover:bg-white/20 rounded-full transition-colors"
              >
                <X size={24} />
              </button>
            </div>
          </div>
          <div className="p-0 flex-1 overflow-auto bg-slate-50 flex flex-col">
            <div className="flex-1 overflow-auto">
              <table className="w-full text-left border-collapse text-[10px] md:text-xs">
                <thead className="bg-slate-200 text-slate-700 uppercase font-black sticky top-0 shadow-sm z-10">
                  <tr>
                    <th className="p-3 border-b border-slate-300 w-[8%]">
                      Nivel
                    </th>
                    <th className="p-3 border-b border-slate-300 w-[5%]">
                      Ubicación (Eje / Clave)
                    </th>
                    <th className="p-3 border-b border-slate-300 w-[5%]">
                      Elemento
                    </th>
                    <th className="p-3 border-b border-slate-300 text-center w-[5%]">
                      Pzas
                    </th>
                    <th className="p-0 border-b border-slate-300 w-[77%]">
                      <table className="w-full">
                        <thead>
                          <tr>
                            <th className="w-[10%] p-2 text-center border-r border-slate-300">
                              Tipo Ref.
                            </th>
                            <th className="w-[8%] p-2 text-center border-r border-slate-300">
                              # Var
                            </th>
                            <th className="w-[8%] p-2 text-center border-r border-slate-300">
                              Tramo
                            </th>
                            <th className="w-[8%] p-2 text-center border-r border-slate-300">
                              Sep.
                            </th>
                            <th className="w-[8%] p-2 text-center border-r border-slate-300">
                              Pzas
                            </th>
                            <th className="w-[8%] p-2 text-center border-r border-slate-300">
                              Long.
                            </th>
                            <th className="w-[8%] p-2 text-center border-r border-slate-300">
                              Gan.
                            </th>
                            <th className="w-[8%] p-2 text-center border-r border-slate-300">
                              Ancl.
                            </th>
                            <th className="w-[8%] p-2 text-center border-r border-slate-300">
                              Tras.
                            </th>
                            <th className="w-[8%] p-2 text-center border-r border-slate-300">
                              ML/Pza
                            </th>
                            <th className="w-[9%] p-2 text-right border-r border-slate-300">
                              Tot.(ML)
                            </th>
                            <th className="w-[9%] p-2 text-right">Tot.(KG)</th>
                          </tr>
                        </thead>
                      </table>
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {estructurasConAcero.length === 0 && (
                    <tr>
                      <td
                        colSpan={5}
                        className="p-8 text-center text-slate-400 font-bold uppercase"
                      >
                        No hay elementos con acero registrado en este nivel.
                      </td>
                    </tr>
                  )}
                  {estructurasConAcero.map((e, i) => {
                    const ePiezas = parseFloat(e.piezas) || 1;
                    return (
                      <tr
                        key={i}
                        className="hover:bg-white transition-colors bg-slate-50/30"
                      >
                        <td className="p-3 border-r border-slate-200 font-bold text-slate-600">
                          {e._nivelInfo}
                        </td>
                        <td className="p-3 border-r border-slate-200 font-black uppercase text-slate-800">
                          {e.eje} {e.clave}
                        </td>
                        <td className="p-3 border-r border-slate-200 uppercase text-[10px] text-slate-600 font-bold">
                          {e.tipo} <br />
                          <span className="text-[8px] text-slate-400">
                            ({e.largo} x {e.ancho} x {e.alto})
                          </span>
                        </td>
                        <td className="p-3 border-r border-slate-200 text-center font-black text-indigo-700 bg-indigo-50/30 text-sm">
                          {ePiezas}
                        </td>
                        <td className="p-0 border-r border-slate-200">
                          <table className="w-full text-[10px]">
                            <tbody className="divide-y divide-slate-100">
                              {e.aceros.map((a, j) => {
                                const calc = calcAceroItem(a);
                                const kgTotal =
                                  Math.round(calc.kg * ePiezas * 100) / 100;
                                return (
                                  <tr key={j} className="hover:bg-blue-50/50">
                                    <td className="w-[10%] p-2 text-center border-r border-slate-100 font-bold text-slate-600 uppercase">
                                      {a.tipo}
                                    </td>
                                    <td className="w-[8%] p-2 text-center border-r border-slate-100 font-black text-slate-800">
                                      #{a.numVarilla}
                                    </td>
                                    <td className="w-[8%] p-2 text-center border-r border-slate-100 font-bold text-indigo-900">
                                      {a.anchoCalc || "-"}
                                    </td>
                                    <td className="w-[8%] p-2 text-center border-r border-slate-100 font-bold text-indigo-900">
                                      {a.separacion || "-"}
                                    </td>
                                    <td className="w-[8%] p-2 text-center border-r border-slate-100 font-black text-blue-900">
                                      {a.piezas || "1"}
                                    </td>
                                    <td className="w-[8%] p-2 text-center border-r border-slate-100 font-black text-blue-800">
                                      {parseFloat(a.longitud || 0).toFixed(2)}
                                    </td>
                                    <td className="w-[8%] p-2 text-center border-r border-slate-100 font-black text-emerald-800">
                                      {a.ganchos || "0"}
                                    </td>
                                    <td className="w-[8%] p-2 text-center border-r border-slate-100 font-black text-emerald-800">
                                      {a.anclaje || "0"}
                                    </td>
                                    <td className="w-[8%] p-2 text-center border-r border-slate-100 font-black text-emerald-800">
                                      {parseFloat(a.traslapes || 0).toFixed(2)}
                                    </td>
                                    <td className="w-[8%] p-2 text-center border-r border-slate-100 font-black text-slate-500">
                                      {calc.mlPorPieza.toFixed(2)}
                                    </td>
                                    <td className="w-[9%] p-2 text-right border-r border-slate-100 font-black text-slate-600">
                                      {(calc.ml * ePiezas).toFixed(2)}
                                    </td>
                                    <td className="w-[9%] p-2 text-right font-black text-blue-700">
                                      {kgTotal.toFixed(2)}
                                    </td>
                                  </tr>
                                );
                              })}
                            </tbody>
                          </table>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <div className="w-full bg-white border-t border-slate-200 p-6 shadow-md z-20 flex flex-col xl:flex-row gap-8 shrink-0">
              <div className="flex-1 flex flex-col md:flex-row gap-6 border-r border-slate-100 pr-4">
                <div className="flex-1">
                  <h4 className="text-[10px] font-black uppercase tracking-widest text-teal-600 mb-3 border-b border-teal-100 pb-1">
                    Cimentación
                  </h4>
                  <div className="grid grid-cols-2 gap-2">
                    {Object.keys(resumenCimentacion).length === 0 && (
                      <span className="text-[9px] text-slate-400 font-bold uppercase">
                        Sin datos
                      </span>
                    )}
                    {Object.entries(resumenCimentacion)
                      .sort((a, b) => Number(a[0]) - Number(b[0]))
                      .map(([num, kg]) => (
                        <div
                          key={num}
                          className="bg-teal-50/50 border border-teal-100 rounded-lg p-2"
                        >
                          <span className="block text-[8px] font-black text-teal-500 uppercase">
                            Varilla #{num}
                          </span>
                          <span className="font-black text-teal-700 text-sm">
                            {kg.toLocaleString(undefined, {
                              minimumFractionDigits: 2,
                            })}{" "}
                            <span className="text-[8px]">kg</span>
                          </span>
                        </div>
                      ))}
                  </div>
                </div>

                <div className="flex-1">
                  <h4 className="text-[10px] font-black uppercase tracking-widest text-blue-600 mb-3 border-b border-blue-100 pb-1">
                    Estructura
                  </h4>
                  <div className="grid grid-cols-2 gap-2">
                    {Object.keys(resumenEstructura).length === 0 && (
                      <span className="text-[9px] text-slate-400 font-bold uppercase">
                        Sin datos
                      </span>
                    )}
                    {Object.entries(resumenEstructura)
                      .sort((a, b) => Number(a[0]) - Number(b[0]))
                      .map(([num, kg]) => (
                        <div
                          key={num}
                          className="bg-blue-50/50 border border-blue-100 rounded-lg p-2"
                        >
                          <span className="block text-[8px] font-black text-blue-500 uppercase">
                            Varilla #{num}
                          </span>
                          <span className="font-black text-blue-700 text-sm">
                            {kg.toLocaleString(undefined, {
                              minimumFractionDigits: 2,
                            })}{" "}
                            <span className="text-[8px]">kg</span>
                          </span>
                        </div>
                      ))}
                  </div>
                </div>

                <div className="flex-1">
                  <h4 className="text-[10px] font-black uppercase tracking-widest text-indigo-600 mb-3 border-b border-indigo-100 pb-1">
                    Muros
                  </h4>
                  <div className="flex flex-col gap-3">
                    {Object.keys(resumenMuros).length === 0 && (
                      <span className="text-[9px] text-slate-400 font-bold uppercase">
                        Sin datos
                      </span>
                    )}
                    {Object.entries(resumenMuros)
                      .sort((a, b) => Number(a[0]) - Number(b[0]))
                      .map(([num, kg]) => (
                        <div
                          key={num}
                          className="bg-indigo-50 border border-indigo-100 rounded-lg p-2.5"
                        >
                          <span className="block text-[9px] font-black text-indigo-500 uppercase tracking-wider mb-1">
                            Varilla #{num}
                          </span>
                          <span className="font-black text-indigo-700 text-lg md:text-xl">
                            {kg.toLocaleString(undefined, {
                              minimumFractionDigits: 2,
                            })}{" "}
                            <span className="text-xs">kg</span>
                          </span>
                        </div>
                      ))}
                  </div>
                </div>
              </div>
              <div className="w-full xl:w-[350px] bg-slate-900 p-6 rounded-3xl text-white shadow-lg relative overflow-hidden flex flex-col justify-center">
                <div className="absolute right-0 top-0 opacity-10 transform scale-150 -translate-y-4 translate-x-4">
                  <Wrench size={100} />
                </div>
                <span className="block text-[10px] font-black text-blue-300 uppercase tracking-widest mb-2 relative z-10">
                  Gran Total Acero (
                  {globalReportLevel === "todos" ? "Global" : "Nivel"})
                </span>
                <div className="flex items-baseline gap-2 relative z-10">
                  <span className="text-4xl font-black text-white tracking-tighter">
                    {totalGlobalAcero.toLocaleString(undefined, {
                      minimumFractionDigits: 2,
                    })}
                  </span>
                  <span className="text-blue-400 font-black">kg</span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  };

  const getMetalRowWeights = (r) => {
    let pzas = parseFloat(r.pzas);
    if (isNaN(pzas)) pzas = 1;
    let ancho = parseFloat(r.ancho);
    if (isNaN(ancho)) ancho = 1;

    let p = 0;
    let w = 0;

    if (r.tipo === "Polín Monten") {
      const pd = POLIN_DATA[r.medidaNominal];
      p = pd?.pesos[r.calibre] || 0;
      w = (parseFloat(r.largo) || 0) * p * pzas;
    } else if (r.tipo === "Viga IPR") {
      const pdList = IPR_DATA[r.medidaNominal] || [];
      const item = pdList.find((x) => x.kg.toString() === r.calibre) ||
        pdList[0] || { kg: 0 };
      p = item.kg;
      w = (parseFloat(r.largo) || 0) * item.kg * pzas;
    } else if (r.tipo === "Perfil HSS Cuadrado") {
      const pd = HSS_DATA[r.medidaNominal];
      p = pd?.pesos[r.calibre] || 0;
      w = (parseFloat(r.largo) || 0) * p * pzas;
    } else if (r.tipo === "Perfil HSS Rectangular") {
      const pd = HSS_RECT_DATA[r.medidaNominal];
      p = pd?.pesos[r.calibre] || 0;
      w = (parseFloat(r.largo) || 0) * p * pzas;
    } else if (r.tipo === "Placas Metálicas") {
      p = PLACAS_DATA[r.calibre] || 0;
      const m2 = (parseFloat(r.largo) || 0) * ancho;
      w = m2 * p * pzas;
    } else if (r.tipo === "Lámina Galvanizada") {
      const pd = LAMINA_DATA[r.medidaNominal];
      p = pd?.pesos[r.calibre] || 0;
      const m2 = (parseFloat(r.largo) || 0) * ancho;
      w = m2 * p * pzas;
    } else if (r.tipo === "Perfil PTR") {
      const pd = PTR_DATA[r.medidaNominal];
      p = pd?.pesos[r.calibre] || 0;
      w = (parseFloat(r.largo) || 0) * p * pzas;
    } else if (r.tipo === "Canales Tipo C") {
      const pd = CANALES_C_DATA[r.medidaNominal];
      p = pd?.pesos[r.calibre] || 0;
      w = (parseFloat(r.largo) || 0) * p * pzas;
    } else if (r.tipo === "Perfiles OC") {
      const pd = PERFILES_OC_DATA[r.medidaNominal];
      p = pd?.pesos[r.calibre] || 0;
      w = (parseFloat(r.largo) || 0) * p * pzas;
    } else if (r.tipo === "Viga IR") {
      p = parseFloat(r.calibre) || 0;
      w = (parseFloat(r.largo) || 0) * p * pzas;
    }

    return { pzas, p, w };
  };

  const renderAnalysisNervadura = () => {
    if (!showAnalysisNervadura) return null;

    const availableNervaduras = estructuras.filter((e) =>
      ["losa nervada", "nervadura"].includes((e.tipo || "").toLowerCase()),
    );

    const toggleSelection = (eId) => {
      setNervaduraAnalysisSelection((prev) =>
        prev.includes(eId) ? prev.filter((id) => id !== eId) : [...prev, eId],
      );
    };

    const toggleAll = () => {
      if (
        nervaduraAnalysisSelection.length === availableNervaduras.length &&
        availableNervaduras.length > 0
      ) {
        setNervaduraAnalysisSelection([]);
      } else {
        setNervaduraAnalysisSelection(availableNervaduras.map((e) => e.id));
      }
    };

    const selectedElements = availableNervaduras.filter((e) =>
      nervaduraAnalysisSelection.includes(e.id),
    );

    let totalCaseton = 0;
    let totalConcreto = 0;
    let totalCimbra = 0;
    let totalCimbraFrontera = 0;
    const acerosDetalle = { 3: 0, 4: 0, 5: 0, 6: 0, 8: 0, 10: 0 };
    let totalAceroVarilla = 0;

    selectedElements.forEach((e) => {
      const isLosaNervada = (e.tipo || "").toLowerCase() === "losa nervada";
      if (isLosaNervada) {
        totalCaseton += getCasetonesTotalVol(e.casetones || [], e.piezas);
      }
      totalConcreto += calcConcreto(e);
      totalCimbra += calcCimbra(e);
      totalCimbraFrontera += calcCimbraFrontera(e);

      const p = parseFloat(e.piezas) || 1;
      (e.aceros || []).forEach((a) => {
        const kg = Math.round(calcAceroItem(a).kg * p * 100) / 100;
        if (kg > 0) {
          totalAceroVarilla += kg;
          const varillaNum = parseInt(a.numVarilla || "0", 10);
          if (varillaNum > 0 && acerosDetalle[varillaNum] !== undefined) {
            acerosDetalle[varillaNum] += kg;
          }
        }
      });
    });

    const m2Parsed = parseFloat(nervaduraAnalysisM2) || 0;

    const exportToExcelConfig = () => {
      let exportHtml = `<table border="1" style="border-collapse: collapse; font-family: Arial, sans-serif;">
        <thead>
          <tr><th colspan="3" style="background-color: #0891b2; color: #fff; padding: 10px; font-size: 14pt;">Análisis de Nervadura y Losa Nervada</th></tr>
          <tr><th colspan="3" style="background-color: #cffafe; padding: 5px;">M2 Totales Analizados: ${m2Parsed.toFixed(2)}</th></tr>
          <tr>
            <th style="background-color: #f1f5f9;">Concepto</th>
            <th style="background-color: #f1f5f9;">Volumen Total</th>
            <th style="background-color: #cffafe; color: #0e7490;">Cant. por m2</th>
          </tr>
        </thead>
        <tbody>
          <tr><td>Vol. Casetones (m3)</td><td style="text-align: center">${totalCaseton.toFixed(4)}</td><td style="text-align: center; color: #0891b2">${(m2Parsed > 0 ? totalCaseton / m2Parsed : 0).toFixed(4)}</td></tr>
          <tr><td>Concreto (m3)</td><td style="text-align: center">${totalConcreto.toFixed(4)}</td><td style="text-align: center; color: #0891b2">${(m2Parsed > 0 ? totalConcreto / m2Parsed : 0).toFixed(4)}</td></tr>
          <tr><td>Cimbra (m2)</td><td style="text-align: center">${totalCimbra.toFixed(4)}</td><td style="text-align: center; color: #0891b2">${(m2Parsed > 0 ? totalCimbra / m2Parsed : 0).toFixed(4)}</td></tr>
          <tr><td>Cimbra Frontera (ml)</td><td style="text-align: center">${totalCimbraFrontera.toFixed(4)}</td><td style="text-align: center; color: #0891b2">${(m2Parsed > 0 ? totalCimbraFrontera / m2Parsed : 0).toFixed(4)}</td></tr>
          <tr><td colspan="3" style="background-color: #f1f5f9; font-weight: bold;">Acero Total por Calibre</td></tr>`;

      [3, 4, 5, 6, 8, 10].forEach((v) => {
        exportHtml += `<tr><td>Acero #${v} (kg)</td><td style="text-align: center">${acerosDetalle[v].toFixed(4)}</td><td style="text-align: center; color: #0891b2">${(m2Parsed > 0 ? acerosDetalle[v] / m2Parsed : 0).toFixed(4)}</td></tr>`;
      });

      exportHtml += `<tr style="font-weight: bold;">
        <td style="text-align: right; background-color: #f1f5f9;">Acero Total (kg)</td>
        <td style="text-align: center; background-color: #f1f5f9;">${totalAceroVarilla.toFixed(4)}</td>
        <td style="text-align: center; background-color: #cffafe; color: #0e7490;">${(m2Parsed > 0 ? totalAceroVarilla / m2Parsed : 0).toFixed(4)}</td>
      </tr>
      </tbody></table><br/>`;

      // Concreto
      exportHtml += `<table border="1" style="border-collapse: collapse; font-family: Arial, sans-serif;">
        <thead>
          <tr><th colspan="8" style="background-color: #f1f5f9; padding: 5px; font-weight: bold;">Generador de Concreto</th></tr>
          <tr style="background-color: #f8fafc; font-size: 10pt;">
            <th>Elemento</th><th>Largo</th><th>Ancho</th><th>Alto</th><th>Pzas</th>
            <th>Vol. Bruto</th><th>Vol. Casetones</th><th style="color: #0e7490;">Concreto Neto (m3)</th>
          </tr>
        </thead>
        <tbody style="font-size: 10pt;">`;
      selectedElements
        .filter((e) => (e.tipo || "").toLowerCase() !== "nervadura")
        .forEach((e) => {
          const p = parseFloat(e.piezas) || 1;
          const volBruto =
            getEffectiveLargo(e) *
            (parseFloat(e.ancho) || 0) *
            (parseFloat(e.alto) || 0) *
            p;
          const volCaseton =
            (e.tipo || "").toLowerCase() === "losa nervada"
              ? getCasetonesTotalVol(e.casetones || [], p)
              : 0;
          exportHtml += `<tr>
          <td>${e.tipo} - ${e.eje} - ${e.clave}</td>
          <td style="text-align: center">${e.largo}</td><td style="text-align: center">${e.ancho}</td>
          <td style="text-align: center">${e.alto}</td><td style="text-align: center">${e.piezas}</td>
          <td style="text-align: center">${volBruto.toFixed(4)}</td>
          <td style="text-align: center">${volCaseton.toFixed(4)}</td>
          <td style="text-align: center; font-weight: bold; color: #0891b2;">${calcConcreto(e).toFixed(4)}</td>
        </tr>`;
        });
      exportHtml += `</tbody></table><br/>`;

      // Cimbra
      exportHtml += `<table border="1" style="border-collapse: collapse; font-family: Arial, sans-serif;">
        <thead>
          <tr><th colspan="8" style="background-color: #f1f5f9; padding: 5px; font-weight: bold;">Generador de Cimbra</th></tr>
          <tr style="background-color: #f8fafc; font-size: 10pt;">
            <th>Elemento</th><th>Largo</th><th>Ancho</th><th>Alto</th><th>Pzas</th>
            <th>Descuenta</th><th style="color: #0e7490;">Cimbra (m2)</th><th style="color: #0e7490;">Cimbra Frontera (ml)</th>
          </tr>
        </thead>
        <tbody style="font-size: 10pt;">`;
      selectedElements
        .filter((e) => (e.tipo || "").toLowerCase() !== "nervadura")
        .forEach((e) => {
          exportHtml += `<tr>
          <td>${e.tipo} - ${e.eje} - ${e.clave}</td>
          <td style="text-align: center">${e.largo}</td><td style="text-align: center">${e.ancho}</td>
          <td style="text-align: center">${e.alto}</td><td style="text-align: center">${e.piezas}</td>
          <td style="text-align: center">${e.descuentaCimbra || 0}</td>
          <td style="text-align: center; font-weight: bold; color: #0891b2;">${calcCimbra(e).toFixed(4)}</td>
          <td style="text-align: center; font-weight: bold; color: #0891b2;">${calcCimbraFrontera(e).toFixed(4)}</td>
        </tr>`;
        });
      exportHtml += `</tbody></table><br/>`;

      // Casetones
      exportHtml += `<table border="1" style="border-collapse: collapse; font-family: Arial, sans-serif;">
        <thead>
          <tr><th colspan="7" style="background-color: #f1f5f9; padding: 5px; font-weight: bold;">Generador de Casetones</th></tr>
          <tr style="background-color: #f8fafc; font-size: 10pt;">
            <th>Elemento Padre</th><th>Clave Casetón</th><th>Largo</th><th>Ancho</th><th>Alto</th>
            <th>Pzas</th><th style="color: #0e7490;">Volumen Total (m3)</th>
          </tr>
        </thead>
        <tbody style="font-size: 10pt;">`;
      selectedElements.forEach((e) => {
        if ((e.tipo || "").toLowerCase() === "losa nervada" && e.casetones) {
          const pP = parseFloat(e.piezas) || 1;
          e.casetones.forEach((c) => {
            const p = parseFloat(c.piezas) || 1;
            const l = parseFloat(c.largo) || 0;
            const an = parseFloat(c.ancho) || 0;
            const al = parseFloat(c.alto) || 0;
            const vol = l * an * al * p * pP;
            exportHtml += `<tr>
              <td>${e.eje} - ${e.clave}</td>
              <td style="text-align: center">${c.clave}</td>
              <td style="text-align: center">${c.largo}</td><td style="text-align: center">${c.ancho}</td>
              <td style="text-align: center">${c.alto}</td><td style="text-align: center">${p * pP}</td>
              <td style="text-align: center; font-weight: bold; color: #0891b2;">${vol.toFixed(4)}</td>
            </tr>`;
          });
        }
      });
      exportHtml += `</tbody></table><br/>`;

      // Acero
      exportHtml += `<table border="1" style="border-collapse: collapse; font-family: Arial, sans-serif;">
        <thead>
          <tr><th colspan="9" style="background-color: #f1f5f9; padding: 5px; font-weight: bold;">Generador de Acero</th></tr>
          <tr style="background-color: #f8fafc; font-size: 10pt;">
            <th>Elemento Padre</th><th>Refuerzo</th><th># Var.</th><th>L (m)</th><th>Sep (m)</th>
            <th>Pzas</th><th>Total ML</th><th>Peso (kg/ml)</th><th style="color: #0e7490;">Total Kg</th>
          </tr>
        </thead>
        <tbody style="font-size: 10pt;">`;
      selectedElements.forEach((e) => {
        const pP = parseFloat(e.piezas) || 1;
        let acerosCorregidos = e.aceros || [];
        const validTypesOptions = getSteelTypesForElement(e.tipo);
        const validTypes = validTypesOptions.map((t) => t.toLowerCase());
        if (validTypes.length > 0) {
          acerosCorregidos = acerosCorregidos.map((a) => {
            if (!validTypes.includes((a.tipo || "").toLowerCase())) {
              return { ...a, tipo: validTypesOptions[0] };
            }
            return a;
          });
        }
        acerosCorregidos.forEach((a) => {
          const calc = calcAceroItem(a);
          const p = parseFloat(a.piezas) || 1;
          const kgl = PESOS_VARILLA[a.numVarilla || "-"] || 0;
          exportHtml += `<tr>
            <td>${e.eje} - ${e.clave}</td>
            <td>${a.tipo}</td>
            <td style="text-align: center">#${a.numVarilla}</td>
            <td style="text-align: center">${calc.mlPorPieza.toFixed(2)}</td>
            <td style="text-align: center">${a.separacion || "-"}</td>
            <td style="text-align: center">${p * pP}</td>
            <td style="text-align: center">${(calc.ml * pP).toFixed(2)}</td>
            <td style="text-align: center; mso-number-format:'0\\.000';">${kgl.toFixed(3)}</td>
            <td style="text-align: center; font-weight: bold; color: #0891b2;">${(Math.round(calc.kg * pP * 100) / 100).toFixed(4)}</td>
          </tr>`;
        });
      });
      exportHtml += `</tbody></table>`;

      exportFormattedExcel(exportHtml, `Analisis_Nervadura_M2`);
    };

    return (
      <div className="bg-white rounded-3xl shadow-xl border border-gray-100 animate-in fade-in zoom-in-95 duration-300 flex flex-col mb-8">
        <div className="bg-cyan-800 text-white p-4 flex justify-between items-center shrink-0 rounded-t-3xl">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-cyan-700/50 rounded-lg border border-cyan-500/50">
              <Calculator size={20} className="text-cyan-100" />
            </div>
            <div>
              <h2 className="text-lg font-black uppercase tracking-wider leading-tight">
                Análisis / m2 Nervadura y Losa Nervada
              </h2>
              <p className="text-[9px] text-cyan-300 font-bold uppercase tracking-widest">
                {activePartida?.nombre}
              </p>
            </div>
          </div>
          <button
            onClick={() => setShowAnalysisNervadura(false)}
            className="p-1.5 bg-cyan-700 hover:bg-cyan-600 rounded-full transition-colors"
          >
            <X size={20} />
          </button>
        </div>

        <div className="p-4 grid grid-cols-1 xl:grid-cols-2 gap-6 bg-slate-50/50 min-h-[400px]">
          <div className="flex flex-col border border-slate-200 rounded-2xl overflow-hidden bg-white shadow-sm">
            <div className="p-3 bg-slate-100 border-b border-slate-200 flex justify-between items-center">
              <h3 className="font-black text-slate-700 uppercase text-xs tracking-widest">
                Elementos Disponibles
              </h3>
              {availableNervaduras.length > 0 && (
                <button
                  onClick={toggleAll}
                  className="text-[10px] font-bold text-cyan-600 uppercase hover:underline"
                >
                  {nervaduraAnalysisSelection.length ===
                  availableNervaduras.length
                    ? "Deseleccionar Todos"
                    : "Seleccionar Todos"}
                </button>
              )}
            </div>
            <div className="overflow-y-auto max-h-[400px] p-2">
              {availableNervaduras.length === 0 ? (
                <div className="p-8 text-center text-slate-400 font-bold uppercase text-xs">
                  No hay losas nervadas o nervaduras capturadas en este nivel.
                </div>
              ) : (
                <div className="flex flex-col gap-1">
                  {availableNervaduras.map((e) => (
                    <div
                      key={e.id}
                      onClick={() => toggleSelection(e.id)}
                      className={`p-2 border rounded-lg cursor-pointer flex items-center gap-3 transition-colors ${nervaduraAnalysisSelection.includes(e.id) ? "bg-cyan-50 border-cyan-200 shadow-sm" : "hover:bg-slate-50 border-transparent hover:border-slate-200"}`}
                    >
                      <button
                        className={`w-5 h-5 rounded flex items-center justify-center shrink-0 ${nervaduraAnalysisSelection.includes(e.id) ? "bg-cyan-600 text-white" : "bg-slate-200 text-transparent"}`}
                      >
                        <CheckSquare size={14} />
                      </button>
                      <div className="flex-1 text-xs">
                        <p className="font-bold text-slate-800">
                          <span className="text-cyan-700">
                            {e.tipo.toUpperCase()}
                          </span>{" "}
                          - {e.eje} - {e.clave}
                        </p>
                        <p className="text-[10px] text-slate-500 uppercase tracking-wider">
                          L: {e.largo} | A: {e.ancho} | H: {e.alto} | Pzas:{" "}
                          {e.piezas}
                        </p>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>

          <div className="flex flex-col gap-4">
            <div className="bg-white border text-center border-slate-200 rounded-2xl p-4 shadow-sm">
              <label className="text-xs font-black text-slate-700 uppercase tracking-widest mb-2 block">
                M2 Totales del Análisis
              </label>
              <input
                type="number"
                value={nervaduraAnalysisM2}
                onChange={(e) => setNervaduraAnalysisM2(e.target.value)}
                placeholder="0.00"
                className="w-full text-center text-xl font-bold p-3 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-cyan-500 focus:border-cyan-500 outline-none"
              />
            </div>

            <div className="bg-white border border-slate-200 rounded-2xl flex-1 shadow-sm flex flex-col overflow-hidden">
              <div className="p-3 bg-slate-800 text-white flex justify-between items-center shrink-0">
                <h3 className="font-black uppercase text-xs tracking-widest">
                  Resultados del Análisis
                </h3>
                <button
                  onClick={exportToExcelConfig}
                  disabled={selectedElements.length === 0}
                  className={`flex items-center gap-2 px-3 py-1.5 rounded-lg text-xs font-bold uppercase transition ${selectedElements.length > 0 ? "bg-emerald-600 hover:bg-emerald-500 text-white" : "bg-slate-700 text-slate-400 cursor-not-allowed"}`}
                >
                  <FileSpreadsheet size={14} />
                  Exportar a Excel
                </button>
              </div>

              <div className="p-4 flex-1 overflow-y-auto flex flex-col gap-8">
                <div>
                  <h4 className="font-black text-slate-700 uppercase text-[10px] mb-2 tracking-widest border-b pb-1">
                    Resumen General
                  </h4>
                  <table className="w-full text-left border-collapse text-[13px]">
                    <thead className="bg-slate-100 text-slate-500 font-bold uppercase text-[9px]">
                      <tr>
                        <th className="p-2 border-b border-r">Concepto</th>
                        <th className="p-2 border-b border-r text-center">
                          Volumen Total
                        </th>
                        <th className="p-2 border-b text-center text-cyan-700">
                          Cant. por m2
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 font-bold text-slate-700 text-[13px]">
                      <tr className="hover:bg-slate-50">
                        <td className="p-2 border-r">Vol. Casetones (m3)</td>
                        <td className="p-2 border-r text-center">
                          {totalCaseton.toFixed(4)}
                        </td>
                        <td className="p-2 text-center text-cyan-600 bg-cyan-50/30">
                          {(m2Parsed > 0 ? totalCaseton / m2Parsed : 0).toFixed(
                            4,
                          )}
                        </td>
                      </tr>
                      <tr className="hover:bg-slate-50">
                        <td className="p-2 border-r">Concreto (m3)</td>
                        <td className="p-2 border-r text-center">
                          {totalConcreto.toFixed(4)}
                        </td>
                        <td className="p-2 text-center text-cyan-600 bg-cyan-50/30">
                          {(m2Parsed > 0
                            ? totalConcreto / m2Parsed
                            : 0
                          ).toFixed(4)}
                        </td>
                      </tr>
                      <tr className="hover:bg-slate-50">
                        <td className="p-2 border-r">Cimbra (m2)</td>
                        <td className="p-2 border-r text-center">
                          {totalCimbra.toFixed(4)}
                        </td>
                        <td className="p-2 text-center text-cyan-600 bg-cyan-50/30">
                          {(m2Parsed > 0 ? totalCimbra / m2Parsed : 0).toFixed(
                            4,
                          )}
                        </td>
                      </tr>
                      <tr className="hover:bg-slate-50">
                        <td className="p-2 border-r">Cimbra Frontera (ml)</td>
                        <td className="p-2 border-r text-center">
                          {totalCimbraFrontera.toFixed(4)}
                        </td>
                        <td className="p-2 text-center text-cyan-600 bg-cyan-50/30">
                          {(m2Parsed > 0
                            ? totalCimbraFrontera / m2Parsed
                            : 0
                          ).toFixed(4)}
                        </td>
                      </tr>
                      <tr className="hover:bg-slate-50">
                        <td className="p-2 border-r bg-slate-50" colSpan={3}>
                          Acero Total por Calibre
                        </td>
                      </tr>
                      {[3, 4, 5, 6, 8, 10].map((v) => (
                        <tr
                          key={`v-${v}`}
                          className="hover:bg-slate-50"
                        >
                          <td className="p-2 border-r pl-6">Acero #{v} (kg)</td>
                          <td className="p-2 border-r text-center">
                            {acerosDetalle[v].toFixed(4)}
                          </td>
                          <td className="p-2 text-center text-cyan-600 bg-cyan-50/30">
                            {(m2Parsed > 0
                              ? acerosDetalle[v] / m2Parsed
                              : 0
                            ).toFixed(4)}
                          </td>
                        </tr>
                      ))}
                      <tr className="bg-slate-100 font-black">
                        <td className="p-2 border-r text-right">
                          Acero Total (kg)
                        </td>
                        <td className="p-2 border-r text-center">
                          {totalAceroVarilla.toFixed(4)}
                        </td>
                        <td className="p-2 text-center text-cyan-700 bg-cyan-100/50">
                          {(m2Parsed > 0
                            ? totalAceroVarilla / m2Parsed
                            : 0
                          ).toFixed(4)}
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                <div className="overflow-x-auto">
                  <h4 className="font-black text-slate-700 uppercase text-[10px] mb-2 tracking-widest border-b pb-1">
                    Generador de Concreto
                  </h4>
                  <table className="w-full text-left border-collapse text-[10px]">
                    <thead className="bg-slate-100 text-slate-500 uppercase text-[9px]">
                      <tr>
                        <th className="p-1 border">Elemento</th>
                        <th className="p-1 border text-center">Largo</th>
                        <th className="p-1 border text-center">Ancho</th>
                        <th className="p-1 border text-center">Alto</th>
                        <th className="p-1 border text-center">Pzas</th>
                        <th className="p-1 border text-center">Vol. Bruto</th>
                        <th className="p-1 border text-center">
                          Vol. Casetones
                        </th>
                        <th className="p-1 border text-center text-cyan-700">
                          Concreto Neto (m3)
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y text-slate-700 font-medium">
                      {selectedElements
                        .filter(
                          (e) => (e.tipo || "").toLowerCase() !== "nervadura",
                        )
                        .map((e, i) => {
                          const p = parseFloat(e.piezas) || 1;
                          const volBruto =
                            getEffectiveLargo(e) *
                            (parseFloat(e.ancho) || 0) *
                            (parseFloat(e.alto) || 0) *
                            p;
                          const volCaseton =
                            (e.tipo || "").toLowerCase() === "losa nervada"
                              ? getCasetonesTotalVol(e.casetones || [], p)
                              : 0;
                          return (
                            <tr key={i} className="hover:bg-slate-50">
                              <td className="p-1 border">
                                {e.tipo} - {e.eje} - {e.clave}
                              </td>
                              <td className="p-1 border text-center">
                                {e.largo}
                              </td>
                              <td className="p-1 border text-center">
                                {e.ancho}
                              </td>
                              <td className="p-1 border text-center">
                                {e.alto}
                              </td>
                              <td className="p-1 border text-center">
                                {e.piezas}
                              </td>
                              <td className="p-1 border text-center">
                                {volBruto.toFixed(4)}
                              </td>
                              <td className="p-1 border text-center">
                                {volCaseton.toFixed(4)}
                              </td>
                              <td className="p-1 border text-center text-cyan-700 font-bold bg-cyan-50/30">
                                {calcConcreto(e).toFixed(4)}
                              </td>
                            </tr>
                          );
                        })}
                    </tbody>
                  </table>
                </div>

                <div className="overflow-x-auto">
                  <h4 className="font-black text-slate-700 uppercase text-[10px] mb-2 tracking-widest border-b pb-1">
                    Generador de Cimbra
                  </h4>
                  <table className="w-full text-left border-collapse text-[10px]">
                    <thead className="bg-slate-100 text-slate-500 uppercase text-[9px]">
                      <tr>
                        <th className="p-1 border">Elemento</th>
                        <th className="p-1 border text-center">Largo</th>
                        <th className="p-1 border text-center">Ancho</th>
                        <th className="p-1 border text-center">Alto</th>
                        <th className="p-1 border text-center">Pzas</th>
                        <th className="p-1 border text-center">Descuenta</th>
                        <th className="p-1 border text-center text-cyan-700">
                          Cimbra (m2)
                        </th>
                        <th className="p-1 border text-center text-cyan-700">
                          Cimbra Frontera (ml)
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y text-slate-700 font-medium">
                      {selectedElements
                        .filter(
                          (e) => (e.tipo || "").toLowerCase() !== "nervadura",
                        )
                        .map((e, i) => (
                          <tr key={i} className="hover:bg-slate-50">
                            <td className="p-1 border">
                              {e.tipo} - {e.eje} - {e.clave}
                            </td>
                            <td className="p-1 border text-center">
                              {e.largo}
                            </td>
                            <td className="p-1 border text-center">
                              {e.ancho}
                            </td>
                            <td className="p-1 border text-center">{e.alto}</td>
                            <td className="p-1 border text-center">
                              {e.piezas}
                            </td>
                            <td className="p-1 border text-center">
                              {e.descuentaCimbra || 0}
                            </td>
                            <td className="p-1 border text-center font-bold text-cyan-700 bg-cyan-50/30">
                              {calcCimbra(e).toFixed(4)}
                            </td>
                            <td className="p-1 border text-center font-bold text-cyan-700 bg-cyan-50/30">
                              {calcCimbraFrontera(e).toFixed(4)}
                            </td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                </div>

                <div className="overflow-x-auto">
                  <h4 className="font-black text-slate-700 uppercase text-[10px] mb-2 tracking-widest border-b pb-1">
                    Generador de Casetones
                  </h4>
                  <table className="w-full text-left border-collapse text-[10px]">
                    <thead className="bg-slate-100 text-slate-500 uppercase text-[9px]">
                      <tr>
                        <th className="p-1 border">Elemento Padre</th>
                        <th className="p-1 border text-center">Casetón</th>
                        <th className="p-1 border text-center">Largo</th>
                        <th className="p-1 border text-center">Ancho</th>
                        <th className="p-1 border text-center">Alto</th>
                        <th className="p-1 border text-center">Pzas</th>
                        <th className="p-1 border text-center text-cyan-700">
                          Volumen Total (m3)
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y text-slate-700 font-medium">
                      {selectedElements.filter(
                        (e) =>
                          (e.tipo || "").toLowerCase() === "losa nervada" &&
                          e.casetones,
                      ).length === 0 && (
                        <tr>
                          <td
                            colSpan={7}
                            className="p-2 text-center text-slate-400"
                          >
                            Sin casetones en elementos seleccionados
                          </td>
                        </tr>
                      )}
                      {selectedElements
                        .filter(
                          (e) =>
                            (e.tipo || "").toLowerCase() === "losa nervada" &&
                            e.casetones,
                        )
                        .map((e) => {
                          const pP = parseFloat(e.piezas) || 1;
                          return e.casetones.map((c, i) => {
                            const p = parseFloat(c.piezas) || 1;
                            const l = parseFloat(c.largo) || 0;
                            const an = parseFloat(c.ancho) || 0;
                            const al = parseFloat(c.alto) || 0;
                            const vol = l * an * al * p * pP;
                            return (
                              <tr
                                key={`${e.id}-${i}`}
                                className="hover:bg-slate-50"
                              >
                                <td className="p-1 border">
                                  {e.eje} - {e.clave}
                                </td>
                                <td className="p-1 border text-center">
                                  {c.clave}
                                </td>
                                <td className="p-1 border text-center">
                                  {c.largo}
                                </td>
                                <td className="p-1 border text-center">
                                  {c.ancho}
                                </td>
                                <td className="p-1 border text-center">
                                  {c.alto}
                                </td>
                                <td className="p-1 border text-center">
                                  {p * pP}
                                </td>
                                <td className="p-1 border text-center font-bold text-cyan-700 bg-cyan-50/30">
                                  {vol.toFixed(4)}
                                </td>
                              </tr>
                            );
                          });
                        })}
                    </tbody>
                  </table>
                </div>

                <div className="overflow-x-auto">
                  <h4 className="font-black text-slate-700 uppercase text-[10px] mb-2 tracking-widest border-b pb-1">
                    Generador de Acero
                  </h4>
                  <table className="w-full text-left border-collapse text-[10px]">
                    <thead className="bg-slate-100 text-slate-500 uppercase text-[9px]">
                      <tr>
                        <th className="p-1 border">Elemento Padre</th>
                        <th className="p-1 border text-center">Refuerzo</th>
                        <th className="p-1 border text-center"># Var.</th>
                        <th className="p-1 border text-center">L (m)</th>
                        <th className="p-1 border text-center">Sep (m)</th>
                        <th className="p-1 border text-center">Pzas</th>
                        <th className="p-1 border text-center">Total ML</th>
                        <th className="p-1 border text-center">Peso (kg/ml)</th>
                        <th className="p-1 border text-center text-cyan-700">
                          Total KG
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y text-slate-700 font-medium">
                      {selectedElements.filter(
                        (e) => e.aceros && e.aceros.length > 0,
                      ).length === 0 && (
                        <tr>
                          <td
                            colSpan={9}
                            className="p-2 text-center text-slate-400"
                          >
                            Sin acero en elementos seleccionados
                          </td>
                        </tr>
                      )}
                      {selectedElements.map((e) => {
                        const pP = parseFloat(e.piezas) || 1;
                        let acerosCorregidos = e.aceros || [];
                        const validTypesOptions = getSteelTypesForElement(
                          e.tipo,
                        );
                        const validTypes = validTypesOptions.map((t) =>
                          t.toLowerCase(),
                        );
                        if (validTypes.length > 0) {
                          acerosCorregidos = acerosCorregidos.map((a) => {
                            if (
                              !validTypes.includes((a.tipo || "").toLowerCase())
                            ) {
                              return { ...a, tipo: validTypesOptions[0] };
                            }
                            return a;
                          });
                        }
                        return acerosCorregidos.map((a, i) => {
                          const calc = calcAceroItem(a);
                          const p = parseFloat(a.piezas) || 1;
                          const kgl = PESOS_VARILLA[a.numVarilla || "-"] || 0;
                          return (
                            <tr
                              key={`${e.id}-${i}`}
                              className="hover:bg-slate-50"
                            >
                              <td className="p-1 border">
                                {e.eje} - {e.clave}
                              </td>
                              <td className="p-1 border text-center">
                                {a.tipo}
                              </td>
                              <td className="p-1 border text-center">
                                #{a.numVarilla}
                              </td>
                              <td className="p-1 border text-center">
                                {calc.mlPorPieza.toFixed(2)}
                              </td>
                              <td className="p-1 border text-center">
                                {a.separacion || "-"}
                              </td>
                              <td className="p-1 border text-center">
                                {p * pP}
                              </td>
                              <td className="p-1 border text-center">
                                {(calc.ml * pP).toFixed(2)}
                              </td>
                              <td className="p-1 border text-center">
                                {kgl.toFixed(3)}
                              </td>
                              <td className="p-1 border text-center font-bold text-cyan-700 bg-cyan-50/30">
                                {(Math.round(calc.kg * pP * 100) / 100).toFixed(
                                  4,
                                )}
                              </td>
                            </tr>
                          );
                        });
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  };

  const renderMetalStructureGenerator = () => {
    const metalRows = metalEstructuras.filter(
      (r) => r.tipo === activeMetalSubmodal,
    );

    const renderGlobalMetalReport = () => {
      const METAL_ORDER = {
        "Perfil HSS Cuadrado": 1,
        "Perfil HSS Rectangular": 2,
        "Viga IPR": 3,
        "Viga IR": 4,
        "Lámina Galvanizada": 5,
        "Canales Tipo C": 6,
        "Polín Monten": 7,
        "Perfil PTR": 8,
        "Perfiles OC": 9,
        "Placas Metálicas": 99,
      };

      const sortedMetalEstructuras = [...metalEstructuras].sort((a, b) => {
        return (METAL_ORDER[a.tipo] || 50) - (METAL_ORDER[b.tipo] || 50);
      });

      let grandTotal = 0;

      let htmlForExport = `<table style="border-collapse: collapse; font-family: Arial, sans-serif; text-align: left;" border="1">
          <thead>
            <tr style="background-color: #2f5496; color: white; font-size: 14pt; font-weight: bold; text-align: center; height: 30pt;">
              <th colspan="11">REPORTE GLOBAL DE ESTRUCTURA METÁLICA</th>
            </tr>
            <tr style="background-color: #f2f6fb; color: #2f5496; font-size: 8pt; font-weight: bold; text-align: center; height: 15pt;">
              <th>Eje</th>
              <th>Clave</th>
              <th>Tipo</th>
              <th>Pzas</th>
              <th>Largo (m)</th>
              <th>Ancho (m)</th>
              <th>Medida/Espesor</th>
              <th>Calibre/Peso</th>
              <th>Peso ML/M2</th>
              <th>Peso/pza</th>
              <th>Peso Total (kg)</th>
            </tr>
          </thead>
          <tbody>`;

      return (
        <div className="flex-1 flex flex-col overflow-hidden bg-slate-50 relative">
          <div className="flex-1 p-6 overflow-y-auto">
            <h3 className="text-xl font-black text-slate-800 uppercase mb-4 flex items-center justify-between">
              <span>Reporte Global de Estructura Metálica</span>
              <button
                onClick={() => {
                  let exportHtml = htmlForExport;
                  sortedMetalEstructuras.forEach((r) => {
                    const { pzas, p, w } = getMetalRowWeights(r);
                    const isArea =
                      r.tipo === "Placas Metálicas" ||
                      r.tipo === "Lámina Galvanizada";
                    exportHtml += `<tr style="font-size: 8pt; text-align: center;">
                      <td>${r.eje || "-"}</td>
                      <td>${r.clave || "-"}</td>
                      <td>${r.tipo}</td>
                      <td>${pzas}</td>
                      <td>${(parseFloat(r.largo) || 0).toFixed(2)}</td>
                      <td>${isArea ? (parseFloat(r.ancho) || 1).toFixed(2) : "-"}</td>
                      <td>${r.tipo === "Viga IR" ? "-" : r.medidaNominal || "-"}</td>
                      <td>${r.calibre || "-"}</td>
                      <td>${p.toFixed(3)}</td>
                      <td>${pzas > 0 ? (w / pzas).toFixed(2) : "0.00"}</td>
                      <td style="font-weight: bold; color: #1e40af;">${w.toFixed(2)}</td>
                    </tr>`;
                  });
                  exportHtml += `</tbody></table>`;
                  const finalTotal = sortedMetalEstructuras.reduce(
                    (acc, r) => acc + getMetalRowWeights(r).w,
                    0,
                  );
                  const descalibre = finalTotal * 0.04;
                  exportHtml += `<br/><table style="border-collapse: collapse; font-family: Arial, sans-serif; text-align: left;" border="1">
                    <tbody>
                      <tr><th style="background-color: #deebf7; color: #2f5496;">Subtotal Estructura Metálica</th><td style="font-weight: bold;">${finalTotal.toFixed(2)} kg</td></tr>
                      <tr><th style="background-color: #fdf3e8; color: #c65911;">Descalibre/Desperdicio (4%)</th><td style="font-weight: bold;">${descalibre.toFixed(2)} kg</td></tr>
                      <tr><th style="background-color: #deebf7; color: #2f5496;">Total Estructura Metálica</th><td style="font-weight: bold;">${(finalTotal + descalibre).toFixed(2)} kg</td></tr>
                    </tbody>
                  </table>`;
                  exportFormattedExcel(
                    exportHtml,
                    `Reporte_Global_Estructura_Metalica`,
                  );
                }}
                className="px-4 py-2 bg-emerald-600 text-white text-xs font-black rounded-lg uppercase tracking-wider flex items-center gap-2 hover:bg-emerald-500 transition-colors shadow"
              >
                Exportar Excel
              </button>
            </h3>

            <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden mb-6">
              <table className="w-full text-left text-[10px] sm:text-xs">
                <thead className="bg-slate-100 text-slate-700 font-black uppercase tracking-wider">
                  <tr>
                    <th className="p-3 border-b text-center">No.</th>
                    <th className="p-3 border-b text-center">Eje</th>
                    <th className="p-3 border-b text-center">Clave</th>
                    <th className="p-3 border-b text-center">Tipo Elemento</th>
                    <th className="p-3 border-b text-center">Pzas</th>
                    <th className="p-3 border-b text-center">Largo (m)</th>
                    <th className="p-3 border-b text-center">Ancho (m)</th>
                    <th className="p-3 border-b text-center">Medida</th>
                    <th className="p-3 border-b text-center">Calibre</th>
                    <th className="p-3 border-b text-center">Peso ML/M2</th>
                    <th className="p-3 border-b text-center bg-blue-50 text-blue-900">
                      Peso Total (kg)
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {sortedMetalEstructuras.length === 0 && (
                    <tr>
                      <td
                        colSpan={11}
                        className="p-6 text-center text-slate-400 font-bold uppercase tracking-widest"
                      >
                        Aún no hay elementos metálicos registrados en este
                        nivel.
                      </td>
                    </tr>
                  )}
                  {sortedMetalEstructuras.map((r, i) => {
                    const { pzas, p, w } = getMetalRowWeights(r);
                    grandTotal += w;
                    return (
                      <tr key={r.id} className="hover:bg-amber-50">
                        <td className="p-3 text-center text-slate-400 font-bold">
                          {i + 1}
                        </td>
                        <td className="p-3 text-center font-bold text-slate-700">
                          {r.eje || "-"}
                        </td>
                        <td className="p-3 text-center font-bold text-slate-700">
                          {r.clave || "-"}
                        </td>
                        <td className="p-3 text-center font-bold text-amber-700">
                          {r.tipo}
                        </td>
                        <td className="p-3 text-center font-bold text-slate-700">
                          {pzas}
                        </td>
                        <td className="p-3 text-center font-bold text-slate-700">
                          {(parseFloat(r.largo) || 0).toFixed(2)}
                        </td>
                        <td className="p-3 text-center font-bold text-slate-700">
                          {r.tipo === "Placas Metálicas" ||
                          r.tipo === "Lámina Galvanizada"
                            ? (parseFloat(r.ancho) || 1).toFixed(2)
                            : "-"}
                        </td>
                        <td className="p-3 text-center font-bold text-slate-600">
                          {r.tipo === "Viga IR" ? "-" : r.medidaNominal || "-"}
                        </td>
                        <td className="p-3 text-center font-bold text-slate-600">
                          {r.calibre || "-"}
                        </td>
                        <td className="p-3 text-center font-black text-slate-500">
                          {p.toFixed(3)}
                        </td>
                        <td className="p-3 text-center font-black text-blue-700 bg-blue-50/30">
                          {w.toFixed(2)}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>

          <div className="bg-slate-100 p-6 flex flex-col md:flex-row gap-4 items-end justify-end border-t border-slate-200">
            <div className="bg-white rounded-2xl p-4 px-6 border border-slate-200 shadow-sm flex flex-col min-w-[250px]">
              <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest mb-1">
                Subtotal Acero Estructural
              </span>
              <div className="flex items-baseline gap-2 justify-end">
                <span className="text-2xl font-black text-slate-700 leading-none">
                  {grandTotal.toFixed(2)}
                </span>
                <span className="text-xs font-black text-slate-400 uppercase">
                  kg
                </span>
              </div>
            </div>
            <div className="bg-orange-50 rounded-2xl p-4 px-6 border border-orange-200 shadow-sm flex flex-col min-w-[250px]">
              <span className="text-[10px] font-black text-orange-600 uppercase tracking-widest mb-1">
                Descalibre/Consumo Adicional (4%)
              </span>
              <div className="flex items-baseline gap-2 justify-end">
                <span className="text-2xl font-black text-orange-700 leading-none">
                  {(grandTotal * 0.04).toFixed(2)}
                </span>
                <span className="text-xs font-black text-orange-400 uppercase">
                  kg
                </span>
              </div>
            </div>
            <div className="bg-blue-600 rounded-2xl p-4 px-6 border border-blue-700 shadow-lg flex flex-col min-w-[250px]">
              <span className="text-[10px] font-black text-blue-200 uppercase tracking-widest mb-1">
                Total (+ 4% Merma)
              </span>
              <div className="flex items-baseline gap-2 justify-end">
                <span className="text-3xl font-black text-white leading-none">
                  {(grandTotal * 1.04).toFixed(2)}
                </span>
                <span className="text-sm font-black text-blue-300 uppercase">
                  kg
                </span>
              </div>
            </div>
          </div>
        </div>
      );
    };

    const updateMetalRow = (idx, fieldOrObj, value) => {
      const realIdx = metalEstructuras.findIndex(
        (r) => r.id === metalRows[idx].id,
      );
      if (realIdx === -1) return;
      const copy = [...metalEstructuras];

      if (typeof fieldOrObj === "object") {
        copy[realIdx] = { ...copy[realIdx], ...fieldOrObj };
      } else {
        copy[realIdx] = { ...copy[realIdx], [fieldOrObj]: value };
      }

      updateActiveMetalEstructuras(copy);
    };

    const addMetalRow = () => {
      let defaultMedida = '4"';
      let defaultCalibre = "14";
      let defaultAncho = "";
      if (activeMetalSubmodal === "Viga IPR") {
        defaultMedida = '4" X 4"';
        defaultCalibre = "19.3";
      } else if (activeMetalSubmodal === "Perfil HSS Cuadrado") {
        defaultMedida = '4" x 4"';
        defaultCalibre = '1/8"';
      } else if (activeMetalSubmodal === "Perfil HSS Rectangular") {
        defaultMedida = '8" x 4"';
        defaultCalibre = '3/16"';
      } else if (activeMetalSubmodal === "Perfil PTR") {
        defaultMedida = '1" x 1"';
        defaultCalibre = "14";
      } else if (activeMetalSubmodal === "Placas Metálicas") {
        defaultMedida = "-";
        defaultCalibre = '1/4"';
        defaultAncho = "1";
      } else if (activeMetalSubmodal === "Lámina Galvanizada") {
        defaultMedida = "R-72";
        defaultCalibre = "26";
        defaultAncho = "1";
      } else if (activeMetalSubmodal === "Canales Tipo C") {
        defaultMedida = '4"';
        defaultCalibre = "8.04";
      } else if (activeMetalSubmodal === "Perfiles OC") {
        defaultMedida = '20"';
        defaultCalibre = '0.500"';
      } else if (activeMetalSubmodal === "Viga IR") {
        defaultMedida = "152";
        defaultCalibre = "12.7";
      }
      updateActiveMetalEstructuras([
        ...metalEstructuras,
        {
          id: `ME-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
          tipo: activeMetalSubmodal,
          eje: "",
          clave: "",
          pzas: "1",
          largo: "0",
          ancho: defaultAncho,
          medidaNominal: defaultMedida,
          calibre: defaultCalibre,
        },
      ]);
    };

    const deleteMetalRow = (id) => {
      updateActiveMetalEstructuras(metalEstructuras.filter((r) => r.id !== id));
    };

    let totalWeight = 0;
    let totalWeightPolin = 0;
    let totalWeightIPR = 0;
    let totalWeightHSS = 0;
    let totalWeightHSSRect = 0;
    let totalWeightPlacas = 0;
    let totalWeightPTR = 0;
    let totalWeightLamina = 0;
    let totalWeightCanalesC = 0;
    let totalWeightPerfilesOC = 0;
    let totalWeightVigaIR = 0;
    metalRows.forEach((r) => {
      let pzas = parseFloat(r.pzas);
      if (isNaN(pzas)) pzas = 1;

      let ancho = parseFloat(r.ancho);
      if (isNaN(ancho)) ancho = 1; // used for m2 in Placas

      if (r.tipo === "Polín Monten") {
        const pd = POLIN_DATA[r.medidaNominal];
        const p = pd?.pesos[r.calibre] || 0;
        const w = (parseFloat(r.largo) || 0) * p * pzas;
        totalWeight += w;
        totalWeightPolin += w;
      } else if (r.tipo === "Viga IPR") {
        const pdList = IPR_DATA[r.medidaNominal] || [];
        const item = pdList.find((x) => x.kg.toString() === r.calibre) ||
          pdList[0] || { kg: 0 };
        const w = (parseFloat(r.largo) || 0) * item.kg * pzas;
        totalWeight += w;
        totalWeightIPR += w;
      } else if (r.tipo === "Perfil HSS Cuadrado") {
        const pd = HSS_DATA[r.medidaNominal];
        const p = pd?.pesos[r.calibre] || 0;
        const w = (parseFloat(r.largo) || 0) * p * pzas;
        totalWeight += w;
        totalWeightHSS += w;
      } else if (r.tipo === "Perfil HSS Rectangular") {
        const pd = HSS_RECT_DATA[r.medidaNominal];
        const p = pd?.pesos[r.calibre] || 0;
        const w = (parseFloat(r.largo) || 0) * p * pzas;
        totalWeight += w;
        totalWeightHSSRect += w;
      } else if (r.tipo === "Placas Metálicas") {
        // Here calibre stores the espesor ('1"'), and 'largo' & 'ancho' calculate m2
        const p = PLACAS_DATA[r.calibre] || 0;
        const m2 = (parseFloat(r.largo) || 0) * ancho;
        const w = m2 * p * pzas;
        totalWeight += w;
        totalWeightPlacas += w;
      } else if (r.tipo === "Lámina Galvanizada") {
        const pd = LAMINA_DATA[r.medidaNominal];
        const p = pd?.pesos[r.calibre] || 0;
        const m2 = (parseFloat(r.largo) || 0) * ancho;
        const w = m2 * p * pzas;
        totalWeight += w;
        totalWeightLamina += w;
      } else if (r.tipo === "Perfil PTR") {
        const pd = PTR_DATA[r.medidaNominal];
        const p = pd?.pesos[r.calibre] || 0;
        const w = (parseFloat(r.largo) || 0) * p * pzas;
        totalWeight += w;
        totalWeightPTR += w;
      } else if (r.tipo === "Canales Tipo C") {
        const pd = CANALES_C_DATA[r.medidaNominal];
        const p = pd?.pesos[r.calibre] || 0;
        const w = (parseFloat(r.largo) || 0) * p * pzas;
        totalWeight += w;
        totalWeightCanalesC += w;
      } else if (r.tipo === "Perfiles OC") {
        const pd = PERFILES_OC_DATA[r.medidaNominal];
        const p = pd?.pesos[r.calibre] || 0;
        const w = (parseFloat(r.largo) || 0) * p * pzas;
        totalWeight += w;
        totalWeightPerfilesOC += w;
      } else if (r.tipo === "Viga IR") {
        const p = parseFloat(r.calibre) || 0;
        const w = (parseFloat(r.largo) || 0) * p * pzas;
        totalWeight += w;
        totalWeightVigaIR += w;
      }
    });

    return (
      <div className="bg-white rounded-3xl shadow-xl border border-gray-100 animate-in fade-in zoom-in-95 duration-300 flex flex-col min-h-[600px]">
        <div className="bg-red-800 text-white p-4 flex justify-between items-center shrink-0 rounded-t-3xl">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-red-700/50 rounded-lg border border-red-600/50">
              <Wrench size={20} className="text-red-200" />
            </div>
            <div>
              <h2 className="text-lg font-black uppercase tracking-wider leading-tight">
                Estructura Metálica
              </h2>
              <p className="text-[9px] text-red-300 font-bold uppercase tracking-widest">
                {activePartida.nombre}
              </p>
            </div>
          </div>
          <div className="flex items-center gap-3">
            <select
              value={activeLevelId}
              onChange={(e) => {
                setActiveLevelId(e.target.value);
                setSelectedMetalRows([]);
              }}
              className="bg-[#450a0a] text-white font-black text-[10px] uppercase p-2 px-4 rounded-lg border border-red-600 outline-none cursor-pointer"
            >
              {niveles.map((n) => (
                <option key={n.id} value={n.id} className="text-slate-800">
                  {n.nombre}
                </option>
              ))}
            </select>
            <button
              onClick={() => setShowMetalStructureGenerator(false)}
              className="p-1.5 bg-red-700 hover:bg-red-600 rounded-full transition-colors"
            >
              <X size={20} />
            </button>
          </div>
        </div>

        <div className="flex flex-1 overflow-hidden">
          <div className="w-48 bg-slate-50 border-r border-slate-200 p-2 flex flex-col gap-2 shrink-0">
            <button
              onClick={() => setActiveMetalSubmodal("Polín Monten")}
              className={`w-full text-left p-3 rounded-xl font-black text-xs uppercase cursor-pointer transition-colors ${activeMetalSubmodal === "Polín Monten" ? "bg-red-600 text-white" : "bg-white text-slate-600 hover:bg-red-50 border border-slate-200"}`}
            >
              Polín Monten
            </button>
            <button
              onClick={() => setActiveMetalSubmodal("Viga IPR")}
              className={`w-full text-left p-3 rounded-xl font-black text-xs uppercase cursor-pointer transition-colors ${activeMetalSubmodal === "Viga IPR" ? "bg-red-600 text-white" : "bg-white text-slate-600 hover:bg-red-50 border border-slate-200"}`}
            >
              Viga IPR
            </button>
            <button
              onClick={() => setActiveMetalSubmodal("Perfil HSS Cuadrado")}
              className={`w-full text-left p-3 rounded-xl font-black text-xs uppercase cursor-pointer transition-colors ${activeMetalSubmodal === "Perfil HSS Cuadrado" ? "bg-red-600 text-white" : "bg-white text-slate-600 hover:bg-red-50 border border-slate-200"}`}
            >
              Perfil HSS Cuad.
            </button>
            <button
              onClick={() => setActiveMetalSubmodal("Perfil HSS Rectangular")}
              className={`w-full text-left p-3 rounded-xl font-black text-xs uppercase cursor-pointer transition-colors ${activeMetalSubmodal === "Perfil HSS Rectangular" ? "bg-red-600 text-white" : "bg-white text-slate-600 hover:bg-red-50 border border-slate-200"}`}
            >
              Perfil HSS Rect.
            </button>
            <button
              onClick={() => setActiveMetalSubmodal("Perfil PTR")}
              className={`w-full text-left p-3 rounded-xl font-black text-xs uppercase cursor-pointer transition-colors ${activeMetalSubmodal === "Perfil PTR" ? "bg-red-600 text-white" : "bg-white text-slate-600 hover:bg-red-50 border border-slate-200"}`}
            >
              Perfil PTR
            </button>
            <button
              onClick={() => setActiveMetalSubmodal("Placas Metálicas")}
              className={`w-full text-left p-3 rounded-xl font-black text-xs uppercase cursor-pointer transition-colors ${activeMetalSubmodal === "Placas Metálicas" ? "bg-red-600 text-white" : "bg-white text-slate-600 hover:bg-red-50 border border-slate-200"}`}
            >
              Placas Metálicas
            </button>
            <button
              onClick={() => setActiveMetalSubmodal("Lámina Galvanizada")}
              className={`w-full text-left p-3 rounded-xl font-black text-xs uppercase cursor-pointer transition-colors ${activeMetalSubmodal === "Lámina Galvanizada" ? "bg-red-600 text-white" : "bg-white text-slate-600 hover:bg-red-50 border border-slate-200"}`}
            >
              Lámina Galvanizada
            </button>
            <button
              onClick={() => setActiveMetalSubmodal("Canales Tipo C")}
              className={`w-full text-left p-3 rounded-xl font-black text-xs uppercase cursor-pointer transition-colors ${activeMetalSubmodal === "Canales Tipo C" ? "bg-red-600 text-white" : "bg-white text-slate-600 hover:bg-red-50 border border-slate-200"}`}
            >
              Canales Tipo C
            </button>
            <button
              onClick={() => setActiveMetalSubmodal("Perfiles OC")}
              className={`w-full text-left p-3 rounded-xl font-black text-xs uppercase cursor-pointer transition-colors ${activeMetalSubmodal === "Perfiles OC" ? "bg-red-600 text-white" : "bg-white text-slate-600 hover:bg-red-50 border border-slate-200"}`}
            >
              Perfiles OC
            </button>
            <button
              onClick={() => setActiveMetalSubmodal("Viga IR")}
              className={`w-full text-left p-3 rounded-xl font-black text-xs uppercase cursor-pointer transition-colors ${activeMetalSubmodal === "Viga IR" ? "bg-red-600 text-white" : "bg-white text-slate-600 hover:bg-red-50 border border-slate-200"}`}
            >
              Viga IR
            </button>
            <button
              onClick={() => setActiveMetalSubmodal("Reporte Global")}
              className={`w-full text-left p-3 rounded-xl font-black text-xs uppercase cursor-pointer transition-colors mt-2 border-t-2 ${activeMetalSubmodal === "Reporte Global" ? "bg-amber-600 text-white border-amber-600" : "bg-white text-slate-800 hover:bg-amber-50 border-slate-200"}`}
            >
              Reporte Global
            </button>
          </div>

          <div className="flex-1 flex flex-col relative overflow-hidden bg-white rounded-br-3xl">
            {activeMetalSubmodal === "Reporte Global" ? (
              renderGlobalMetalReport()
            ) : (
              <>
                <div className="overflow-x-auto flex-1 p-4">
                  <table className="text-xs text-left border-collapse w-full">
                    <thead className="bg-slate-100 text-slate-700 uppercase tracking-wider select-none sticky top-0 z-10 font-black">
                      <tr>
                        <th className="p-2 border border-slate-200 text-center w-[40px]">
                          No.
                        </th>
                        <th className="p-2 border border-slate-200 text-center">
                          Eje
                        </th>
                        <th className="p-2 border border-slate-200 text-center">
                          Clave
                        </th>
                        <th className="p-2 border border-slate-200 text-center">
                          Pzas
                        </th>
                        <th className="p-2 border border-slate-200 text-center">
                          Largo (m)
                        </th>
                        {activeMetalSubmodal === "Placas Metálicas" ||
                        activeMetalSubmodal === "Lámina Galvanizada" ? (
                          <>
                            <th className="p-2 border border-slate-200 text-center">
                              Ancho (m)
                            </th>
                            <th className="p-2 border border-slate-200 text-center">
                              Área (m²)
                            </th>
                            <th className="p-2 border border-slate-200 text-center">
                              {activeMetalSubmodal === "Lámina Galvanizada"
                                ? "Perfil"
                                : "Espesor"}
                            </th>
                            {activeMetalSubmodal === "Lámina Galvanizada" && (
                              <th className="p-2 border border-slate-200 text-center">
                                Calibre
                              </th>
                            )}
                            <th className="p-2 border border-slate-200 text-center">
                              Peso/m²
                            </th>
                          </>
                        ) : (
                          <>
                            <th className="p-2 border border-slate-200 text-center">
                              {activeMetalSubmodal === "Canales Tipo C" ||
                              activeMetalSubmodal === "Viga IR"
                                ? "Peralte"
                                : activeMetalSubmodal === "Perfiles OC"
                                  ? "Designación (tamaños)"
                                  : "Medida Nominal"}
                            </th>
                            <th className="p-2 border border-slate-200 text-center">
                              Dimensiones
                            </th>
                            <th className="p-2 border border-slate-200 text-center">
                              {activeMetalSubmodal === "Canales Tipo C" ||
                              activeMetalSubmodal === "Viga IR"
                                ? "Peso (kg/m)"
                                : activeMetalSubmodal === "Perfiles OC"
                                  ? "Espesor (in)"
                                  : "Calibre"}
                            </th>
                            <th className="p-2 border border-slate-200 text-center">
                              Peso/ml
                            </th>
                          </>
                        )}
                        <th className="p-2 border border-slate-200 text-center bg-blue-50 text-blue-800">
                          Peso/pza
                        </th>
                        <th className="p-2 border border-slate-200 text-center bg-blue-100 text-blue-900">
                          Peso Total (kg)
                        </th>
                        <th className="p-2 border border-slate-200 text-center"></th>
                      </tr>
                    </thead>
                    <tbody>
                      {metalRows.map((row, idx) => {
                        let pesoMl = 0;
                        let dimString = "-";
                        let optionsCalibre = [];
                        let optionsMedida = [];

                        if (row.tipo === "Polín Monten") {
                          const pd = POLIN_DATA[row.medidaNominal] || {
                            dim: "-",
                            pesos: {},
                          };
                          pesoMl = pd.pesos[row.calibre] || 0;
                          dimString = pd.dim;
                          optionsCalibre = Object.keys(pd.pesos).filter(
                            (k) => pd.pesos[k] !== null,
                          );
                          optionsMedida = Object.keys(POLIN_DATA);
                        } else if (row.tipo === "Viga IPR") {
                          const pdList = IPR_DATA[row.medidaNominal] || [];
                          const item = pdList.find(
                            (x) => x.kg.toString() === row.calibre,
                          ) ||
                            pdList[0] || { kg: 0, h: 0 };
                          pesoMl = item.kg;
                          dimString = `H: ${item.h} mm`;
                          optionsCalibre = pdList.map((x) => x.kg.toString());
                          optionsMedida = Object.keys(IPR_DATA);
                        } else if (row.tipo === "Perfil HSS Cuadrado") {
                          const pd = HSS_DATA[row.medidaNominal] || {
                            dim: "-",
                            pesos: {},
                          };
                          pesoMl = pd.pesos[row.calibre] || 0;
                          dimString = pd.dim;
                          optionsCalibre = Object.keys(pd.pesos).filter(
                            (k) => pd.pesos[k] !== null,
                          );
                          optionsMedida = Object.keys(HSS_DATA);
                        } else if (row.tipo === "Perfil HSS Rectangular") {
                          const pd = HSS_RECT_DATA[row.medidaNominal] || {
                            dim: "-",
                            pesos: {},
                          };
                          pesoMl = pd.pesos[row.calibre] || 0;
                          dimString = pd.dim;
                          optionsCalibre = Object.keys(pd.pesos).filter(
                            (k) => pd.pesos[k] !== null,
                          );
                          optionsMedida = Object.keys(HSS_RECT_DATA);
                        } else if (row.tipo === "Perfil PTR") {
                          const pd = PTR_DATA[row.medidaNominal] || {
                            dim: "-",
                            pesos: {},
                          };
                          pesoMl = pd.pesos[row.calibre] || 0;
                          dimString = pd.dim;
                          optionsCalibre = Object.keys(pd.pesos).filter(
                            (k) => pd.pesos[k] !== null,
                          );
                          optionsMedida = Object.keys(PTR_DATA);
                        } else if (row.tipo === "Placas Metálicas") {
                          pesoMl = PLACAS_DATA[row.calibre] || 0; // pesoMl here means pesoM2
                          optionsCalibre = Object.keys(PLACAS_DATA);
                        } else if (row.tipo === "Lámina Galvanizada") {
                          const pd = LAMINA_DATA[row.medidaNominal] || {
                            pesos: {},
                          };
                          pesoMl = pd.pesos[row.calibre] || 0;
                          optionsCalibre = Object.keys(pd.pesos).filter(
                            (k) => pd.pesos[k] !== null,
                          );
                          optionsMedida = Object.keys(LAMINA_DATA);
                        } else if (row.tipo === "Canales Tipo C") {
                          const pd = CANALES_C_DATA[row.medidaNominal] || {
                            dim: "-",
                            pesos: {},
                          };
                          pesoMl = pd.pesos[row.calibre] || 0;
                          dimString = pd.dim;
                          optionsCalibre = Object.keys(pd.pesos).filter(
                            (k) => pd.pesos[k] !== null,
                          );
                          optionsMedida = Object.keys(CANALES_C_DATA);
                        } else if (row.tipo === "Perfiles OC") {
                          const pd = PERFILES_OC_DATA[row.medidaNominal] || {
                            dim: "-",
                            pesos: {},
                          };
                          pesoMl = pd.pesos[row.calibre] || 0;
                          dimString = pd.dim;
                          optionsCalibre = Object.keys(pd.pesos).filter(
                            (k) => pd.pesos[k] !== null,
                          );
                          optionsMedida = Object.keys(PERFILES_OC_DATA);
                        } else if (row.tipo === "Viga IR") {
                          const pdList = VIGA_IR_DATA[row.medidaNominal] || [];
                          pesoMl = parseFloat(row.calibre) || 0;
                          dimString = "-";
                          optionsCalibre = pdList.map((x) => x.toString());
                          optionsMedida = Object.keys(VIGA_IR_DATA);
                        }
                        const pzas = isNaN(parseFloat(row.pzas))
                          ? 1
                          : parseFloat(row.pzas);

                        let pesoPza = 0;
                        if (
                          row.tipo === "Placas Metálicas" ||
                          row.tipo === "Lámina Galvanizada"
                        ) {
                          const w = parseFloat(row.ancho) || 0;
                          const l = parseFloat(row.largo) || 0;
                          pesoPza = w * l * pesoMl;
                        } else {
                          pesoPza = (parseFloat(row.largo) || 0) * pesoMl;
                        }
                        const totalKg = pesoPza * pzas;

                        return (
                          <tr
                            key={row.id}
                            className="hover:bg-slate-50 transition-colors group"
                          >
                            <td className="p-2 border border-slate-200 text-center text-[10px] text-slate-400 font-bold">
                              {idx + 1}
                            </td>
                            <td className="p-0 border border-slate-200">
                              <input
                                type="text"
                                className="w-full text-center p-2 bg-transparent outline-none uppercase font-bold"
                                value={row.eje || ""}
                                onChange={(e) =>
                                  updateMetalRow(idx, "eje", e.target.value)
                                }
                              />
                            </td>
                            <td className="p-0 border border-slate-200">
                              <input
                                type="text"
                                className="w-full text-center p-2 bg-transparent outline-none uppercase font-bold"
                                value={row.clave || ""}
                                onChange={(e) =>
                                  updateMetalRow(idx, "clave", e.target.value)
                                }
                              />
                            </td>
                            <td className="p-0 border border-slate-200">
                              <input
                                type="number"
                                className="w-full text-center p-2 bg-transparent outline-none font-bold text-slate-800"
                                value={row.pzas !== undefined ? row.pzas : "1"}
                                onChange={(e) =>
                                  updateMetalRow(idx, "pzas", e.target.value)
                                }
                              />
                            </td>
                            <td className="p-0 border border-slate-200">
                              <input
                                type="number"
                                className="w-full text-center p-2 bg-transparent outline-none font-bold"
                                value={row.largo || ""}
                                onChange={(e) =>
                                  updateMetalRow(idx, "largo", e.target.value)
                                }
                              />
                            </td>

                            {row.tipo === "Placas Metálicas" ||
                            row.tipo === "Lámina Galvanizada" ? (
                              <>
                                <td className="p-0 border border-slate-200">
                                  <input
                                    type="number"
                                    className="w-full text-center p-2 bg-transparent outline-none font-bold"
                                    value={
                                      row.ancho !== undefined ? row.ancho : ""
                                    }
                                    onChange={(e) =>
                                      updateMetalRow(
                                        idx,
                                        "ancho",
                                        e.target.value,
                                      )
                                    }
                                  />
                                </td>
                                <td className="p-2 border border-slate-200 text-center font-bold text-slate-500 whitespace-nowrap">
                                  {(
                                    (parseFloat(row.ancho) || 0) *
                                    (parseFloat(row.largo) || 0)
                                  ).toFixed(2)}
                                </td>
                                {row.tipo === "Lámina Galvanizada" && (
                                  <td className="p-1 border border-slate-200 text-center">
                                    <select
                                      className="w-full text-center p-1 bg-transparent outline-none font-bold cursor-pointer"
                                      value={row.medidaNominal}
                                      onChange={(e) => {
                                        const newMedida = e.target.value;
                                        let newCalibre = row.calibre;
                                        const newPd = LAMINA_DATA[newMedida];
                                        if (
                                          newPd &&
                                          newPd.pesos[newCalibre] === undefined
                                        ) {
                                          newCalibre =
                                            Object.keys(newPd.pesos).find(
                                              (k) =>
                                                newPd.pesos[k] !== undefined &&
                                                newPd.pesos[k] !== null,
                                            ) || "26";
                                        }
                                        updateMetalRow(idx, {
                                          medidaNominal: newMedida,
                                          calibre: newCalibre,
                                        });
                                      }}
                                    >
                                      {optionsMedida.map((m) => (
                                        <option key={m} value={m}>
                                          {m}
                                        </option>
                                      ))}
                                    </select>
                                  </td>
                                )}
                                <td className="p-1 border border-slate-200 text-center">
                                  <select
                                    className="w-full text-center p-1 bg-transparent outline-none font-bold cursor-pointer"
                                    value={row.calibre}
                                    onChange={(e) =>
                                      updateMetalRow(
                                        idx,
                                        "calibre",
                                        e.target.value,
                                      )
                                    }
                                  >
                                    {optionsCalibre.map((c) => (
                                      <option key={c} value={c}>
                                        {c}
                                      </option>
                                    ))}
                                  </select>
                                </td>
                              </>
                            ) : (
                              <>
                                <td className="p-1 border border-slate-200 text-center">
                                  <select
                                    className="w-full text-center p-1 bg-transparent outline-none font-bold cursor-pointer"
                                    value={row.medidaNominal}
                                    onChange={(e) => {
                                      const newMedida = e.target.value;
                                      let newCalibre = row.calibre;
                                      if (row.tipo === "Polín Monten") {
                                        const newPd = POLIN_DATA[newMedida];
                                        if (
                                          newPd &&
                                          newPd.pesos[newCalibre] === undefined
                                        ) {
                                          newCalibre =
                                            Object.keys(newPd.pesos).find(
                                              (k) =>
                                                newPd.pesos[k] !== undefined &&
                                                newPd.pesos[k] !== null,
                                            ) || "14";
                                        }
                                      } else if (row.tipo === "Viga IPR") {
                                        const newPdList =
                                          IPR_DATA[newMedida] || [];
                                        if (
                                          !newPdList.find(
                                            (x) =>
                                              x.kg.toString() === newCalibre,
                                          )
                                        ) {
                                          newCalibre = newPdList[0]
                                            ? newPdList[0].kg.toString()
                                            : "0";
                                        }
                                      } else if (
                                        row.tipo === "Perfil HSS Cuadrado"
                                      ) {
                                        const newPd = HSS_DATA[newMedida];
                                        if (
                                          newPd &&
                                          newPd.pesos[newCalibre] === undefined
                                        ) {
                                          newCalibre =
                                            Object.keys(newPd.pesos).find(
                                              (k) =>
                                                newPd.pesos[k] !== undefined &&
                                                newPd.pesos[k] !== null,
                                            ) || '1/8"';
                                        }
                                      } else if (
                                        row.tipo === "Perfil HSS Rectangular"
                                      ) {
                                        const newPd = HSS_RECT_DATA[newMedida];
                                        if (
                                          newPd &&
                                          newPd.pesos[newCalibre] === undefined
                                        ) {
                                          newCalibre =
                                            Object.keys(newPd.pesos).find(
                                              (k) =>
                                                newPd.pesos[k] !== undefined &&
                                                newPd.pesos[k] !== null,
                                            ) || '3/16"';
                                        }
                                      } else if (row.tipo === "Perfil PTR") {
                                        const newPd = PTR_DATA[newMedida];
                                        if (
                                          newPd &&
                                          newPd.pesos[newCalibre] === undefined
                                        ) {
                                          newCalibre =
                                            Object.keys(newPd.pesos).find(
                                              (k) =>
                                                newPd.pesos[k] !== undefined &&
                                                newPd.pesos[k] !== null,
                                            ) || "14";
                                        }
                                      } else if (
                                        row.tipo === "Canales Tipo C"
                                      ) {
                                        const newPd = CANALES_C_DATA[newMedida];
                                        if (
                                          newPd &&
                                          newPd.pesos[newCalibre] === undefined
                                        ) {
                                          newCalibre =
                                            Object.keys(newPd.pesos).find(
                                              (k) =>
                                                newPd.pesos[k] !== undefined &&
                                                newPd.pesos[k] !== null,
                                            ) || "8.04";
                                        }
                                      } else if (row.tipo === "Perfiles OC") {
                                        const newPd =
                                          PERFILES_OC_DATA[newMedida];
                                        if (
                                          newPd &&
                                          newPd.pesos[newCalibre] === undefined
                                        ) {
                                          newCalibre =
                                            Object.keys(newPd.pesos).find(
                                              (k) =>
                                                newPd.pesos[k] !== undefined &&
                                                newPd.pesos[k] !== null,
                                            ) || '0.500"';
                                        }
                                      } else if (row.tipo === "Viga IR") {
                                        const newPdList =
                                          VIGA_IR_DATA[newMedida] || [];
                                        if (
                                          !newPdList.includes(
                                            parseFloat(newCalibre),
                                          )
                                        ) {
                                          newCalibre = newPdList[0]
                                            ? newPdList[0].toString()
                                            : "0";
                                        }
                                      }
                                      updateMetalRow(idx, {
                                        medidaNominal: newMedida,
                                        calibre: newCalibre,
                                      });
                                    }}
                                  >
                                    {optionsMedida.map((m) => (
                                      <option key={m} value={m}>
                                        {m}
                                      </option>
                                    ))}
                                  </select>
                                </td>
                                <td className="p-2 border border-slate-200 text-center font-bold text-slate-500 whitespace-nowrap">
                                  {dimString}
                                </td>
                                <td className="p-1 border border-slate-200 text-center">
                                  <select
                                    className="w-full text-center p-1 bg-transparent outline-none font-bold cursor-pointer"
                                    value={row.calibre}
                                    onChange={(e) =>
                                      updateMetalRow(
                                        idx,
                                        "calibre",
                                        e.target.value,
                                      )
                                    }
                                  >
                                    {optionsCalibre.map((c) => (
                                      <option key={c} value={c}>
                                        {row.tipo === "Viga IPR" ||
                                        row.tipo === "Canales Tipo C" ||
                                        row.tipo === "Viga IR"
                                          ? `${parseFloat(c).toFixed(1)} kg/m`
                                          : c}
                                      </option>
                                    ))}
                                  </select>
                                </td>
                              </>
                            )}
                            <td className="p-2 border border-slate-200 text-center font-black text-blue-600">
                              {pesoMl > 0 ? pesoMl.toFixed(3) : "-"}
                            </td>
                            <td className="p-2 border border-slate-200 text-center font-black text-blue-800">
                              {pesoPza > 0 ? pesoPza.toFixed(2) : "-"}
                            </td>
                            <td className="p-2 border border-slate-200 text-center font-black bg-blue-100 text-blue-900">
                              {totalKg > 0 ? totalKg.toFixed(2) : "-"}
                            </td>
                            <td className="p-1 border border-slate-200 text-center">
                              <button
                                onClick={() => deleteMetalRow(row.id)}
                                className="text-slate-300 hover:text-red-500 transition-colors p-1"
                              >
                                <Trash2 size={14} />
                              </button>
                            </td>
                          </tr>
                        );
                      })}
                      <tr>
                        <td
                          colSpan="12"
                          className="p-2 border border-slate-200"
                        >
                          <button
                            onClick={addMetalRow}
                            className="w-full py-2 bg-slate-50 border border-slate-200 border-dashed rounded-lg text-slate-500 hover:bg-slate-100 hover:text-slate-700 font-bold uppercase text-[10px] tracking-widest flex items-center justify-center gap-2 transition-colors"
                          >
                            <Plus size={14} /> Añadir Fila
                          </button>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                <div className="bg-slate-50 p-4 border-t border-slate-200 flex flex-col gap-2 shrink-0 rounded-br-3xl">
                  {totalWeightPolin > 0 && (
                    <div className="flex justify-end w-full">
                      <div className="bg-white rounded-2xl p-4 px-6 border border-slate-200 shadow-sm flex items-center gap-6 justify-between min-w-[300px]">
                        <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest">
                          Total Polín Monten
                        </span>
                        <div className="flex items-baseline gap-2">
                          <span className="text-2xl font-black text-slate-800 leading-none">
                            {totalWeightPolin.toFixed(2)}
                          </span>
                          <span className="text-sm font-black text-slate-400 uppercase">
                            kg
                          </span>
                        </div>
                      </div>
                    </div>
                  )}
                  {totalWeightIPR > 0 && (
                    <div className="flex justify-end w-full">
                      <div className="bg-white rounded-2xl p-4 px-6 border border-slate-200 shadow-sm flex items-center gap-6 justify-between min-w-[300px]">
                        <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest">
                          Total Viga IPR
                        </span>
                        <div className="flex items-baseline gap-2">
                          <span className="text-2xl font-black text-slate-800 leading-none">
                            {totalWeightIPR.toFixed(2)}
                          </span>
                          <span className="text-sm font-black text-slate-400 uppercase">
                            kg
                          </span>
                        </div>
                      </div>
                    </div>
                  )}
                  {totalWeightHSS > 0 && (
                    <div className="flex justify-end w-full">
                      <div className="bg-white rounded-2xl p-4 px-6 border border-slate-200 shadow-sm flex items-center gap-6 justify-between min-w-[300px]">
                        <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest">
                          Total Perfil HSS Cuad.
                        </span>
                        <div className="flex items-baseline gap-2">
                          <span className="text-2xl font-black text-slate-800 leading-none">
                            {totalWeightHSS.toFixed(2)}
                          </span>
                          <span className="text-sm font-black text-slate-400 uppercase">
                            kg
                          </span>
                        </div>
                      </div>
                    </div>
                  )}
                  {totalWeightHSSRect > 0 && (
                    <div className="flex justify-end w-full">
                      <div className="bg-white rounded-2xl p-4 px-6 border border-slate-200 shadow-sm flex items-center gap-6 justify-between min-w-[300px]">
                        <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest">
                          Total Perfil HSS Rect.
                        </span>
                        <div className="flex items-baseline gap-2">
                          <span className="text-2xl font-black text-slate-800 leading-none">
                            {totalWeightHSSRect.toFixed(2)}
                          </span>
                          <span className="text-sm font-black text-slate-400 uppercase">
                            kg
                          </span>
                        </div>
                      </div>
                    </div>
                  )}
                  {totalWeightPlacas > 0 && (
                    <div className="flex justify-end w-full">
                      <div className="bg-white rounded-2xl p-4 px-6 border border-slate-200 shadow-sm flex items-center gap-6 justify-between min-w-[300px]">
                        <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest">
                          Total Placas Metálicas
                        </span>
                        <div className="flex items-baseline gap-2">
                          <span className="text-2xl font-black text-slate-800 leading-none">
                            {totalWeightPlacas.toFixed(2)}
                          </span>
                          <span className="text-sm font-black text-slate-400 uppercase">
                            kg
                          </span>
                        </div>
                      </div>
                    </div>
                  )}
                  {totalWeightLamina > 0 && (
                    <div className="flex justify-end w-full">
                      <div className="bg-white rounded-2xl p-4 px-6 border border-slate-200 shadow-sm flex items-center gap-6 justify-between min-w-[300px]">
                        <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest">
                          Total Lámina Galvanizada
                        </span>
                        <div className="flex items-baseline gap-2">
                          <span className="text-2xl font-black text-slate-800 leading-none">
                            {totalWeightLamina.toFixed(2)}
                          </span>
                          <span className="text-sm font-black text-slate-400 uppercase">
                            kg
                          </span>
                        </div>
                      </div>
                    </div>
                  )}
                  {totalWeightPTR > 0 && (
                    <div className="flex justify-end w-full">
                      <div className="bg-white rounded-2xl p-4 px-6 border border-slate-200 shadow-sm flex items-center gap-6 justify-between min-w-[300px]">
                        <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest">
                          Total Perfil PTR
                        </span>
                        <div className="flex items-baseline gap-2">
                          <span className="text-2xl font-black text-slate-800 leading-none">
                            {totalWeightPTR.toFixed(2)}
                          </span>
                          <span className="text-sm font-black text-slate-400 uppercase">
                            kg
                          </span>
                        </div>
                      </div>
                    </div>
                  )}
                  {totalWeightCanalesC > 0 && (
                    <div className="flex justify-end w-full">
                      <div className="bg-white rounded-2xl p-4 px-6 border border-slate-200 shadow-sm flex items-center gap-6 justify-between min-w-[300px]">
                        <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest">
                          Total Canales Tipo C
                        </span>
                        <div className="flex items-baseline gap-2">
                          <span className="text-2xl font-black text-slate-800 leading-none">
                            {totalWeightCanalesC.toFixed(2)}
                          </span>
                          <span className="text-sm font-black text-slate-400 uppercase">
                            kg
                          </span>
                        </div>
                      </div>
                    </div>
                  )}
                  {totalWeightPerfilesOC > 0 && (
                    <div className="flex justify-end w-full">
                      <div className="bg-white rounded-2xl p-4 px-6 border border-slate-200 shadow-sm flex items-center gap-6 justify-between min-w-[300px]">
                        <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest">
                          Total Perfiles OC
                        </span>
                        <div className="flex items-baseline gap-2">
                          <span className="text-2xl font-black text-slate-800 leading-none">
                            {totalWeightPerfilesOC.toFixed(2)}
                          </span>
                          <span className="text-sm font-black text-slate-400 uppercase">
                            kg
                          </span>
                        </div>
                      </div>
                    </div>
                  )}
                  {totalWeightVigaIR > 0 && (
                    <div className="flex justify-end w-full">
                      <div className="bg-white rounded-2xl p-4 px-6 border border-slate-200 shadow-sm flex items-center gap-6 justify-between min-w-[300px]">
                        <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest">
                          Total Viga IR
                        </span>
                        <div className="flex items-baseline gap-2">
                          <span className="text-2xl font-black text-slate-800 leading-none">
                            {totalWeightVigaIR.toFixed(2)}
                          </span>
                          <span className="text-sm font-black text-slate-400 uppercase">
                            kg
                          </span>
                        </div>
                      </div>
                    </div>
                  )}
                  <div className="flex justify-end w-full">
                    <div className="bg-blue-600 rounded-2xl p-4 px-6 border border-blue-700 shadow-lg flex items-center gap-6 justify-between min-w-[300px]">
                      <span className="text-[10px] font-black text-blue-200 uppercase tracking-widest">
                        Peso Total
                      </span>
                      <div className="flex items-baseline gap-2">
                        <span className="text-2xl font-black text-white leading-none">
                          {totalWeight.toFixed(2)}
                        </span>
                        <span className="text-sm font-black text-blue-300 uppercase">
                          kg
                        </span>
                      </div>
                    </div>
                  </div>
                </div>
              </>
            )}
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-[#fcfdfe] font-sans text-slate-900 p-4 md:p-6 flex flex-col">
      {renderSteelSubmodal()}
      {renderWallSubmodal()}
      {renderCasetonSubmodal()}
      {renderConceptSelectorModal()}
      {renderGeneratorModal()}
      {renderLevelReport()}
      {renderGlobalSteelReport()}
      <header className="max-w-7xl mx-auto mb-6 w-full shrink-0">
        <div className="flex justify-between items-start">
          {activePartidaId ? (
            <button
              onClick={() => {
                setActivePartidaId(null);
                setShowSettings(false);
                setShowWallGenerator(false);
                setShowStructureGenerator(false);
                setShowMetalStructureGenerator(false);
                setShowAnalysisNervadura(false);
                setActiveWallSubmodal(null);
                setActiveSteelSubmodal(null);
                setActiveCasetonSubmodal(null);
                setActiveLevelReport(null);
              }}
              className="flex items-center gap-2 text-slate-400 hover:text-blue-600 font-bold text-[10px] uppercase tracking-[0.2em] transition-colors mb-4"
            >
              <ChevronLeft size={16} /> Volver a Opciones de Obra
            </button>
          ) : (
            <button
              onClick={onBack}
              className="flex items-center gap-2 text-slate-400 hover:text-blue-600 font-bold text-[10px] uppercase tracking-[0.2em] transition-colors mb-4"
            >
              <ChevronLeft size={16} /> Volver al Gestor de Obras
            </button>
          )}
        </div>
        <div className="flex justify-between items-end flex-wrap gap-4">
          <div className="flex flex-col gap-4">
            <div className="flex items-center gap-4">
              <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center text-white font-black text-sm">
                X
              </div>
              <div className="h-6 w-px bg-slate-200"></div>
              <span className="text-blue-600 font-black uppercase tracking-[0.4em] text-[8px]">
                Asset Management
              </span>
              <div className="flex items-center gap-2 px-3 py-1 bg-slate-50 border border-slate-100 rounded-lg">
                <div
                  className={`w-1.5 h-1.5 rounded-full ${isSavingAny ? "bg-amber-400 animate-pulse shadow-[0_0_8px_rgba(251,191,36,0.5)]" : "bg-teal-500 shadow-[0_0_8px_rgba(20,184,166,0.3)]"}`}
                ></div>
                <span
                  className={`text-[8px] font-black uppercase tracking-widest ${isSavingAny ? "text-amber-600" : "text-teal-600"}`}
                >
                  {isSavingAny ? "Sincronizando..." : "Guardado"}
                </span>
              </div>
            </div>
            <div>
              <h1 className="text-4xl font-black tracking-tighter mb-2">
                {obraInfo.nombre || ""}
              </h1>
              {activePartidaId ? (
                <p className="text-blue-600 font-black flex items-center gap-3 uppercase text-xs tracking-widest">
                  <Layers size={14} /> PARTIDA:{" "}
                  <span className="text-slate-600">{activePartida.nombre}</span>
                </p>
              ) : (
                <p className="text-slate-400 font-black flex items-center gap-3 uppercase text-[10px] tracking-widest">
                  <LayoutDashboard size={14} /> MENÚ DE PARTIDAS
                </p>
              )}
            </div>
          </div>
          <div className="flex gap-3 flex-wrap">
            {activePartidaId && (
              <>
                <button
                  onClick={() => {
                    setShowAnalysisNervadura(!showAnalysisNervadura);
                    setShowMetalStructureGenerator(false);
                    setShowWallGenerator(false);
                    setShowStructureGenerator(false);
                    setShowSettings(false);
                  }}
                  className={`px-4 py-2.5 rounded-xl font-black uppercase text-[10px] shadow-lg transition-all flex items-center gap-2 ${showAnalysisNervadura ? "bg-cyan-700 text-white" : "bg-white border text-slate-700 hover:bg-cyan-50"}`}
                >
                  <Calculator size={16} /> Analisis/m2 Nerv.
                </button>
                <button
                  onClick={() => {
                    setShowMetalStructureGenerator(
                      !showMetalStructureGenerator,
                    );
                    setShowAnalysisNervadura(false);
                    setShowWallGenerator(false);
                    setShowStructureGenerator(false);
                    setShowSettings(false);
                  }}
                  className={`px-4 py-2.5 rounded-xl font-black uppercase text-[10px] shadow-lg transition-all flex items-center gap-2 ${showMetalStructureGenerator ? "bg-red-700 text-white" : "bg-white border text-slate-700 hover:bg-red-50"}`}
                >
                  <Wrench size={16} /> Est. Metálica
                </button>
                <button
                  onClick={() => {
                    setShowStructureGenerator(!showStructureGenerator);
                    setShowAnalysisNervadura(false);
                    setShowWallGenerator(false);
                    setShowMetalStructureGenerator(false);
                    setShowSettings(false);
                  }}
                  className={`px-4 py-2.5 rounded-xl font-black uppercase text-[10px] shadow-lg transition-all flex items-center gap-2 ${showStructureGenerator ? "bg-amber-600 text-white" : "bg-white border text-slate-700 hover:bg-amber-50"}`}
                >
                  <Box size={16} /> Estructura
                </button>
                <button
                  onClick={() => {
                    setShowWallGenerator(!showWallGenerator);
                    setShowAnalysisNervadura(false);
                    setShowStructureGenerator(false);
                    setShowMetalStructureGenerator(false);
                    setShowSettings(false);
                  }}
                  className={`px-4 py-2.5 rounded-xl font-black uppercase text-[10px] shadow-lg transition-all flex items-center gap-2 ${showWallGenerator ? "bg-indigo-700 text-white" : "bg-white border text-slate-700 hover:bg-indigo-50"}`}
                >
                  <Building2 size={16} /> Muros
                </button>
              </>
            )}
            <button
              onClick={() => {
                setShowSettings(!showSettings);
                setShowAnalysisNervadura(false);
                setShowWallGenerator(false);
                setShowStructureGenerator(false);
                setShowMetalStructureGenerator(false);
              }}
              className={`px-4 py-2.5 rounded-xl font-black uppercase text-[10px] shadow-lg flex items-center gap-2 ${showSettings && !showWallGenerator && !showStructureGenerator && !showMetalStructureGenerator && !showAnalysisNervadura ? "bg-blue-600 text-white" : "bg-white border text-slate-700 hover:bg-slate-50"}`}
            >
              <Settings size={16} /> Ajustes
            </button>
          </div>
        </div>
      </header>
      <main className="max-w-7xl mx-auto w-full flex-1 relative">
        {showAnalysisNervadura ? (
          renderAnalysisNervadura()
        ) : showMetalStructureGenerator ? (
          renderMetalStructureGenerator()
        ) : showStructureGenerator ? (
          <div className="bg-white rounded-3xl shadow-xl border border-gray-100 animate-in fade-in zoom-in-95 duration-300 flex flex-col">
            <div className="bg-amber-800 text-white p-4 flex justify-between items-center shrink-0">
              <div className="flex items-center gap-3">
                <div className="p-2 bg-amber-700 rounded-lg">
                  <Box size={20} className="text-amber-100" />
                </div>
                <div>
                  <h2 className="text-lg font-black uppercase tracking-wider leading-tight">
                    Generador de Estructuras y Acero
                  </h2>
                  <p className="text-[9px] text-amber-200 font-bold uppercase tracking-widest">
                    {activePartida.nombre}
                  </p>
                </div>
              </div>
              <div className="flex items-center gap-3">
                <select
                  value={activeLevelId}
                  onChange={(e) => {
                    setActiveLevelId(e.target.value);
                    setSelectedStructureRows([]);
                  }}
                  className="bg-amber-900/50 text-white font-black text-[10px] uppercase p-2 px-4 rounded-lg border border-amber-600 outline-none cursor-pointer"
                >
                  {niveles.map((n) => (
                    <option key={n.id} value={n.id} className="text-slate-800">
                      {n.nombre}
                    </option>
                  ))}
                </select>
                <button
                  onClick={() => setShowStructureGenerator(false)}
                  className="p-1.5 bg-amber-700 hover:bg-amber-600 rounded-full transition-colors"
                >
                  <X size={20} />
                </button>
              </div>
            </div>
            <div
              className={`overflow-x-auto bg-white ${estructuras.length > 12 ? "max-h-[450px] overflow-y-auto overscroll-contain" : ""}`}
            >
              <table className="text-left border-collapse text-xs table-fixed w-full min-w-[900px]">
                <thead className="bg-slate-100 text-slate-500 font-black uppercase tracking-wider border-b-2 border-slate-200 text-[9px] sticky top-0 z-30">
                  <tr>
                    <th
                      className="p-2 text-center border-r border-slate-200 relative select-none bg-slate-100"
                      rowSpan={2}
                      style={{ width: structColWidths.select }}
                    >
                      <button
                        onClick={toggleAllStructureRows}
                        className="hover:text-amber-600 transition-colors"
                      >
                        {estructuras.length > 0 &&
                        selectedStructureRows.length === estructuras.length ? (
                          <CheckSquare size={14} />
                        ) : (
                          <Square size={14} />
                        )}
                      </button>
                    </th>
                    <th
                      className="p-2 text-center border-r relative select-none bg-slate-100"
                      rowSpan={2}
                      style={{ width: structColWidths.no }}
                    >
                      No.
                    </th>
                    <th
                      className="p-2 border-r relative select-none bg-slate-100"
                      rowSpan={2}
                      style={{ width: structColWidths.eje }}
                    >
                      Eje
                    </th>
                    <th
                      className="p-2 border-r relative select-none bg-slate-100"
                      rowSpan={2}
                      style={{ width: structColWidths.clave }}
                    >
                      Clave
                    </th>
                    <th
                      className="p-2 border-r text-center relative select-none leading-tight bg-slate-100"
                      rowSpan={2}
                      style={{ width: structColWidths.tipo }}
                    >
                      Tipo Elemento
                    </th>
                    <th
                      className="p-2 text-center border-r bg-indigo-50 text-indigo-800 relative select-none"
                      rowSpan={2}
                      style={{ width: structColWidths.i || 30 }}
                    >
                      I
                    </th>
                    <th
                      className="p-2 text-center border-r bg-blue-50 text-blue-800 relative select-none"
                      rowSpan={2}
                      style={{ width: structColWidths.largo }}
                    >
                      Largo (m)
                    </th>
                    <th
                      className="p-2 text-center border-r bg-blue-50 text-blue-800 relative select-none"
                      rowSpan={2}
                      style={{ width: structColWidths.ancho }}
                    >
                      Ancho (m)
                    </th>
                    <th
                      className="p-2 text-center border-r bg-blue-50 text-blue-800 relative select-none leading-tight"
                      rowSpan={2}
                      style={{ width: structColWidths.alto }}
                    >
                      Esp/Alto
                    </th>
                    <th
                      className="p-2 text-center border-r bg-blue-50 text-blue-800 relative select-none"
                      rowSpan={2}
                      style={{ width: structColWidths.piezas }}
                    >
                      Pzas
                    </th>
                    <th
                      className="p-1 border-b text-center bg-sky-100 text-sky-900 select-none"
                      colSpan={4}
                    >
                      <div className="flex items-center justify-between px-2">
                        <span className="leading-tight uppercase text-[8px] font-black">
                          Volúmenes / Áreas Totales
                        </span>
                        <button
                          onClick={() => setShowGlobalSteelReport(true)}
                          className="flex items-center gap-1 bg-slate-800 text-white px-2 py-1 rounded text-[8px] uppercase font-black hover:bg-slate-700 shadow-sm transition-transform active:scale-95 cursor-pointer shrink-0"
                        >
                          <ListPlus size={10} /> Reporte Global
                        </button>
                      </div>
                    </th>
                    <th
                      className="p-2 text-center border-l relative select-none bg-slate-100"
                      rowSpan={2}
                      style={{ width: structColWidths.action }}
                    >
                      <div className="flex justify-center items-center gap-1">
                        <button
                          onClick={handleCopySelectedStructures}
                          className={`p-1 rounded transition-colors ${selectedStructureRows.length > 0 ? "text-blue-600 hover:bg-blue-100" : "text-slate-300 cursor-not-allowed"}`}
                          disabled={selectedStructureRows.length === 0}
                          title="Copiar Seleccionados"
                        >
                          <Copy size={14} />
                        </button>
                        <button
                          onClick={() => {
                            if (selectedStructureRows.length > 0) {
                              updateActiveEstructuras((prev) =>
                                prev.filter(
                                  (x) => !selectedStructureRows.includes(x.id),
                                ),
                              );
                              setSelectedStructureRows([]);
                            }
                          }}
                          className={`p-1 rounded transition-colors ${selectedStructureRows.length > 0 ? "text-red-500 hover:bg-red-50" : "text-slate-300 cursor-not-allowed"}`}
                          disabled={selectedStructureRows.length === 0}
                          title="Borrar Seleccionados"
                        >
                          <Trash2 size={14} />
                        </button>
                      </div>
                    </th>
                  </tr>
                  <tr>
                    <th
                      className="p-1 border-r text-center bg-sky-200/50 text-sky-800 relative select-none"
                      style={{ width: structColWidths.concreto }}
                    >
                      Concreto
                      <br />
                      (m3)
                    </th>
                    <th
                      className="p-1 border-r text-center bg-amber-100/50 text-amber-800 relative select-none"
                      style={{ width: structColWidths.cimbra }}
                    >
                      Cimbra(m2/ml)
                    </th>
                    <th
                      className="p-1 border-r text-center bg-orange-100/50 text-orange-800 relative select-none"
                      style={{ width: structColWidths.cimbraFrontera }}
                    >
                      Cim.Front(ml)
                    </th>
                    <th
                      className="p-1 border-r text-center bg-slate-200/50 text-slate-800 relative select-none"
                      style={{ width: structColWidths.acero }}
                    >
                      Acero (kg)
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {estructuras.map((e, i) => {
                    const concreto = calcConcreto(e),
                      cimbra = calcCimbra(e),
                      cimbraFrontera = calcCimbraFrontera(e),
                      aceroKg = calcAceroTotalKg(e.aceros, e.piezas);
                    return (
                      <tr
                        key={e.id}
                        className={`transition-colors ${selectedStructureRows.includes(e.id) ? "bg-amber-50/50" : "hover:bg-slate-50"}`}
                      >
                        <td className="p-1 text-center border-r border-slate-200">
                          <button
                            onClick={() => toggleStructureRow(e.id)}
                            className={`${selectedStructureRows.includes(e.id) ? "text-amber-600" : "text-slate-300 hover:text-amber-500"} transition-colors`}
                          >
                            {selectedStructureRows.includes(e.id) ? (
                              <CheckSquare size={14} />
                            ) : (
                              <Square size={14} />
                            )}
                          </button>
                        </td>
                        <td className="p-1 text-center font-black text-slate-400 border-r">
                          {i + 1}
                        </td>
                        <td className="p-0 border-r">
                          <DebouncedCell
                            value={e.eje}
                            onChange={(v) =>
                              updateEstructuraField(e.id, "eje", v)
                            }
                            className="w-full p-1.5 bg-transparent font-bold uppercase outline-none focus:bg-white text-xs text-center"
                          />
                        </td>
                        <td className="p-0 border-r">
                          <DebouncedCell
                            value={e.clave}
                            onChange={(v) =>
                              updateEstructuraField(e.id, "clave", v)
                            }
                            className="w-full p-1.5 bg-transparent font-bold uppercase outline-none focus:bg-white text-xs text-center"
                          />
                        </td>
                        <td className="p-0 border-r relative">
                          <select
                            className="w-full p-1.5 bg-transparent outline-none font-bold text-slate-700 uppercase cursor-pointer text-[10px]"
                            value={e.tipo || ""}
                            onChange={(ev) =>
                              updateEstructuraField(
                                e.id,
                                "tipo",
                                ev.target.value,
                              )
                            }
                          >
                            <option value="" disabled>
                              Sel...
                            </option>
                            <option value="zapata aislada">
                              Zapata Aislada
                            </option>
                            <option value="zapata corrida">
                              Zapata Corrida
                            </option>
                            <option value="dado">Dado</option>
                            <option value="contratrabe">Contratrabe</option>
                            <option value="muro">Muro</option>
                            <option value="muro curvo">Muro Curvo</option>
                            <option value="losa">Losa</option>
                            <option value="losa nervada">Losa Nervada</option>
                            <option value="losa vigueta h=35">Losa Vigueta h=35</option>
                            <option value="losa vigueta h=25">Losa Vigueta h=25</option>
                            <option value="losa doble vigueta h=25">Losa Doble Vigueta h=25</option>
                            <option value="losa doble vigueta h=35">Losa Doble Vigueta h=35</option>
                            <option value="escalera papelillo">
                              Escalera Papelillo
                            </option>
                            <option value="rampa de escalera">
                              Rampa de Escalera
                            </option>
                            <option value="columna">Columna</option>
                            <option value="columna circular">
                              Columna Circular
                            </option>
                            <option value="pilas">Pilas</option>
                            <option value="trabe">Trabe</option>
                            <option value="nervadura">Nervadura</option>
                          </select>
                          {["trabe", "contratrabe", "columna", "columna circular", "muro", "muro curvo"].includes((e.tipo || "").toLowerCase()) && (
                            <Layout
                              size={10}
                              className="absolute top-1 right-1 text-sky-500 animate-pulse pointer-events-none"
                            />
                          )}
                        </td>
                        <td className="p-0 border-r bg-indigo-50/10 text-center">
                          <input
                            type="checkbox"
                            checked={Boolean(e.i)}
                            onChange={(ev) =>
                              updateEstructuraField(
                                e.id,
                                "i",
                                ev.target.checked,
                              )
                            }
                            disabled={
                              ![
                                "trabe",
                                "contratrabe",
                                "zapata corrida",
                              ].includes((e.tipo || "").toLowerCase())
                            }
                            className="w-3 h-3 cursor-pointer accent-indigo-600"
                          />
                        </td>
                        <td className="p-0 border-r bg-blue-50/20">
                          <DebouncedCell
                            value={e.largo}
                            onChange={(v) =>
                              updateEstructuraField(e.id, "largo", v)
                            }
                            className="w-full p-1.5 bg-transparent text-center text-blue-900 font-bold text-xs"
                          />
                        </td>
                        <td className="p-0 border-r bg-blue-50/20">
                          <DebouncedCell
                            value={e.ancho}
                            onChange={(v) =>
                              updateEstructuraField(e.id, "ancho", v)
                            }
                            className="w-full p-1.5 bg-transparent text-center text-blue-900 font-bold text-xs"
                          />
                        </td>
                        <td className="p-0 border-r bg-blue-50/20">
                          <div className="relative group">
                            <DebouncedCell
                              value={e.alto}
                              onChange={(v) =>
                                updateEstructuraField(e.id, "alto", v)
                              }
                              onDoubleClick={() => handleStructAltoDoubleClick(e.id)}
                              className="w-full p-1.5 bg-transparent text-center text-blue-900 font-bold text-xs"
                              title={e.structEspesorLosa || e.structEspesorZapata ? `Ajuste Losa: ${e.structEspesorLosa}m, Zapata: ${e.structEspesorZapata}m` : "Doble click para descontar espesor de losa o zapata"}
                            />
                            {["trabe", "contratrabe", "columna", "columna circular", "muro", "muro curvo"].includes((e.tipo || "").toLowerCase().trim()) && (
                              <Layout
                                size={12}
                                className="absolute top-1 right-1 text-sky-500 font-black animate-pulse-slow"
                              />
                            )}
                            {parseFloat(e.descuentoLosa || 0) > 0 && (
                              <div className="absolute -top-1 -right-1 bg-red-500 text-white text-[8px] px-1 rounded-full font-black pointer-events-none">
                                -{e.descuentoLosa}
                              </div>
                            )}
                          </div>
                        </td>
                        <td className="p-0 border-r bg-blue-50/20">
                          <DebouncedCell
                            value={e.piezas}
                            onChange={(v) =>
                              updateEstructuraField(e.id, "piezas", v)
                            }
                            className="w-full p-1.5 bg-transparent font-black text-center text-slate-700 text-xs"
                          />
                        </td>
                        <td
                          className={`p-0 border-r ${e.tipo?.toLowerCase() === "losa nervada" ? "bg-sky-50 transition-colors hover:bg-sky-100 cursor-pointer" : e.tipo?.toLowerCase() === "nervadura" ? "bg-slate-50 cursor-not-allowed opacity-60" : "bg-sky-50 transition-colors"}`}
                          onClick={() =>
                            e.tipo?.toLowerCase() === "losa nervada"
                              ? setActiveCasetonSubmodal(e.id)
                              : null
                          }
                        >
                          {e.tipo?.toLowerCase() === "losa nervada" ? (
                            <div
                              className="w-full h-full p-1 text-center font-black text-sky-700 flex justify-center items-center gap-1 group"
                              title="Modificar Casetones"
                            >
                              {concreto > 0 ? (
                                <span className="text-sm">
                                  {concreto.toFixed(2)}
                                </span>
                              ) : (
                                <span className="text-slate-400 font-normal text-[10px]">
                                  0.00
                                </span>
                              )}
                              <Box
                                size={12}
                                className="text-sky-400 group-hover:text-sky-600 transition-colors"
                              />
                            </div>
                          ) : e.tipo?.toLowerCase() === "nervadura" ? (
                            <div className="w-full h-full p-1 text-center font-black flex justify-center items-center text-sm cursor-not-allowed text-slate-500">
                              -
                            </div>
                          ) : (
                            <div className="w-full h-full p-1 text-center font-black text-sky-700 flex justify-center items-center text-sm cursor-default">
                              {concreto > 0 ? concreto.toFixed(2) : "-"}
                            </div>
                          )}
                        </td>
                        <td className="p-1 border-r text-center font-black text-amber-700 bg-amber-50 text-sm">
                          {e.tipo?.toLowerCase() === "nervadura"
                            ? "-"
                            : cimbra > 0
                              ? cimbra.toFixed(2)
                              : "-"}
                        </td>
                        <td className="p-1 border-r text-center font-black text-orange-700 bg-orange-50 text-sm">
                          {e.tipo?.toLowerCase() === "nervadura"
                            ? "-"
                            : cimbraFrontera > 0
                              ? cimbraFrontera.toFixed(2)
                              : "-"}
                        </td>
                        <td
                          className={`p-0 border-r ${e.tipo?.toLowerCase().includes("vigueta") ? "bg-slate-50 cursor-not-allowed" : "bg-slate-100/50 hover:bg-slate-200 cursor-pointer"} transition-colors`}
                          onClick={() => {
                            if (!e.tipo?.toLowerCase().includes("vigueta")) {
                              setActiveSteelSubmodal(e.id);
                            }
                          }}
                        >
                          <div
                            className={`w-full h-full p-2 text-center font-black flex justify-center items-center gap-1 group ${e.tipo?.toLowerCase().includes("vigueta") ? "opacity-30" : "text-slate-700"}`}
                          >
                            {!e.tipo?.toLowerCase().includes("vigueta") && aceroKg > 0 ? (
                              <span className="text-sm">
                                {aceroKg.toFixed(2)}
                              </span>
                            ) : (
                              <span className="text-slate-400 font-normal text-[10px]">
                                {e.tipo?.toLowerCase().includes("vigueta") ? "-" : "0.00"}
                              </span>
                            )}
                            <Wrench
                              size={12}
                              className="text-slate-300 group-hover:text-slate-600 transition-colors"
                            />
                          </div>
                        </td>
                        <td className="p-1 border-l text-center">
                          <div className="flex justify-center items-center gap-0.5">
                            <button
                              onClick={() => handleCopyStructureRow(e)}
                              className="p-1 rounded text-slate-400 hover:text-blue-600"
                            >
                              <Copy size={14} />
                            </button>
                            <button
                              onClick={() => handlePasteStructureRow(e.id)}
                              disabled={!copiedStructureRow}
                              className="p-1 rounded text-slate-400 disabled:opacity-10"
                            >
                              <ClipboardPaste size={14} />
                            </button>
                            <button
                              onClick={() =>
                                updateActiveEstructuras((prev) =>
                                  prev.filter((x) => x.id !== e.id),
                                )
                              }
                              className="p-1 rounded text-slate-400 hover:text-red-500"
                            >
                              <Trash2 size={14} />
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <div className="p-4 bg-white border-t border-slate-100 flex gap-3 flex-wrap shrink-0">
              <button
                onClick={() =>
                  updateActiveEstructuras((prev) => {
                    const lastRow = prev[prev.length - 1];
                    const nextClave = lastRow
                      ? getNextClaveValue(lastRow.clave)
                      : "";
                    return [
                      ...prev,
                      {
                        id: `E-${Date.now()}`,
                        eje: lastRow?.eje || "",
                        clave: nextClave,
                        tipo: lastRow?.tipo || "",
                        i: false,
                        largo: "",
                        ancho: "",
                        alto: "",
                        piezas: "1",
                        aceros: [],
                      },
                    ];
                  })
                }
                className="px-6 py-2 border-2 border-dashed border-amber-300 text-amber-600 font-black rounded-lg hover:bg-amber-50 text-[10px] uppercase tracking-wider"
              >
                + Agregar Elemento
              </button>
              <button
                onClick={() =>
                  updateActiveEstructuras((prev) => {
                    const lastRow = prev[prev.length - 1];
                    let currentClave = lastRow ? lastRow.clave : "";
                    const newRows = [];
                    for (let i = 0; i < 5; i++) {
                      const next = getNextClaveValue(currentClave);
                      newRows.push({
                        id: `E-${Date.now()}-${i}`,
                        eje: lastRow?.eje || "",
                        clave: next,
                        tipo: lastRow?.tipo || "",
                        i: false,
                        largo: "",
                        ancho: "",
                        alto: "",
                        piezas: "1",
                        aceros: [],
                      });
                      currentClave = next;
                    }
                    return [...prev, ...newRows];
                  })
                }
                className="px-6 py-2 border-2 border-dashed border-amber-300 text-amber-600 font-black rounded-lg hover:bg-amber-50 text-[10px] uppercase tracking-wider"
              >
                + Agregar 5 Elementos
              </button>
              <button
                onClick={handleCopySelectedStructures}
                className={`px-6 py-2 rounded-lg font-black text-[10px] uppercase flex items-center gap-2 transition-colors ${selectedStructureRows.length > 0 ? "bg-blue-600 text-white hover:bg-blue-700 shadow-md" : "bg-slate-100 hover:bg-slate-200 text-slate-600"}`}
              >
                <Copy size={14} />{" "}
                {copyStatus === "copied-structure-list"
                  ? "¡Copiado!"
                  : selectedStructureRows.length > 0
                    ? `Copiar (${selectedStructureRows.length})`
                    : "Copiar Todas"}
              </button>
              <button
                onClick={handlePasteSelectedStructures}
                disabled={
                  !copiedStructureRowsList ||
                  copiedStructureRowsList.length === 0
                }
                className={`px-6 py-2 rounded-lg font-black text-[10px] uppercase flex items-center gap-2 transition-all ${copiedStructureRowsList && copiedStructureRowsList.length > 0 ? "bg-emerald-100 hover:bg-emerald-200 text-emerald-700 shadow-sm cursor-pointer" : "bg-slate-50 text-slate-300 cursor-not-allowed"}`}
              >
                <ClipboardPaste size={14} /> Pegar Datos
              </button>
              <button
                onClick={() => {
                  const levelNameExport =
                    niveles.find((n) => n.id === activeLevelId)?.nombre ||
                    "TODOS LOS NIVELES";
                  const renderExVal = (v) => {
                    const num = parseFloat(v);
                    if (isNaN(num) || num === 0) return "-";
                    return num.toFixed(2);
                  };

                  let html = `<table style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; text-align: center;">
                      <colgroup>
                         <col width="30" />
                         <col width="50" />
                         <col width="80" />
                         <col width="100" />
                         <col width="150" />
                         <col width="50" />
                         <col width="80" />
                         <col width="80" />
                         <col width="80" />
                         <col width="60" />
                         <col width="100" />
                         <col width="100" />
                         <col width="120" />
                         <col width="100" />
                      </colgroup>
                      <thead>
                      <tr style="height: 15pt;">
                         <td colspan="14" style="border: none; background-color: #ffffff;"></td>
                      </tr>
                      <tr style="height: 20pt;">
                         <td style="border: none; background-color: #ffffff;"></td>
                         <th colspan="13" style="background-color: #9f490e; color: #ffffff; font-size: 12pt; font-weight: bold; text-align: center; vertical-align: middle; border: 1px solid #9f490e; text-transform: uppercase;">
                            GENERADOR DE ESTRUCTURAS Y ACERO - NIVEL: ${levelNameExport}
                         </th>
                      </tr>
                      <tr style="height: 15pt;">
                         <td colspan="14" style="border: none; background-color: #ffffff;"></td>
                      </tr>
                      <tr style="background-color: #f2f6fb; color: #2f5496; font-size: 8pt; font-weight: bold; text-align: center; height: 15pt;">
                         <td style="border: none; background-color: #ffffff;"></td>
                         <th rowspan="2" style="border: 1px solid #b4c6e7; vertical-align: middle;">NO.</th>
                         <th rowspan="2" style="border: 1px solid #b4c6e7; vertical-align: middle;">EJE</th>
                         <th rowspan="2" style="border: 1px solid #b4c6e7; vertical-align: middle;">CLAVE</th>
                         <th rowspan="2" style="border: 1px solid #b4c6e7; vertical-align: middle;">TIPO ELEMENTO</th>
                         <th rowspan="2" style="border: 1px solid #b4c6e7; vertical-align: middle;">I</th>
                         <th rowspan="2" style="border: 1px solid #b4c6e7; vertical-align: middle;">LARGO (M)</th>
                         <th rowspan="2" style="border: 1px solid #b4c6e7; vertical-align: middle;">ANCHO (M)</th>
                         <th rowspan="2" style="border: 1px solid #b4c6e7; vertical-align: middle;">ALTO (M)</th>
                         <th rowspan="2" style="border: 1px solid #b4c6e7; vertical-align: middle;">PZAS</th>
                         <th colspan="5" style="border: 1px solid #b4c6e7; background-color: #deebf7; color: #2f5496; vertical-align: middle;">VOLÚMENES / ÁREAS TOTALES</th>
                      </tr>
                      <tr style="font-size: 8pt; font-weight: bold; height: 15pt; text-align: center;">
                         <td style="border: none; background-color: #ffffff;"></td>
                         <th style="background-color: #deebf7; color: #2f5496; border: 1px solid #b4c6e7; vertical-align: middle;">CONCRETO (M3)</th>
                         <th style="background-color: #deebf7; color: #2f5496; border: 1px solid #b4c6e7; vertical-align: middle;">CASETONES (M3)</th>
                         <th style="background-color: #fdf3e8; color: #c65911; border: 1px solid #b4c6e7; vertical-align: middle;">CIMBRA (M2)</th>
                         <th style="background-color: #fdf3e8; color: #c65911; border: 1px solid #b4c6e7; vertical-align: middle;">CIM. FRONT(ML)</th>
                         <th style="background-color: #f2f6fb; color: #44546a; border: 1px solid #b4c6e7; vertical-align: middle;">ACERO (KG)</th>
                      </tr>
                   </thead><tbody>`;

                  estructuras.forEach((e, i) => {
                    const concreto = calcConcreto(e);
                    const cimbra = calcCimbra(e);
                    const cimbraFrontera = calcCimbraFrontera(e);
                    const aceroKg = calcAceroTotalKg(e.aceros, e.piezas);
                    const isLosaNerv = (e.tipo || "").toLowerCase() === "losa nervada";
                    const casetones = isLosaNerv ? getCasetonesTotalVol(e.casetones || [], e.piezas) : 0;

                    html += `<tr style="text-align: center; background-color: #ffffff; height: 15pt;">
                         <td style="border: none; background-color: #ffffff;"></td>
                         <td style="border: 1px solid #b4c6e7; color: #a6a6a6; font-size: 8pt; font-weight: bold; vertical-align: middle; mso-number-format:'0';">${i + 1}</td>
                         <td style="border: 1px solid #b4c6e7; color: #2f5496; font-weight: bold; text-transform: uppercase; font-size: 8pt; vertical-align: middle;">${e.eje || ""}</td>
                         <td style="border: 1px solid #b4c6e7; color: #2f5496; font-weight: bold; text-transform: uppercase; font-size: 8pt; vertical-align: middle;">${e.clave || ""}</td>
                         <td style="border: 1px solid #b4c6e7; color: #595959; font-weight: bold; text-transform: uppercase; font-size: 8pt; vertical-align: middle;">${e.tipo || ""}</td>
                         <td style="border: 1px solid #b4c6e7; color: #2f5496; font-weight: bold; font-size: 8pt; vertical-align: middle; text-transform: uppercase;">${e.i ? "SI" : "NO"}</td>
                         <td style="border: 1px solid #b4c6e7; color: #2f5496; font-weight: bold; font-size: 8pt; vertical-align: middle; mso-number-format:'0\\.00';">${renderExVal(e.largo)}</td>
                         <td style="border: 1px solid #b4c6e7; color: #2f5496; font-weight: bold; font-size: 8pt; vertical-align: middle; mso-number-format:'0\\.00';">${renderExVal(e.ancho)}</td>
                         <td style="border: 1px solid #b4c6e7; color: #2f5496; font-weight: bold; font-size: 8pt; vertical-align: middle; mso-number-format:'0\\.00';">${renderExVal(e.alto)}</td>
                         <td style="border: 1px solid #b4c6e7; color: #000000; font-weight: bold; font-size: 8pt; vertical-align: middle; mso-number-format:'0\\.00';">${renderExVal(e.piezas || 1)}</td>
                         <td style="border: 1px solid #b4c6e7; color: #2f5496; font-weight: bold; font-size: 8pt; vertical-align: middle; mso-number-format:'0\\.00';">${concreto > 0 ? concreto.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #b4c6e7; color: #2f5496; font-weight: bold; font-size: 8pt; vertical-align: middle; mso-number-format:'0\\.00';">${casetones > 0 || isLosaNerv ? casetones.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #b4c6e7; color: #e36c09; font-weight: bold; font-size: 8pt; vertical-align: middle; mso-number-format:'0\\.00';">${cimbra > 0 ? cimbra.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #b4c6e7; color: #e36c09; font-weight: bold; font-size: 8pt; vertical-align: middle; mso-number-format:'0\\.00';">${cimbraFrontera > 0 ? cimbraFrontera.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #b4c6e7; color: #44546a; font-weight: bold; font-size: 8pt; vertical-align: middle; mso-number-format:'0\\.00';">${aceroKg > 0 ? aceroKg.toFixed(2) : "-"}</td>
                      </tr>`;
                  });
                  html += `</tbody></table><br/><br/>`;

                  html += `<table style="border-collapse: collapse; font-family: Arial, sans-serif; text-align: center;">
                      <thead>
                         <tr style="height: 15pt;">
                            <td style="width: 30px; border: none;"></td>
                            <th colspan="5" style="background-color: #0d705a; color: white; font-size: 12pt; padding: 6px; border: 1px solid #0d705a; vertical-align: middle;">
                               VOLÚMENES GENERALES DE CIMENTACIÓN
                            </th>
                         </tr>
                         <tr style="background-color: #caffeb; color: #0d705a; font-size: 9pt; height: 15pt;">
                            <td style="border: none; background-color: #ffffff;"></td>
                            <th style="border: 1px solid #7bc5b2; width: 150px; vertical-align: middle;">CONCRETO (M3)</th>
                            <th style="border: 1px solid #7bc5b2; width: 120px; vertical-align: middle;">PLANTILLA (M2)</th>
                            <th style="border: 1px solid #7bc5b2; width: 120px; vertical-align: middle;">EXCAVACIÓN (M3)</th>
                            <th style="border: 1px solid #7bc5b2; width: 120px; vertical-align: middle;">AFINE (M2)</th>
                            <th style="border: 1px solid #7bc5b2; width: 120px; vertical-align: middle;">RELLENO (M3)</th>
                         </tr>
                      </thead>
                      <tbody>
                         <tr style="text-align: center; background-color: #ffffff; height: 15pt;">
                            <td style="border: none;"></td>
                            <td style="border: 1px solid #7bc5b2; font-weight: bold; color: #0d705a; font-size: 10pt; mso-number-format:'0\\.00'; vertical-align: middle;">${estructuraSummary.totalConcretoCimentacion.toFixed(2)}</td>
                            <td style="border: 1px solid #7bc5b2; font-weight: bold; color: #0d705a; font-size: 10pt; mso-number-format:'0\\.00'; vertical-align: middle;">${estructuraSummary.totalPlantilla.toFixed(2)}</td>
                            <td style="border: 1px solid #7bc5b2; font-weight: bold; color: #0d705a; font-size: 10pt; mso-number-format:'0\\.00'; vertical-align: middle;">${estructuraSummary.totalExcavacion.toFixed(2)}</td>
                            <td style="border: 1px solid #7bc5b2; font-weight: bold; color: #0d705a; font-size: 10pt; mso-number-format:'0\\.00'; vertical-align: middle;">${estructuraSummary.totalAfine.toFixed(2)}</td>
                            <td style="border: 1px solid #7bc5b2; font-weight: bold; color: #0d705a; font-size: 10pt; mso-number-format:'0\\.00'; vertical-align: middle;">${estructuraSummary.totalRelleno.toFixed(2)}</td>
                         </tr>
                      </tbody>
                   </table><br/>`;

                  html += `<table style="border-collapse: collapse; font-family: Arial, sans-serif; text-align: center;">
                      <thead>
                         <tr style="height: 15pt;">
                            <td style="width: 30px; border: none;"></td>
                            <th colspan="20" style="background-color: #e36c09; color: white; font-size: 12pt; padding: 6px; border: 1px solid #e36c09; vertical-align: middle;">
                               DESGLOSE POR GRUPOS PARA GENERADORES
                            </th>
                         </tr>
                         <tr style="background-color: #fdf3e8; color: #c65911; font-size: 9pt; height: 15pt;">
                            <td style="border: none; background-color: #ffffff;"></td>
                            <th style="border: 1px solid #f8cbad; width: 250px; vertical-align: middle;">GRUPO / TIPO</th>
                            <th style="border: 1px solid #f8cbad; width: 120px; vertical-align: middle;">CONCRETO (M3)</th>
                            <th style="border: 1px solid #f8cbad; width: 120px; vertical-align: middle;">CASETONES (M3)</th>
                            <th style="border: 1px solid #f8cbad; width: 120px; vertical-align: middle;">CIMBRA (M2)</th>
                            <th style="border: 1px solid #f8cbad; width: 120px; vertical-align: middle;">CIMBRA 0-3M</th>
                            <th style="border: 1px solid #f8cbad; width: 120px; vertical-align: middle;">CIMBRA 3-6M</th>
                            <th style="border: 1px solid #f8cbad; width: 120px; vertical-align: middle;">CIMBRA &gt;6M</th>
                            <th style="border: 1px solid #f8cbad; width: 120px; vertical-align: middle;">CIMB CURVA 0-3M</th>
                             <th style="border: 1px solid #f8cbad; width: 120px; vertical-align: middle;">CIMB CURVA 3-6M</th>
                             <th style="border: 1px solid #f8cbad; width: 120px; vertical-align: middle;">CIMB CURVA &gt;6M</th>
                             <th style="border: 1px solid #f8cbad; width: 140px; vertical-align: middle;">CIM. FRONT (ML)</th>
                             <th style="border: 1px solid #f8cbad; width: 150px; vertical-align: middle;">CIMBRA ESCAL. (M2/ML)</th>
                             <th style="border: 1px solid #f8cbad; width: 150px; vertical-align: middle;">CIMBRA RAMPA (M2/ML)</th>
                             <th style="border: 1px solid #f8cbad; width: 120px; vertical-align: middle;">ANDAMIAJE</th>
                            <th style="border: 1px solid #f8cbad; width: 120px; vertical-align: middle;">OBTURACIÓN (M2)</th>
                            <th style="border: 1px solid #f8cbad; width: 120px; vertical-align: middle;">MALLA REF. (M2)</th>
                            <th style="border: 1px solid #f8cbad; width: 120px; vertical-align: middle;">PASOS MUROS (PZA)</th>
                            <th style="border: 1px solid #f8cbad; width: 120px; vertical-align: middle;">EMPLAYER (M2)</th>
                            <th style="border: 1px solid #f8cbad; width: 140px; vertical-align: middle;">ANCLAJES VERT. (PZA)</th>
                            <th style="border: 1px solid #f8cbad; width: 140px; vertical-align: middle;">ACERO TOTAL (KG)</th>
                         </tr>
                      </thead>
                      <tbody>`;

                  const orderArr = [
                    "Dado",
                    "Zapatas",
                    "Contratrabes",
                    "Columnas",
                    "Columnas Circulares",
                    "Losas de Vigueta",
                    "Losas",
                    "Trabes",
                    "Nervaduras",
                    "Muros"
                  ];
                  Object.entries(estructuraSummary.breakdown)
                    .sort(([tipoA, dataA], [tipoB, dataB]) => {
                      if (dataA.isCimentacion && !dataB.isCimentacion)
                        return -1;
                      if (!dataA.isCimentacion && dataB.isCimentacion) return 1;
                      
                      const idxA = orderArr.indexOf(tipoA);
                      const idxB = orderArr.indexOf(tipoB);
                      
                      if (idxA !== -1 && idxB !== -1) return idxA - idxB;
                      if (idxA !== -1) return -1;
                      if (idxB !== -1) return 1;
                      
                      return tipoA.localeCompare(tipoB);
                    })
                    .forEach(([tipo, data]) => {
                      const hasAnclajes =
                        data.anclajesDetalle &&
                        Object.values(data.anclajesDetalle).some((v) => v > 0);
                      const anclajesStr = hasAnclajes
                        ? Object.entries(data.anclajesDetalle)
                            .filter(([, v]) => v > 0)
                            .map(([k, v]) => `#${k}: ${v.toFixed(0)}`)
                            .join("\n")
                        : "-";

                      if (
                        data.concreto === 0 &&
                        data.cimbra === 0 &&
                        data.cimbraFrontera === 0 &&
                        data.aceroKg === 0 &&
                        data.casetones === 0 &&
                        data.excavacion === 0 &&
                        data.plantilla === 0 &&
                        data.obturacionMuros === 0 &&
                        data.mallaRefuerzo === 0 &&
                        data.pasosMuros === 0 &&
                        data.emplayerColumnas === 0 &&
                        data.andamiajeColumnas === 0 &&
                        data.andamiajeTrabes === 0 &&
                        data.cimbraMuros_0_3 === 0 &&
                        data.cimbraMuros_3_6 === 0 &&
                        data.cimbraMuros_6_9 === 0 &&
                        data.cimbraColumnas_0_3 === 0 &&
                        data.cimbraColumnas_3_6 === 0 &&
                        data.cimbraColumnas_6_9 === 0 &&
                        !hasAnclajes
                      )
                        return;
                      const isCim = data.isCimentacion;
                      const concText = isCim
                        ? "-"
                        : data.concreto > 0
                          ? data.concreto.toFixed(2)
                          : "-";

                      const isColumnGrp =
                        tipo === "Columnas" || tipo === "Columnas Circulares";
                      const andamiajeVal = isColumnGrp
                        ? data.andamiajeColumnas
                        : tipo === "Trabes"
                          ? data.andamiajeTrabes
                          : 0;

                      html += `<tr style="text-align: center; background-color: #ffffff; height: 15pt;">
                         <td style="border: none;"></td>
                         <td style="border: 1px solid #f8cbad; font-weight: bold; text-align: left; padding-left: 8px; color: #000000; font-size: 10pt; vertical-align: middle;">${tipo}</td>
                         <td style="border: 1px solid #f8cbad; color: #2f5496; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${concText}</td>
                         <td style="border: 1px solid #f8cbad; color: #2f5496; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${data.casetones > 0 || data.hasLosaNervada ? data.casetones.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #e36c09; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${data.cimbra > 0 && tipo !== "Muros" && !isColumnGrp ? data.cimbra.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #e36c09; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${tipo === "Muros" && data.cimbraMuros_0_3 > 0 ? data.cimbraMuros_0_3.toFixed(2) : isColumnGrp && data.cimbraColumnas_0_3 > 0 ? data.cimbraColumnas_0_3.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #e36c09; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${tipo === "Muros" && data.cimbraMuros_3_6 > 0 ? data.cimbraMuros_3_6.toFixed(2) : isColumnGrp && data.cimbraColumnas_3_6 > 0 ? data.cimbraColumnas_3_6.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #e36c09; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${tipo === "Muros" && data.cimbraMuros_6_9 > 0 ? data.cimbraMuros_6_9.toFixed(2) : isColumnGrp && data.cimbraColumnas_6_9 > 0 ? data.cimbraColumnas_6_9.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #e36c09; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${tipo === "Muros" && data.cimbraMuroCurvo_0_3 > 0 ? data.cimbraMuroCurvo_0_3.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #e36c09; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${tipo === "Muros" && data.cimbraMuroCurvo_3_6 > 0 ? data.cimbraMuroCurvo_3_6.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #e36c09; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${tipo === "Muros" && data.cimbraMuroCurvo_6_9 > 0 ? data.cimbraMuroCurvo_6_9.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #e36c09; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${data.cimbraFrontera > 0 ? data.cimbraFrontera.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #e36c09; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${data.cimbraEscaleraPapelillo > 0 ? data.cimbraEscaleraPapelillo.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #e36c09; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${data.cimbraRampaEscalera > 0 ? data.cimbraRampaEscalera.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #548235; font-weight: bold; font-size: 10pt; vertical-align: middle; ${isColumnGrp ? "" : `mso-number-format:'0\\.00';`}">${andamiajeVal > 0 ? (isColumnGrp ? andamiajeVal : andamiajeVal.toFixed(2)) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #548235; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${data.obturacionMuros > 0 ? data.obturacionMuros.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #548235; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${data.mallaRefuerzo > 0 ? data.mallaRefuerzo.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #548235; font-weight: bold; font-size: 10pt; vertical-align: middle;">${data.pasosMuros > 0 ? data.pasosMuros : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #548235; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${data.emplayerColumnas > 0 ? data.emplayerColumnas.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #548235; font-weight: bold; font-size: 10pt; vertical-align: middle;">${anclajesStr}</td>
                         <td style="border: 1px solid #f8cbad; color: #2f5496; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${data.aceroKg > 0 && tipo !== "Losas de Vigueta" ? data.aceroKg.toFixed(2) : "-"}</td>
                      </tr>`;
                    });
                  html += `</tbody></table>`;

                  exportFormattedExcel(
                    html,
                    `Generador_Estructuras_${niveles.find((n) => n.id === activeLevelId)?.nombre || "Nivel"}`,
                  );
                }}
                className="px-6 py-2 bg-emerald-600 text-white font-black rounded-lg hover:bg-emerald-500 text-[10px] uppercase tracking-wider flex items-center gap-2 ml-auto"
              >
                <FileDown size={14} /> Exportar Excel
              </button>
            </div>
            <div className="bg-slate-50 p-6 border-t border-slate-200 shrink-0">
              <div className="flex items-center justify-between mb-6">
                <h3 className="text-sm font-black uppercase tracking-widest text-slate-800 flex items-center gap-2">
                  <LayoutDashboard size={18} className="text-amber-600" />{" "}
                  Resumen y Desglose para Exportación
                </h3>
              </div>
              <div className="flex flex-col lg:flex-row gap-8">
                <div className="w-full lg:w-[30%] flex flex-col gap-4">
                  <div className="bg-white rounded-2xl p-5 border border-slate-100 shadow-sm relative overflow-hidden">
                    <div className="absolute left-0 top-0 bottom-0 w-2 bg-teal-400"></div>
                    <span className="text-[10px] font-black text-teal-600 uppercase tracking-widest block mb-1">
                      Concreto Cimentación
                    </span>
                    <div className="flex justify-between items-end">
                      <span className="text-3xl font-black text-slate-800">
                        {estructuraSummary.totalConcretoCimentacion.toFixed(2)}
                      </span>
                      <span className="text-teal-400 font-black text-xl">
                        m3
                      </span>
                    </div>
                  </div>
                  <div className="bg-white rounded-2xl p-5 border border-slate-100 shadow-sm relative overflow-hidden">
                    <div className="absolute left-0 top-0 bottom-0 w-2 bg-blue-400"></div>
                    <span className="text-[10px] font-black text-blue-600 uppercase tracking-widest block mb-1">
                      Concreto Estructura
                    </span>
                    <div className="flex justify-between items-end">
                      <span className="text-3xl font-black text-slate-800">
                        {estructuraSummary.totalConcretoEstructura.toFixed(2)}
                      </span>
                      <span className="text-blue-400 font-black text-xl">
                        m3
                      </span>
                    </div>
                  </div>
                  <div className="bg-white rounded-2xl p-5 border border-slate-100 shadow-sm relative overflow-hidden">
                    <div className="absolute left-0 top-0 bottom-0 w-2 bg-indigo-400"></div>
                    <span className="text-[10px] font-black text-indigo-600 uppercase tracking-widest block mb-1">
                      Concreto Columnas
                    </span>
                    <div className="flex justify-between items-end">
                      <span className="text-3xl font-black text-slate-800">
                        {estructuraSummary.totalConcretoColumnas.toFixed(2)}
                      </span>
                      <span className="text-indigo-400 font-black text-xl">
                        m3
                      </span>
                    </div>
                  </div>
                  <div className="bg-white rounded-2xl p-5 border border-slate-100 shadow-sm relative overflow-hidden">
                    <div className="absolute left-0 top-0 bottom-0 w-2 bg-purple-400"></div>
                    <span className="text-[10px] font-black text-purple-600 uppercase tracking-widest block mb-1">
                      Concreto Muros
                    </span>
                    <div className="flex justify-between items-end">
                      <span className="text-3xl font-black text-slate-800">
                        {estructuraSummary.totalConcretoMuros.toFixed(2)}
                      </span>
                      <span className="text-purple-400 font-black text-xl">
                        m3
                      </span>
                    </div>
                  </div>
                  <div className="bg-white rounded-2xl p-5 border border-slate-100 shadow-sm relative overflow-hidden">
                    <div className="absolute left-0 top-0 bottom-0 w-2 bg-amber-400"></div>
                    <span className="text-[10px] font-black text-amber-600 uppercase tracking-widest block mb-1">
                      Total Cimbra
                    </span>
                    <div className="flex justify-between items-end">
                      <span className="text-3xl font-black text-slate-800">
                        {estructuraSummary.totalCimbra.toFixed(2)}
                      </span>
                      <span className="text-amber-400 font-black text-xl">
                        m2/ml
                      </span>
                    </div>
                  </div>

                  {estructuraSummary.totalAreaViguetasGlobal > 0 && (
                    <div className="bg-white rounded-2xl p-5 border border-slate-100 shadow-sm relative overflow-hidden">
                      <div className="absolute left-0 top-0 bottom-0 w-2 bg-teal-600"></div>
                      <span className="text-[10px] font-black text-teal-600 uppercase tracking-widest block mb-1">
                        Total Área Losa Vigueta
                      </span>
                      <div className="flex justify-between items-end">
                        <span className="text-3xl font-black text-slate-800">
                          {estructuraSummary.totalAreaViguetasGlobal.toFixed(2)}
                        </span>
                        <span className="text-teal-600 font-black text-xl">
                          m²
                        </span>
                      </div>
                    </div>
                  )}

                  {estructuraSummary.totalViguetaH25 > 0 && (
                    <div className="bg-white rounded-2xl p-5 border border-slate-100 shadow-sm relative overflow-hidden">
                      <div className="absolute left-0 top-0 bottom-0 w-2 bg-teal-400"></div>
                      <span className="text-[10px] font-black text-teal-600 uppercase tracking-widest block mb-1">
                        Losa Vigueta h=25
                      </span>
                      <div className="flex justify-between items-end">
                        <span className="text-3xl font-black text-slate-800">
                          {estructuraSummary.totalViguetaH25.toFixed(2)}
                        </span>
                        <span className="text-teal-400 font-black text-xl">
                          m²
                        </span>
                      </div>
                    </div>
                  )}

                  {estructuraSummary.totalViguetaH35 > 0 && (
                    <div className="bg-white rounded-2xl p-5 border border-slate-100 shadow-sm relative overflow-hidden">
                      <div className="absolute left-0 top-0 bottom-0 w-2 bg-teal-500"></div>
                      <span className="text-[10px] font-black text-teal-700 uppercase tracking-widest block mb-1">
                        Losa Vigueta h=35
                      </span>
                      <div className="flex justify-between items-end">
                        <span className="text-3xl font-black text-slate-800">
                          {estructuraSummary.totalViguetaH35.toFixed(2)}
                        </span>
                        <span className="text-teal-500 font-black text-xl">
                          m²
                        </span>
                      </div>
                    </div>
                  )}

                  {estructuraSummary.totalDobleViguetaH25 > 0 && (
                    <div className="bg-white rounded-2xl p-5 border border-slate-100 shadow-sm relative overflow-hidden">
                      <div className="absolute left-0 top-0 bottom-0 w-2 bg-teal-600"></div>
                      <span className="text-[10px] font-black text-teal-800 uppercase tracking-widest block mb-1">
                        Doble Vigueta h=25
                      </span>
                      <div className="flex justify-between items-end">
                        <span className="text-3xl font-black text-slate-800">
                          {estructuraSummary.totalDobleViguetaH25.toFixed(2)}
                        </span>
                        <span className="text-teal-600 font-black text-xl">
                          m²
                        </span>
                      </div>
                    </div>
                  )}

                  {estructuraSummary.totalDobleViguetaH35 > 0 && (
                    <div className="bg-white rounded-2xl p-5 border border-slate-100 shadow-sm relative overflow-hidden">
                      <div className="absolute left-0 top-0 bottom-0 w-2 bg-teal-700"></div>
                      <span className="text-[10px] font-black text-teal-900 uppercase tracking-widest block mb-1">
                        Doble Vigueta h=35
                      </span>
                      <div className="flex justify-between items-end">
                        <span className="text-3xl font-black text-slate-800">
                          {estructuraSummary.totalDobleViguetaH35.toFixed(2)}
                        </span>
                        <span className="text-teal-700 font-black text-xl">
                          m²
                        </span>
                      </div>
                    </div>
                  )}

                  {estructuraSummary.totalCasetonesGeneral > 0 && (
                    <div className="bg-white rounded-2xl p-5 border border-slate-100 shadow-sm relative overflow-hidden">
                      <div className="absolute left-0 top-0 bottom-0 w-2 bg-sky-400"></div>
                      <span className="text-[10px] font-black text-sky-600 uppercase tracking-widest block mb-1">
                        Total Casetones
                      </span>
                      <div className="flex justify-between items-end">
                        <span className="text-3xl font-black text-slate-800">
                          {estructuraSummary.totalCasetonesGeneral.toFixed(2)}
                        </span>
                        <span className="text-sky-500 font-black text-xl">
                          m3
                        </span>
                      </div>
                    </div>
                  )}
                  {estructuraSummary.totalPlantilla > 0 && (
                    <div className="bg-white rounded-2xl p-5 border border-slate-100 shadow-sm relative overflow-hidden">
                      <div className="absolute left-0 top-0 bottom-0 w-2 bg-slate-400"></div>
                      <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest block mb-1">
                        Total Plantilla
                      </span>
                      <div className="flex justify-between items-end">
                        <span className="text-3xl font-black text-slate-800">
                          {estructuraSummary.totalPlantilla.toFixed(2)}
                        </span>
                        <span className="text-slate-400 font-black text-xl">
                          m2
                        </span>
                      </div>
                    </div>
                  )}
                  {estructuraSummary.totalExcavacion > 0 && (
                    <div className="bg-white rounded-2xl p-5 border border-slate-100 shadow-sm relative overflow-hidden">
                      <div className="absolute left-0 top-0 bottom-0 w-2 bg-red-400"></div>
                      <span className="text-[10px] font-black text-red-500 uppercase tracking-widest block mb-1">
                        Total Excavación
                      </span>
                      <div className="flex justify-between items-end">
                        <span className="text-3xl font-black text-slate-800">
                          {estructuraSummary.totalExcavacion.toFixed(2)}
                        </span>
                        <span className="text-red-400 font-black text-xl">
                          m3
                        </span>
                      </div>
                    </div>
                  )}
                  {estructuraSummary.totalAfine > 0 && (
                    <div className="bg-white rounded-2xl p-5 border border-slate-100 shadow-sm relative overflow-hidden">
                      <div className="absolute left-0 top-0 bottom-0 w-2 bg-yellow-400"></div>
                      <span className="text-[10px] font-black text-yellow-600 uppercase tracking-widest block mb-1">
                        Total Afine
                      </span>
                      <div className="flex justify-between items-end">
                        <span className="text-3xl font-black text-slate-800">
                          {estructuraSummary.totalAfine.toFixed(2)}
                        </span>
                        <span className="text-yellow-500 font-black text-xl">
                          m2
                        </span>
                      </div>
                    </div>
                  )}
                  {estructuraSummary.totalRelleno > 0 && (
                    <div className="bg-white rounded-2xl p-5 border border-slate-100 shadow-sm relative overflow-hidden">
                      <div className="absolute left-0 top-0 bottom-0 w-2 bg-lime-400"></div>
                      <span className="text-[10px] font-black text-lime-600 uppercase tracking-widest block mb-1">
                        Total Relleno
                      </span>
                      <div className="flex justify-between items-end">
                        <span className="text-3xl font-black text-slate-800">
                          {estructuraSummary.totalRelleno.toFixed(2)}
                        </span>
                        <span className="text-lime-500 font-black text-xl">
                          m3
                        </span>
                      </div>
                    </div>
                  )}
                </div>

                <div className="w-full lg:w-[70%] bg-white rounded-3xl p-6 border border-slate-200 shadow-sm flex flex-col">
                  <h4 className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-6 border-b border-slate-100 pb-2">
                    Desglose por grupos para generadores
                  </h4>
                  <div className="space-y-4 w-full">
                    {(estructuraSummary.totalConcretoCimentacion > 0 ||
                      estructuraSummary.totalExcavacion > 0) && (
                      <div className="bg-teal-50/50 border border-teal-200 rounded-2xl p-4 flex flex-col xl:flex-row items-center gap-4 shadow-sm">
                        <div className="font-black text-teal-900 uppercase tracking-widest text-sm w-40 shrink-0 flex flex-col">
                          <span>Volúmenes Cimentación</span>
                          <span className="text-[8px] opacity-60">
                            (Zapatas, Dados, CT)
                          </span>
                        </div>
                        <div className="flex flex-wrap gap-3 flex-1">
                          {estructuraSummary.totalConcretoCimentacion > 0 && (
                            <div className="bg-white border border-teal-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                              <div>
                                <span className="block text-[8px] font-black text-teal-600 uppercase">
                                  Concreto
                                </span>
                                <span className="text-sm font-black text-slate-800">
                                  {estructuraSummary.totalConcretoCimentacion.toFixed(
                                    2,
                                  )}{" "}
                                  m3
                                </span>
                              </div>
                              <button
                                onClick={() =>
                                  handleCopyStructure(
                                    "GLOBAL_CIMENTACION",
                                    "concreto",
                                  )
                                }
                                className="text-teal-400 hover:text-teal-600"
                              >
                                <Clipboard size={14} />
                              </button>
                            </div>
                          )}
                          {estructuraSummary.totalPlantilla > 0 && (
                            <div className="bg-white border border-teal-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                              <div>
                                <span className="block text-[8px] font-black text-teal-600 uppercase">
                                  Plantilla
                                </span>
                                <span className="text-sm font-black text-slate-800">
                                  {estructuraSummary.totalPlantilla.toFixed(2)}{" "}
                                  m2
                                </span>
                              </div>
                              <button
                                onClick={() =>
                                  handleCopyStructure(
                                    "GLOBAL_CIMENTACION",
                                    "plantilla",
                                  )
                                }
                                className="text-teal-400 hover:text-teal-600"
                              >
                                <Clipboard size={14} />
                              </button>
                            </div>
                          )}
                          {estructuraSummary.totalExcavacion > 0 && (
                            <div className="bg-white border border-teal-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                              <div>
                                <span className="block text-[8px] font-black text-teal-600 uppercase">
                                  Excavación
                                </span>
                                <span className="text-sm font-black text-slate-800">
                                  {estructuraSummary.totalExcavacion.toFixed(2)}{" "}
                                  m3
                                </span>
                              </div>
                              <button
                                onClick={() =>
                                  handleCopyStructure(
                                    "GLOBAL_CIMENTACION",
                                    "excavacion",
                                  )
                                }
                                className="text-teal-400 hover:text-teal-600"
                              >
                                <Clipboard size={14} />
                              </button>
                            </div>
                          )}
                          {estructuraSummary.totalAfine > 0 && (
                            <div className="bg-white border border-teal-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                              <div>
                                <span className="block text-[8px] font-black text-teal-600 uppercase">
                                  Afine
                                </span>
                                <span className="text-sm font-black text-slate-800">
                                  {estructuraSummary.totalAfine.toFixed(2)} m2
                                </span>
                              </div>
                              <button
                                onClick={() =>
                                  handleCopyStructure(
                                    "GLOBAL_CIMENTACION",
                                    "afine",
                                  )
                                }
                                className="text-teal-400 hover:text-teal-600"
                              >
                                <Clipboard size={14} />
                              </button>
                            </div>
                          )}
                          {estructuraSummary.totalRelleno > 0 && (
                            <div className="bg-white border border-teal-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                              <div>
                                <span className="block text-[8px] font-black text-teal-600 uppercase">
                                  Relleno
                                </span>
                                <span className="text-sm font-black text-slate-800">
                                  {estructuraSummary.totalRelleno.toFixed(2)} m3
                                </span>
                              </div>
                              <button
                                onClick={() =>
                                  handleCopyStructure(
                                    "GLOBAL_CIMENTACION",
                                    "relleno",
                                  )
                                }
                                className="text-teal-400 hover:text-teal-600"
                              >
                                <Clipboard size={14} />
                              </button>
                            </div>
                          )}
                        </div>
                      </div>
                    )}
                    {estructuraSummary.totalAceroCimentacion > 0 && (
                      <div className="bg-teal-50/50 border border-teal-200 rounded-2xl p-4 flex flex-col xl:flex-row items-center gap-4 shadow-sm">
                        <div className="font-black text-teal-900 uppercase tracking-widest text-sm w-40 shrink-0 flex flex-col">
                          <span>Acero Cimentación</span>
                          <span className="text-[8px] opacity-60">
                            (Zapatas, Dados, CT)
                          </span>
                        </div>
                        <div className="flex flex-wrap gap-3 flex-1">
                          <div className="bg-slate-900 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-md border-b-2 border-teal-500">
                            <div>
                              <span className="block text-[8px] font-black text-teal-300 uppercase">
                                Total Acero
                              </span>
                              <span className="text-sm font-black text-white">
                                {estructuraSummary.totalAceroCimentacion.toFixed(
                                  2,
                                )}{" "}
                                kg
                              </span>
                            </div>
                            <button
                              onClick={() =>
                                handleCopyStructure(
                                  "GLOBAL_CIMENTACION",
                                  "acero-total",
                                )
                              }
                              className="text-teal-300 hover:text-white"
                            >
                              <Clipboard size={14} />
                            </button>
                          </div>
                          {Object.entries(
                            estructuraSummary.aceroCimentacionDetalle,
                          )
                            .sort((a, b) => Number(a[0]) - Number(b[0]))
                            .map(([num, kg]) => (
                              <div
                                key={num}
                                className="bg-white border border-teal-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[120px] shadow-sm"
                              >
                                <div>
                                  <span className="block text-[8px] font-black text-teal-600 uppercase">
                                    Acero #{num}
                                  </span>
                                  <span className="text-sm font-black text-slate-800">
                                    {kg.toFixed(2)} kg
                                  </span>
                                </div>
                                <button
                                  onClick={() =>
                                    handleCopyStructure(
                                      "GLOBAL_CIMENTACION",
                                      `acero-${num}`,
                                    )
                                  }
                                  className="text-slate-400 hover:text-teal-600"
                                >
                                  <Clipboard size={14} />
                                </button>
                              </div>
                            ))}
                        </div>
                      </div>
                    )}
                    {Object.entries(estructuraSummary.breakdown).some(
                      ([, data]) =>
                        data.isCimentacion &&
                        (data.cimbra > 0 || data.cimbraFrontera > 0),
                    ) && (
                      <div className="flex flex-wrap gap-4">
                        {Object.entries(estructuraSummary.breakdown)
                          .filter(
                            ([, data]) =>
                              data.isCimentacion &&
                              (data.cimbra > 0 || data.cimbraFrontera > 0),
                          )
                          .sort(([tipoA], [tipoB]) =>
                            tipoA.localeCompare(tipoB),
                          )
                          .map(([tipo, data]) => (
                            <div
                              key={tipo}
                              className="bg-amber-50/50 border border-amber-200 rounded-xl p-3 flex flex-col items-center gap-3 shadow-sm flex-1 min-w-[200px] max-w-[280px]"
                            >
                              <div className="font-black text-amber-900 uppercase tracking-widest text-[10px] text-center w-full">
                                {tipo}
                              </div>
                              <div className="flex flex-col gap-2 w-full">
                                {data.cimbra > 0 && (
                                  <div className="bg-amber-100 border border-amber-200 rounded-lg p-2.5 flex items-center justify-between gap-3 w-full shadow-sm">
                                    <div>
                                      <span className="block text-[8px] font-black text-amber-700 uppercase">
                                        Cimbra
                                      </span>
                                      <span className="text-sm font-black text-amber-900">
                                        {data.cimbra.toFixed(2)} m2/ml
                                      </span>
                                    </div>
                                    <button
                                      onClick={() =>
                                        handleCopyStructure(tipo, "cimbra")
                                      }
                                      className="text-amber-500 hover:text-amber-700"
                                    >
                                      <Clipboard size={14} />
                                    </button>
                                  </div>
                                )}
                                {data.cimbraFrontera > 0 && (
                                  <div className="bg-orange-100 border border-orange-200 rounded-lg p-2.5 flex items-center justify-between gap-3 w-full shadow-sm">
                                    <div>
                                      <span className="block text-[8px] font-black text-orange-700 uppercase">
                                        Cim. Frontera
                                      </span>
                                      <span className="text-sm font-black text-orange-900">
                                        {data.cimbraFrontera.toFixed(2)} ml
                                      </span>
                                    </div>
                                    <button
                                      onClick={() =>
                                        handleCopyStructure(
                                          tipo,
                                          "cimbraFrontera",
                                        )
                                      }
                                      className="text-orange-500 hover:text-orange-700"
                                    >
                                      <Clipboard size={14} />
                                    </button>
                                  </div>
                                )}
                              </div>
                            </div>
                          ))}
                      </div>
                    )}
                    {(estructuraSummary.totalAceroEstructuraGlobal > 0 ||
                      estructuraSummary.totalConcretoEstructura > 0) && (
                      <div className="bg-blue-50/50 border border-blue-200 rounded-2xl p-4 flex flex-col xl:flex-row items-center gap-4 shadow-sm">
                        <div className="font-black text-blue-900 uppercase tracking-widest text-sm w-40 shrink-0 flex flex-col">
                          <span>Acero y Concreto en Estructura</span>
                          <span className="text-[8px] opacity-60">
                            (Losas, Trabes, Col, Col. Circ, Nerv, Esc, Ramp)
                          </span>
                        </div>
                        <div className="flex flex-wrap gap-3 flex-1">
                          {estructuraSummary.totalConcretoEstructura > 0 && (
                            <div className="bg-white border border-blue-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                              <div>
                                <span className="block text-[8px] font-black text-blue-600 uppercase">
                                  Concreto Estructura
                                </span>
                                <span className="text-sm font-black text-slate-800">
                                  {estructuraSummary.totalConcretoEstructura.toFixed(
                                    2,
                                  )}{" "}
                                  m3
                                </span>
                              </div>
                              <button
                                onClick={() =>
                                  handleCopyStructure(
                                    "Concreto en Estructuras",
                                    "concreto",
                                  )
                                }
                                className="text-blue-400 hover:text-blue-600"
                              >
                                <Clipboard size={14} />
                              </button>
                            </div>
                          )}
                          {estructuraSummary.totalAceroEstructuraGlobal > 0 && (
                            <div className="bg-slate-900 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-md border-b-2 border-blue-500">
                              <div>
                                <span className="block text-[8px] font-black text-blue-300 uppercase">
                                  Total Acero
                                </span>
                                <span className="text-sm font-black text-white">
                                  {estructuraSummary.totalAceroEstructuraGlobal.toFixed(
                                    2,
                                  )}{" "}
                                  kg
                                </span>
                              </div>
                              <button
                                onClick={() =>
                                  handleCopyStructure(
                                    "GLOBAL_ESTRUCTURA_ACERO",
                                    "acero-total",
                                  )
                                }
                                className="text-blue-300 hover:text-white"
                              >
                                <Clipboard size={14} />
                              </button>
                            </div>
                          )}
                          {Object.entries(
                            estructuraSummary.aceroEstructuraGlobalDetalle,
                          )
                            .sort((a, b) => Number(a[0]) - Number(b[0]))
                            .map(([num, kg]) => (
                              <div
                                key={num}
                                className="bg-white border border-blue-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[120px] shadow-sm"
                              >
                                <div>
                                  <span className="block text-[8px] font-black text-blue-600 uppercase">
                                    Acero #{num}
                                  </span>
                                  <span className="text-sm font-black text-slate-800">
                                    {kg.toFixed(2)} kg
                                  </span>
                                </div>
                                <button
                                  onClick={() =>
                                    handleCopyStructure(
                                      "GLOBAL_ESTRUCTURA_ACERO",
                                      `acero-${num}`,
                                    )
                                  }
                                  className="text-slate-400 hover:text-blue-600"
                                >
                                  <Clipboard size={14} />
                                </button>
                              </div>
                            ))}
                        </div>
                      </div>
                    )}
                    {Object.entries(estructuraSummary.breakdown).some(
                      ([tipo]) =>
                        tipo === "Columnas" || tipo === "Columnas Circulares",
                    ) && (
                      <div className="flex flex-col md:flex-row gap-4 mt-4 w-full">
                        {Object.entries(estructuraSummary.breakdown)
                          .filter(
                            ([tipo]) =>
                              tipo === "Columnas" ||
                              tipo === "Columnas Circulares",
                          )
                          .sort(([tipoA], [tipoB]) =>
                            tipoA.localeCompare(tipoB),
                          )
                          .map(([tipo, data]) => {
                            if (
                              data.concreto === 0 &&
                              data.cimbra === 0 &&
                              data.cimbraColumnas_0_3 === 0 &&
                              data.cimbraColumnas_3_6 === 0 &&
                              data.cimbraColumnas_6_9 === 0 &&
                              data.aceroKg === 0 &&
                              data.andamiajeColumnas === 0 &&
                              data.emplayerColumnas === 0
                            )
                              return null;

                            return (
                              <div
                                key={tipo}
                                className="bg-amber-50/30 border border-amber-100 rounded-2xl p-4 flex flex-col items-center gap-4 flex-1"
                              >
                                <div className="font-black text-amber-900 uppercase tracking-widest text-sm w-full text-center mb-2">
                                  {tipo}
                                </div>
                                <div className="flex flex-col gap-3 w-full">
                                  {data.concreto > 0 && (
                                    <div className="bg-white border border-amber-200 rounded-xl p-3 flex items-center justify-between gap-4 w-full shadow-sm">
                                      <div>
                                        <span className="block text-[8px] font-black text-amber-600 uppercase">
                                          Concreto
                                        </span>
                                        <span className="text-sm font-black text-slate-800">
                                          {data.concreto.toFixed(2)} m3
                                        </span>
                                      </div>
                                      <button
                                        onClick={() =>
                                          handleCopyStructure(tipo, "concreto")
                                        }
                                        className="text-amber-400 hover:text-amber-600"
                                      >
                                        <Clipboard size={14} />
                                      </button>
                                    </div>
                                  )}
                                  {data.cimbraColumnas_0_3 > 0 && (
                                    <div className="bg-amber-100 border border-amber-200 rounded-xl p-3 flex items-center justify-between gap-4 w-full shadow-sm">
                                      <div>
                                        <span className="block text-[8px] font-black text-amber-700 uppercase">
                                          Cimbra 0-3m
                                        </span>
                                        <span className="text-sm font-black text-amber-900">
                                          {data.cimbraColumnas_0_3.toFixed(2)}{" "}
                                          m2/ml
                                        </span>
                                      </div>
                                      <button
                                        onClick={() =>
                                          handleCopyStructure(
                                            tipo,
                                            "cimbraColumnas_0_3",
                                          )
                                        }
                                        className="text-amber-500 hover:text-amber-700"
                                      >
                                        <Clipboard size={14} />
                                      </button>
                                    </div>
                                  )}
                                  {data.cimbraColumnas_3_6 > 0 && (
                                    <div className="bg-orange-100 border border-orange-200 rounded-xl p-3 flex items-center justify-between gap-4 w-full shadow-sm">
                                      <div>
                                        <span className="block text-[8px] font-black text-orange-700 uppercase">
                                          Cimbra 3-6m
                                        </span>
                                        <span className="text-sm font-black text-orange-900">
                                          {data.cimbraColumnas_3_6.toFixed(2)}{" "}
                                          m2/ml
                                        </span>
                                      </div>
                                      <button
                                        onClick={() =>
                                          handleCopyStructure(
                                            tipo,
                                            "cimbraColumnas_3_6",
                                          )
                                        }
                                        className="text-orange-500 hover:text-orange-700"
                                      >
                                        <Clipboard size={14} />
                                      </button>
                                    </div>
                                  )}
                                  {data.cimbraColumnas_6_9 > 0 && (
                                    <div className="bg-red-100 border border-red-200 rounded-xl p-3 flex items-center justify-between gap-4 w-full shadow-sm">
                                      <div>
                                        <span className="block text-[8px] font-black text-red-700 uppercase">
                                          Cimbra {">"} 6m
                                        </span>
                                        <span className="text-sm font-black text-red-900">
                                          {data.cimbraColumnas_6_9.toFixed(2)}{" "}
                                          m2/ml
                                        </span>
                                      </div>
                                      <button
                                        onClick={() =>
                                          handleCopyStructure(
                                            tipo,
                                            "cimbraColumnas_6_9",
                                          )
                                        }
                                        className="text-red-500 hover:text-red-700"
                                      >
                                        <Clipboard size={14} />
                                      </button>
                                    </div>
                                  )}
                                  {data.andamiajeColumnas > 0 && (
                                    <div className="bg-white border border-slate-200 rounded-xl p-3 flex items-center justify-between gap-4 w-full shadow-sm">
                                      <div>
                                        <span className="block text-[8px] font-black text-slate-500 uppercase">
                                          Andamiaje
                                        </span>
                                        <span className="text-sm font-black text-slate-800">
                                          {typeof data.andamiajeColumnas ===
                                          "number"
                                            ? data.andamiajeColumnas.toFixed(2)
                                            : "Sí"}
                                        </span>
                                      </div>
                                      {typeof data.andamiajeColumnas ===
                                        "number" && (
                                        <button
                                          onClick={() =>
                                            handleCopyStructure(
                                              tipo,
                                              "andamiajeColumnas",
                                            )
                                          }
                                          className="text-slate-400 hover:text-slate-600"
                                        >
                                          <Clipboard size={14} />
                                        </button>
                                      )}
                                    </div>
                                  )}
                                  {data.emplayerColumnas > 0 && (
                                    <div className="bg-white border border-slate-200 rounded-xl p-3 flex items-center justify-between gap-4 w-full shadow-sm">
                                      <div>
                                        <span className="block text-[8px] font-black text-slate-500 uppercase">
                                          Emplayer
                                        </span>
                                        <span className="text-sm font-black text-slate-800">
                                          {data.emplayerColumnas.toFixed(2)} m2
                                        </span>
                                      </div>
                                      <button
                                        onClick={() =>
                                          handleCopyStructure(
                                            tipo,
                                            "emplayerColumnas",
                                          )
                                        }
                                        className="text-slate-400 hover:text-slate-600"
                                      >
                                        <Clipboard size={14} />
                                      </button>
                                    </div>
                                  )}
                                </div>
                              </div>
                            );
                          })}
                      </div>
                    )}
                    {Object.entries(estructuraSummary.breakdown)
                      .sort(([tipoA, dataA], [tipoB, dataB]) => {
                        const orderArr = [
                          "Dado",
                          "Zapatas",
                          "Contratrabes",
                          "Columnas",
                          "Columnas Circulares",
                          "Losas de Vigueta",
                          "Losas",
                          "Trabes",
                          "Nervaduras",
                          "Muros"
                        ];
                        const idxA = orderArr.indexOf(tipoA);
                        const idxB = orderArr.indexOf(tipoB);
                        
                        if (idxA !== -1 && idxB !== -1) return idxA - idxB;
                        if (idxA !== -1) return -1;
                        if (idxB !== -1) return 1;
                        
                        return tipoA.localeCompare(tipoB);
                      })
                      .map(([tipo, data]) => {
                        if (
                          data.isCimentacion ||
                          tipo === "Concreto en Estructuras" ||
                          tipo === "Columnas" ||
                          tipo === "Columnas Circulares"
                        )
                          return null;
                        return (
                          <div
                            key={tipo}
                            className="bg-amber-50/30 border border-amber-100 rounded-2xl p-4 flex flex-col xl:flex-row items-center gap-4"
                          >
                            <div className="font-black text-amber-900 uppercase tracking-widest text-sm w-40 shrink-0">
                              {tipo}
                            </div>
                            <div className="flex flex-wrap gap-3 flex-1">
                              {data.areaViguetasH25 > 0 && (
                                <div className="bg-white border border-slate-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-slate-500 uppercase">
                                      Losa Vigueta h=25
                                    </span>
                                    <span className="text-sm font-black text-slate-800">
                                      {data.areaViguetasH25.toFixed(2)} m2
                                    </span>
                                  </div>
                                  <button
                                    onClick={() => handleCopyStructure(tipo, "areaViguetaH25")}
                                    className="text-slate-400 hover:text-slate-600"
                                  >
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {data.areaViguetasH35 > 0 && (
                                <div className="bg-white border border-slate-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-slate-500 uppercase">
                                      Losa Vigueta h=35
                                    </span>
                                    <span className="text-sm font-black text-slate-800">
                                      {data.areaViguetasH35.toFixed(2)} m2
                                    </span>
                                  </div>
                                  <button
                                    onClick={() => handleCopyStructure(tipo, "areaViguetaH35")}
                                    className="text-slate-400 hover:text-slate-600"
                                  >
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {data.areaDobleViguetaH25 > 0 && (
                                <div className="bg-white border border-slate-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-slate-500 uppercase">
                                      Doble Vigueta h=25
                                    </span>
                                    <span className="text-sm font-black text-slate-800">
                                      {data.areaDobleViguetaH25.toFixed(2)} m2
                                    </span>
                                  </div>
                                  <button
                                    onClick={() => handleCopyStructure(tipo, "areaDobleViguetaH25")}
                                    className="text-slate-400 hover:text-slate-600"
                                  >
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {data.areaDobleViguetaH35 > 0 && (
                                <div className="bg-white border border-slate-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-slate-500 uppercase">
                                      Doble Vigueta h=35
                                    </span>
                                    <span className="text-sm font-black text-slate-800">
                                      {data.areaDobleViguetaH35.toFixed(2)} m2
                                    </span>
                                  </div>
                                  <button
                                    onClick={() => handleCopyStructure(tipo, "areaDobleViguetaH35")}
                                    className="text-slate-400 hover:text-slate-600"
                                  >
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {data.areaViguetasGENERIC > 0 && (
                                <div className="bg-white border border-slate-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-slate-500 uppercase">
                                      Losa Vigueta Genérica
                                    </span>
                                    <span className="text-sm font-black text-slate-800">
                                      {data.areaViguetasGENERIC.toFixed(2)} m2
                                    </span>
                                  </div>
                                  <button
                                    onClick={() => handleCopyStructure(tipo, "areaViguetasGENERIC")}
                                    className="text-slate-400 hover:text-slate-600"
                                  >
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {data.cimbraFrontera > 0 && (
                                <div className="bg-orange-100 border border-orange-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-orange-600 uppercase">
                                      Cim. Frontera
                                    </span>
                                    <span className="text-sm font-black text-orange-800">
                                      {data.cimbraFrontera.toFixed(2)} ml
                                    </span>
                                  </div>
                                  <button
                                    onClick={() =>
                                      handleCopyStructure(
                                        tipo,
                                        "cimbraFrontera",
                                      )
                                    }
                                    className="text-orange-500 hover:text-orange-700"
                                  >
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {data.concreto > 0 && !data.isCimentacion && (
                                <div className="bg-white border border-amber-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-amber-600 uppercase">
                                      Concreto
                                    </span>
                                    <span className="text-sm font-black text-slate-800">
                                      {data.concreto.toFixed(2)} m3
                                    </span>
                                  </div>
                                  <button
                                    onClick={() =>
                                      handleCopyStructure(tipo, "concreto")
                                    }
                                    className="text-amber-400 hover:text-amber-600"
                                  >
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {(data.casetones > 0 || data.hasLosaNervada) && (
                                <div className="bg-sky-50 border border-sky-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-sky-600 uppercase">
                                      Casetones
                                    </span>
                                    <span className="text-sm font-black text-slate-800">
                                      {data.casetones.toFixed(2)} m3
                                    </span>
                                  </div>
                                  <button
                                    onClick={() =>
                                      handleCopyStructure(tipo, "caseton")
                                    }
                                    className="text-sky-400 hover:text-sky-600"
                                  >
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {data.cimbra > 0 &&
                                tipo !== "Muros" &&
                                tipo !== "Columnas" &&
                                tipo !== "Columnas Circulares" &&
                                tipo !== "Losas de Vigueta" && (
                                  <div className="bg-amber-100 border border-amber-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                    <div>
                                      <span className="block text-[8px] font-black text-amber-700 uppercase">
                                        Cimbra
                                      </span>
                                      <span className="text-sm font-black text-amber-900">
                                        {data.cimbra.toFixed(2)} m2/ml
                                      </span>
                                    </div>
                                    <button
                                      onClick={() =>
                                        handleCopyStructure(tipo, "cimbra")
                                      }
                                      className="text-amber-500 hover:text-amber-700"
                                    >
                                      <Clipboard size={14} />
                                    </button>
                                  </div>
                                )}
                              {data.andamiajeTrabes > 0 && (
                                <div className="bg-white border border-slate-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-slate-500 uppercase">
                                      Andamiaje Trabes
                                    </span>
                                    <span className="text-sm font-black text-slate-800">
                                      {data.andamiajeTrabes.toFixed(2)} ml
                                    </span>
                                  </div>
                                  <button
                                    onClick={() =>
                                      handleCopyStructure(
                                        tipo,
                                        "andamiajeTrabes",
                                      )
                                    }
                                    className="text-slate-400 hover:text-slate-600"
                                  >
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {tipo === "Muros" && data.cimbraMuros_0_3 > 0 && (
                                <div className="bg-amber-100 border border-amber-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-amber-700 uppercase">
                                      Cimbra 0-3m
                                    </span>
                                    <span className="text-sm font-black text-amber-900">
                                      {data.cimbraMuros_0_3.toFixed(2)} m2/ml
                                    </span>
                                  </div>
                                  <button
                                    onClick={() =>
                                      handleCopyStructure(
                                        tipo,
                                        "cimbraMuros_0_3",
                                      )
                                    }
                                    className="text-amber-500 hover:text-amber-700"
                                  >
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {tipo === "Muros" && data.cimbraMuros_3_6 > 0 && (
                                <div className="bg-orange-100 border border-orange-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-orange-700 uppercase">
                                      Cimbra 3-6m
                                    </span>
                                    <span className="text-sm font-black text-orange-900">
                                      {data.cimbraMuros_3_6.toFixed(2)} m2/ml
                                    </span>
                                  </div>
                                  <button
                                    onClick={() =>
                                      handleCopyStructure(
                                        tipo,
                                        "cimbraMuros_3_6",
                                      )
                                    }
                                    className="text-orange-500 hover:text-orange-700"
                                  >
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {tipo === "Muros" && data.cimbraMuros_6_9 > 0 && (
                                <div className="bg-red-100 border border-red-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-red-700 uppercase">
                                      Cimbra {">"} 6m
                                    </span>
                                    <span className="text-sm font-black text-red-900">
                                      {data.cimbraMuros_6_9.toFixed(2)} m2/ml
                                    </span>
                                  </div>
                                  <button
                                    onClick={() =>
                                      handleCopyStructure(
                                        tipo,
                                        "cimbraMuros_6_9",
                                      )
                                    }
                                    className="text-red-500 hover:text-red-700"
                                  >
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {tipo === "Muros" &&
                                data.cimbraMuroCurvo_0_3 > 0 && (
                                  <div className="bg-amber-100 border border-amber-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                    <div>
                                      <span className="block text-[8px] font-black text-amber-700 uppercase">
                                        Cimbra Curva 0-3m
                                      </span>
                                      <span className="text-sm font-black text-amber-900">
                                        {data.cimbraMuroCurvo_0_3.toFixed(2)}{" "}
                                        m2/ml
                                      </span>
                                    </div>
                                    <button
                                      onClick={() =>
                                        handleCopyStructure(
                                          tipo,
                                          "cimbraMuroCurvo_0_3",
                                        )
                                      }
                                      className="text-amber-500 hover:text-amber-700"
                                    >
                                      <Clipboard size={14} />
                                    </button>
                                  </div>
                                )}
                              {tipo === "Muros" &&
                                data.cimbraMuroCurvo_3_6 > 0 && (
                                  <div className="bg-orange-100 border border-orange-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                    <div>
                                      <span className="block text-[8px] font-black text-orange-700 uppercase">
                                        Cimbra Curva 3-6m
                                      </span>
                                      <span className="text-sm font-black text-orange-900">
                                        {data.cimbraMuroCurvo_3_6.toFixed(2)}{" "}
                                        m2/ml
                                      </span>
                                    </div>
                                    <button
                                      onClick={() =>
                                        handleCopyStructure(
                                          tipo,
                                          "cimbraMuroCurvo_3_6",
                                        )
                                      }
                                      className="text-orange-500 hover:text-orange-700"
                                    >
                                      <Clipboard size={14} />
                                    </button>
                                  </div>
                                )}
                              {tipo === "Muros" &&
                                data.cimbraMuroCurvo_6_9 > 0 && (
                                  <div className="bg-red-100 border border-red-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                    <div>
                                      <span className="block text-[8px] font-black text-red-700 uppercase">
                                        Cimbra Curva {">"} 6m
                                      </span>
                                      <span className="text-sm font-black text-red-900">
                                        {data.cimbraMuroCurvo_6_9.toFixed(2)}{" "}
                                        m2/ml
                                      </span>
                                    </div>
                                    <button
                                      onClick={() =>
                                        handleCopyStructure(
                                          tipo,
                                          "cimbraMuroCurvo_6_9",
                                        )
                                      }
                                      className="text-red-500 hover:text-red-700"
                                    >
                                      <Clipboard size={14} />
                                    </button>
                                  </div>
                                )}
                              {(tipo === "Columnas" ||
                                tipo === "Columnas Circulares") &&
                                data.cimbraColumnas_0_3 > 0 && (
                                  <div className="bg-amber-100 border border-amber-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                    <div>
                                      <span className="block text-[8px] font-black text-amber-700 uppercase">
                                        Cimbra 0-3m
                                      </span>
                                      <span className="text-sm font-black text-amber-900">
                                        {data.cimbraColumnas_0_3.toFixed(2)}{" "}
                                        m2/ml
                                      </span>
                                    </div>
                                    <button
                                      onClick={() =>
                                        handleCopyStructure(
                                          tipo,
                                          "cimbraColumnas_0_3",
                                        )
                                      }
                                      className="text-amber-500 hover:text-amber-700"
                                    >
                                      <Clipboard size={14} />
                                    </button>
                                  </div>
                                )}
                              {(tipo === "Columnas" ||
                                tipo === "Columnas Circulares") &&
                                data.cimbraColumnas_3_6 > 0 && (
                                  <div className="bg-orange-100 border border-orange-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                    <div>
                                      <span className="block text-[8px] font-black text-orange-700 uppercase">
                                        Cimbra 3-6m
                                      </span>
                                      <span className="text-sm font-black text-orange-900">
                                        {data.cimbraColumnas_3_6.toFixed(2)}{" "}
                                        m2/ml
                                      </span>
                                    </div>
                                    <button
                                      onClick={() =>
                                        handleCopyStructure(
                                          tipo,
                                          "cimbraColumnas_3_6",
                                        )
                                      }
                                      className="text-orange-500 hover:text-orange-700"
                                    >
                                      <Clipboard size={14} />
                                    </button>
                                  </div>
                                )}
                              {(tipo === "Columnas" ||
                                tipo === "Columnas Circulares") &&
                                data.cimbraColumnas_6_9 > 0 && (
                                  <div className="bg-red-100 border border-red-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                    <div>
                                      <span className="block text-[8px] font-black text-red-700 uppercase">
                                        Cimbra {">"} 6m
                                      </span>
                                      <span className="text-sm font-black text-red-900">
                                        {data.cimbraColumnas_6_9.toFixed(2)}{" "}
                                        m2/ml
                                      </span>
                                    </div>
                                    <button
                                      onClick={() =>
                                        handleCopyStructure(
                                          tipo,
                                          "cimbraColumnas_6_9",
                                        )
                                      }
                                      className="text-red-500 hover:text-red-700"
                                    >
                                      <Clipboard size={14} />
                                    </button>
                                  </div>
                                )}
                              {data.obturacionMuros > 0 && (
                                <div className="bg-white border border-slate-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-slate-500 uppercase">
                                      Obturación Muros
                                    </span>
                                    <span className="text-sm font-black text-slate-800">
                                      {data.obturacionMuros.toFixed(2)} m2/ml
                                    </span>
                                  </div>
                                  <button
                                    onClick={() =>
                                      handleCopyStructure(
                                        tipo,
                                        "obturacionMuros",
                                      )
                                    }
                                    className="text-slate-400 hover:text-slate-600"
                                  >
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {data.mallaRefuerzo > 0 && (
                                <div className="bg-white border border-slate-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-slate-500 uppercase">
                                      Malla Refuerzo
                                    </span>
                                    <span className="text-sm font-black text-slate-800">
                                      {data.mallaRefuerzo.toFixed(2)} m2/ml
                                    </span>
                                  </div>
                                  <button
                                    onClick={() =>
                                      handleCopyStructure(tipo, "mallaRefuerzo")
                                    }
                                    className="text-slate-400 hover:text-slate-600"
                                  >
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {data.pasosMuros > 0 && (
                                <div className="bg-white border border-slate-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-slate-500 uppercase">
                                      Pasos Muros
                                    </span>
                                    <span className="text-sm font-black text-slate-800">
                                      {data.pasosMuros} pzas
                                    </span>
                                  </div>
                                  <button
                                    onClick={() =>
                                      handleCopyStructure(tipo, "pasosMuros")
                                    }
                                    className="text-slate-400 hover:text-slate-600"
                                  >
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {data.emplayerColumnas > 0 && (
                                <div className="bg-white border border-slate-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-slate-500 uppercase">
                                      Emplayer
                                    </span>
                                    <span className="text-sm font-black text-slate-800">
                                      {data.emplayerColumnas.toFixed(2)} m2
                                    </span>
                                  </div>
                                  <button
                                    onClick={() =>
                                      handleCopyStructure(
                                        tipo,
                                        "emplayerColumnas",
                                      )
                                    }
                                    className="text-slate-400 hover:text-slate-600"
                                  >
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {data.anclajesDetalle &&
                                Object.entries(data.anclajesDetalle).map(
                                  ([nv, pzas]) => (
                                    <div
                                      key={nv}
                                      className="bg-indigo-50 border border-indigo-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm"
                                    >
                                      <div>
                                        <span className="block text-[8px] font-black text-indigo-500 uppercase">
                                          Anclaje Var #{nv}
                                        </span>
                                        <span className="text-sm font-black text-indigo-900">
                                          {pzas.toFixed(0)} pzas
                                        </span>
                                      </div>
                                      <button
                                        onClick={() =>
                                          handleCopyStructure(
                                            tipo,
                                            `anclajeVar_${nv}`,
                                          )
                                        }
                                        className="text-indigo-400 hover:text-indigo-600"
                                      >
                                        <Clipboard size={14} />
                                      </button>
                                    </div>
                                  ),
                                )}
                              {data.andamiajeColumnas > 0 && (
                                <div className="bg-white border border-slate-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-slate-500 uppercase">
                                      Andamiaje Col ({">"}3m)
                                    </span>
                                    <span className="text-sm font-black text-slate-800">
                                      {data.andamiajeColumnas} pzas
                                    </span>
                                  </div>
                                  <button
                                    onClick={() =>
                                      handleCopyStructure(
                                        tipo,
                                        "andamiajeColumnas",
                                      )
                                    }
                                    className="text-slate-400 hover:text-slate-600"
                                  >
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {data.excavacion > 0 && !data.isCimentacion && (
                                <div className="bg-amber-100 border border-amber-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-amber-700 uppercase">
                                      Excavación
                                    </span>
                                    <span className="text-sm font-black text-amber-900">
                                      {data.excavacion.toFixed(2)} m3
                                    </span>
                                  </div>
                                  <button
                                    onClick={() =>
                                      handleCopyStructure(tipo, "excavacion")
                                    }
                                    className="text-amber-500 hover:text-amber-700"
                                  >
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {data.relleno > 0 && !data.isCimentacion && (
                                <div className="bg-orange-100 border border-orange-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-orange-700 uppercase">
                                      Relleno
                                    </span>
                                    <span className="text-sm font-black text-orange-900">
                                      {data.relleno.toFixed(2)} m3
                                    </span>
                                  </div>
                                  <button
                                    onClick={() =>
                                      handleCopyStructure(tipo, "relleno")
                                    }
                                    className="text-orange-500 hover:text-orange-700"
                                  >
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {data.cimbraEscaleraPapelillo > 0 && (
                                <div className="bg-amber-100 border border-amber-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-amber-700 uppercase">
                                      Cimbra Esc. Papelillo
                                    </span>
                                    <span className="text-sm font-black text-amber-900">
                                      {data.cimbraEscaleraPapelillo.toFixed(2)} m2/ml
                                    </span>
                                  </div>
                                  <button
                                    onClick={() =>
                                      handleCopyStructure(
                                        tipo,
                                        "cimbraEscaleraPapelillo",
                                      )
                                    }
                                    className="text-amber-500 hover:text-amber-700"
                                  >
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {data.cimbraRampaEscalera > 0 && (
                                <div className="bg-amber-100 border border-amber-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-amber-700 uppercase">
                                      Cimbra Rampa Esc.
                                    </span>
                                    <span className="text-sm font-black text-amber-900">
                                      {data.cimbraRampaEscalera.toFixed(2)} m2/ml
                                    </span>
                                  </div>
                                  <button
                                    onClick={() =>
                                      handleCopyStructure(
                                        tipo,
                                        "cimbraRampaEscalera",
                                      )
                                    }
                                    className="text-amber-500 hover:text-amber-700"
                                  >
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {data.aceroKg > 0 &&
                                tipo !== "Losas de Vigueta" &&
                                !data.isCimentacion &&
                                ![
                                  "losa",
                                  "losa nervada",
                                  "trabe",
                                  "nervadura",
                                  "columna",
                                  "columna circular",
                                  "escalera papelillo",
                                  "rampa de escalera",
                                ].includes(data.tipoOriginal) && (
                                  <>
                                    <div className="bg-slate-900 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-md border-b-2 border-slate-500">
                                      <div>
                                        <span className="block text-[8px] font-black text-slate-300 uppercase">
                                          Total Acero
                                        </span>
                                        <span className="text-sm font-black text-white">
                                          {data.aceroKg.toFixed(2)} kg
                                        </span>
                                      </div>
                                      <button
                                        onClick={() =>
                                          handleCopyStructure(
                                            tipo,
                                            "acero-total",
                                          )
                                        }
                                        className="text-slate-300 hover:text-white"
                                      >
                                        <Clipboard size={14} />
                                      </button>
                                    </div>
                                    {Object.entries(data.aceroDetalle || {})
                                      .sort(
                                        (a, b) => Number(a[0]) - Number(b[0]),
                                      )
                                      .map(([num, kg]) => (
                                        <div
                                          key={num}
                                          className="bg-white border border-slate-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[120px] shadow-sm"
                                        >
                                          <div>
                                            <span className="block text-[8px] font-black text-slate-600 uppercase">
                                              Acero #{num}
                                            </span>
                                            <span className="text-sm font-black text-slate-800">
                                              {kg.toFixed(2)} kg
                                            </span>
                                          </div>
                                          <button
                                            onClick={() =>
                                              handleCopyStructure(
                                                tipo,
                                                `acero-${num}`,
                                              )
                                            }
                                            className="text-slate-400 hover:text-slate-600"
                                          >
                                            <Clipboard size={14} />
                                          </button>
                                        </div>
                                      ))}
                                  </>
                                )}
                            </div>
                          </div>
                        );
                      })}
                  </div>
                </div>
              </div>
            </div>
          </div>
        ) : showWallGenerator ? (
          <div className="bg-white rounded-3xl shadow-xl border border-gray-100 animate-in fade-in zoom-in-95 duration-300 flex flex-col">
            <div className="bg-[#312e81] text-white p-4 flex justify-between items-center shrink-0">
              <div className="flex items-center gap-3">
                <div className="p-2 bg-indigo-800/50 rounded-lg border border-indigo-600/50">
                  <Building2 size={20} className="text-indigo-200" />
                </div>
                <div>
                  <h2 className="text-lg font-black uppercase tracking-wider leading-tight">
                    Matriz de Muros
                  </h2>
                  <p className="text-[9px] text-indigo-300 font-bold uppercase tracking-widest">
                    {activePartida.nombre}
                  </p>
                </div>
              </div>
              <div className="flex items-center gap-3">
                <select
                  value={activeLevelId}
                  onChange={(e) => setActiveLevelId(e.target.value)}
                  className="bg-[#1e1b4b] text-white font-black text-[10px] uppercase p-2 px-4 rounded-lg border border-[#4338ca] outline-none cursor-pointer"
                >
                  {niveles.map((n) => (
                    <option key={n.id} value={n.id} className="text-slate-800">
                      {n.nombre}
                    </option>
                  ))}
                </select>
                <button
                  onClick={() => setShowWallGenerator(false)}
                  className="p-1.5 bg-[#4338ca] hover:bg-indigo-500 rounded-full transition-colors"
                >
                  <X size={20} />
                </button>
              </div>
            </div>
            <div
              className={`overflow-x-auto bg-white ${muros.length > 12 ? "max-h-[450px] overflow-y-auto overscroll-contain" : ""}`}
            >
              <table className="w-full text-left border-collapse text-[10px] xl:text-xs">
                <thead className="bg-slate-50 text-slate-500 font-black uppercase tracking-wider border-b-2 border-slate-200 sticky top-0 z-30 shadow-sm text-[9px] md:text-[10px]">
                  <tr>
                    <th
                      className="p-2 text-center border-r border-slate-200 relative select-none bg-slate-50"
                      rowSpan={2}
                      style={{ width: wallColWidths.select || 40 }}
                    >
                      <button
                        onClick={toggleAllWallRows}
                        className="hover:text-indigo-600 transition-colors"
                      >
                        {muros.length > 0 &&
                        selectedWallRows.length === muros.length ? (
                          <CheckSquare size={14} />
                        ) : (
                          <Square size={14} />
                        )}
                      </button>
                    </th>
                    <th
                      className="p-2 text-center border-r border-slate-200"
                      rowSpan={2}
                    >
                      No.
                    </th>
                    <th
                      className="p-2 border-r border-slate-200 text-center"
                      rowSpan={2}
                    >
                      Eje
                    </th>
                    <th
                      className="p-2 border-r border-slate-200 text-center"
                      rowSpan={2}
                    >
                      Clave
                    </th>
                    <th
                      className="p-2 text-center border-r border-slate-200 text-blue-700 bg-slate-50"
                      rowSpan={2}
                    >
                      Largo
                    </th>
                    <th
                      className="p-2 text-center border-r border-slate-200 text-blue-700 bg-slate-50"
                      rowSpan={2}
                    >
                      Ancho
                    </th>
                    <th
                      className="p-2 text-center border-r border-slate-200 text-blue-700 bg-slate-50"
                      rowSpan={2}
                    >
                      Alto
                    </th>
                    <th
                      className="p-2 text-center border-r border-slate-200 bg-slate-200 text-slate-700"
                      rowSpan={2}
                    >
                      Bruto
                    </th>
                    <th
                      className="p-2 text-center border-r border-slate-200 bg-red-50 text-red-600"
                      rowSpan={2}
                    >
                      -Huecos
                    </th>
                    <th
                      className="p-2 text-center border-r border-slate-200 bg-orange-50 text-orange-600"
                      rowSpan={2}
                    >
                      -Cast.
                    </th>
                    <th
                      className="p-2 text-center border-r border-slate-200 bg-emerald-100 text-emerald-800 font-black"
                      rowSpan={2}
                    >
                      Neta
                    </th>
                    <th
                      className="p-1 border-b border-r border-slate-200 text-center text-purple-700"
                      colSpan={2}
                    >
                      Aplanado
                    </th>
                    <th
                      className="p-1 border-b border-r border-slate-200 text-center text-fuchsia-700"
                      colSpan={2}
                    >
                      Recubrimiento
                    </th>
                    <th
                      className="p-2 text-center border-l border-slate-200"
                      rowSpan={2}
                    >
                      <div className="flex justify-center items-center gap-1">
                        <button
                          onClick={handleCopySelectedWalls}
                          className={`p-1 rounded transition-colors ${selectedWallRows.length > 0 ? "text-indigo-600 hover:bg-indigo-100" : "text-slate-300 cursor-not-allowed"}`}
                          disabled={selectedWallRows.length === 0}
                          title="Copiar Seleccionados"
                        >
                          <Copy size={14} />
                        </button>
                        <button
                          onClick={() => {
                            if (selectedWallRows.length > 0) {
                              updateActiveMuros((prev) =>
                                prev.filter(
                                  (x) => !selectedWallRows.includes(x.id),
                                ),
                              );
                              setSelectedWallRows([]);
                            }
                          }}
                          className={`p-1 rounded transition-colors ${selectedWallRows.length > 0 ? "text-red-500 hover:bg-red-50" : "text-slate-300 cursor-not-allowed"}`}
                          disabled={selectedWallRows.length === 0}
                          title="Borrar Seleccionados"
                        >
                          <Trash2 size={14} />
                        </button>
                      </div>
                    </th>
                  </tr>
                  <tr>
                    <th className="p-1 border-r border-slate-200 text-center text-purple-700">
                      C1
                    </th>
                    <th className="p-1 border-r border-slate-200 text-center text-purple-700">
                      C2
                    </th>
                    <th className="p-1 border-r border-slate-200 text-center text-fuchsia-700">
                      C1
                    </th>
                    <th className="p-1 border-r border-slate-200 text-center text-fuchsia-700">
                      C2
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {muros.map((m, i) => {
                    const bruto = calcWallArea(m.largo, m.alto),
                      dedH = getWallHuecosTotal(m.huecos),
                      dedC = getWallCastillosTotal(m.castillos, m.alto),
                      neto = Math.max(0, bruto - dedH - dedC);
                    return (
                      <tr
                        key={m.id}
                        className={`transition-colors ${selectedWallRows.includes(m.id) ? "bg-indigo-50/50" : "hover:bg-slate-50"}`}
                      >
                        <td className="p-1 text-center border-r border-slate-200">
                          <button
                            onClick={() => toggleWallRow(m.id)}
                            className={`${selectedWallRows.includes(m.id) ? "text-indigo-600" : "text-slate-300 hover:text-indigo-500"} transition-colors`}
                          >
                            {selectedWallRows.includes(m.id) ? (
                              <CheckSquare size={14} />
                            ) : (
                              <Square size={14} />
                            )}
                          </button>
                        </td>
                        <td className="p-1.5 text-center font-black text-slate-400 border-r">
                          {i + 1}
                        </td>
                        <td className="p-0 border-r">
                          <DebouncedCell
                            value={m.eje}
                            onChange={(v) =>
                              updateActiveMuros((prev) =>
                                prev.map((x) =>
                                  x.id === m.id ? { ...x, eje: v } : x,
                                ),
                              )
                            }
                            className="w-full p-2 bg-transparent font-bold uppercase outline-none text-[10px] md:text-xs text-center"
                          />
                        </td>
                        <td className="p-0 border-r">
                          <DebouncedCell
                            value={m.clave}
                            onChange={(v) =>
                              updateActiveMuros((prev) =>
                                prev.map((x) =>
                                  x.id === m.id ? { ...x, clave: v } : x,
                                ),
                              )
                            }
                            className="w-full p-2 bg-transparent font-bold uppercase outline-none text-[10px] md:text-xs text-center"
                          />
                        </td>
                        <td className="p-0 border-r bg-blue-50/10">
                          <DebouncedCell
                            value={m.largo}
                            onChange={(v) =>
                              updateActiveMuros((prev) =>
                                prev.map((x) =>
                                  x.id === m.id ? { ...x, largo: v } : x,
                                ),
                              )
                            }
                            className="w-full p-2 bg-transparent text-center text-blue-900 font-bold text-[10px] md:text-xs"
                          />
                        </td>
                        <td className="p-0 border-r bg-blue-50/10">
                          <DebouncedCell
                            value={m.ancho}
                            onChange={(v) =>
                              updateActiveMuros((prev) =>
                                prev.map((x) =>
                                  x.id === m.id ? { ...x, ancho: v } : x,
                                ),
                              )
                            }
                            className="w-full p-2 bg-transparent text-center text-blue-900 font-bold text-[10px] md:text-xs"
                          />
                        </td>
                        <td className="p-0 border-r bg-blue-50/10">
                          <DebouncedCell
                            value={m.alto}
                            onChange={(v) =>
                              updateActiveMuros((prev) =>
                                prev.map((x) =>
                                  x.id === m.id ? { ...x, alto: v } : x,
                                ),
                              )
                            }
                            className="w-full p-2 bg-transparent text-center text-blue-900 font-bold text-[10px] md:text-xs"
                          />
                        </td>
                        <td className="p-2 border-r text-center font-black text-slate-500 bg-slate-50/50 text-[10px] md:text-xs">
                          {bruto > 0 ? bruto.toFixed(2) : "-"}
                        </td>
                        <td
                          className="p-0 border-r bg-red-50/30 hover:bg-red-100 transition-colors cursor-pointer"
                          onClick={() =>
                            setActiveWallSubmodal({
                              wallId: m.id,
                              type: "huecos",
                            })
                          }
                        >
                          <div className="w-full h-full p-2 text-center font-black text-red-600 text-[10px] md:text-xs">
                            {dedH > 0 ? dedH.toFixed(2) : "0.00"}
                          </div>
                        </td>
                        <td
                          className="p-0 border-r bg-orange-50/30 hover:bg-orange-100 transition-colors cursor-pointer"
                          onClick={() =>
                            setActiveWallSubmodal({
                              wallId: m.id,
                              type: "castillos",
                            })
                          }
                        >
                          <div className="w-full h-full p-2 text-center font-black text-orange-600 text-[10px] md:text-xs">
                            {dedC > 0 ? dedC.toFixed(2) : "0.00"}
                          </div>
                        </td>
                        <td className="p-2 border-r text-center font-black text-emerald-700 bg-emerald-50 text-[10px] md:text-sm">
                          {neto > 0 ? neto.toFixed(2) : "-"}
                        </td>
                        <td className="p-0 border-r">
                          <DebouncedCell
                            value={m.tipoAplanadoC1}
                            onChange={(v) =>
                              updateActiveMuros((prev) =>
                                prev.map((x) =>
                                  x.id === m.id
                                    ? { ...x, tipoAplanadoC1: v }
                                    : x,
                                ),
                              )
                            }
                            className="w-full p-2 bg-transparent text-center font-bold text-slate-700 text-[9px] md:text-[10px] uppercase"
                          />
                        </td>
                        <td className="p-0 border-r">
                          <DebouncedCell
                            value={m.tipoAplanadoC2}
                            onChange={(v) =>
                              updateActiveMuros((prev) =>
                                prev.map((x) =>
                                  x.id === m.id
                                    ? { ...x, tipoAplanadoC2: v }
                                    : x,
                                ),
                              )
                            }
                            className="w-full p-2 bg-transparent text-center font-bold text-slate-700 text-[9px] md:text-[10px] uppercase"
                          />
                        </td>
                        <td className="p-0 border-r">
                          <DebouncedCell
                            value={m.tipoRecubrimientoC1}
                            onChange={(v) =>
                              updateActiveMuros((prev) =>
                                prev.map((x) =>
                                  x.id === m.id
                                    ? { ...x, tipoRecubrimientoC1: v }
                                    : x,
                                ),
                              )
                            }
                            className="w-full p-2 bg-transparent text-center font-bold text-slate-700 text-[9px] md:text-[10px] uppercase"
                          />
                        </td>
                        <td className="p-0 border-r">
                          <DebouncedCell
                            value={m.tipoRecubrimientoC2}
                            onChange={(v) =>
                              updateActiveMuros((prev) =>
                                prev.map((x) =>
                                  x.id === m.id
                                    ? { ...x, tipoRecubrimientoC2: v }
                                    : x,
                                ),
                              )
                            }
                            className="w-full p-2 bg-transparent text-center font-bold text-slate-700 text-[9px] md:text-[10px] uppercase"
                          />
                        </td>
                        <td className="p-1 text-center border-l border-slate-200">
                          <div className="flex justify-center items-center gap-0.5">
                            <button
                              onClick={() => handleCopyWallRow(m)}
                              className="p-1 rounded text-slate-400 hover:text-indigo-600"
                            >
                              <Copy size={14} />
                            </button>
                            <button
                              onClick={() => handlePasteWallRow(m.id)}
                              disabled={!copiedWallRow}
                              className="p-1 rounded text-slate-400 disabled:opacity-10"
                            >
                              <ClipboardPaste size={14} />
                            </button>
                            <button
                              onClick={() =>
                                updateActiveMuros((prev) =>
                                  prev.filter((x) => x.id !== m.id),
                                )
                              }
                              className="text-slate-300 hover:text-red-500 p-1"
                            >
                              <Trash2 size={14} />
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <div className="p-4 bg-white border-t border-slate-100 flex gap-3 flex-wrap shrink-0">
              <button
                onClick={() =>
                  updateActiveMuros((prev) => {
                    const lastRow = prev[prev.length - 1];
                    const nextClave = lastRow
                      ? getNextClaveValue(lastRow.clave)
                      : "";
                    return [
                      ...prev,
                      {
                        id: `M-${Date.now()}`,
                        eje: lastRow?.eje || "",
                        clave: nextClave,
                        largo: "",
                        ancho: "",
                        alto: "",
                        huecos: [],
                        castillos: [],
                        tipoAplanadoC1: "",
                        tipoAplanadoC2: "",
                        tipoRecubrimientoC1: "",
                        tipoRecubrimientoC2: "",
                      },
                    ];
                  })
                }
                className="px-6 py-2 border-2 border-dashed border-indigo-200 text-indigo-600 font-black rounded-lg hover:bg-indigo-50 text-[10px] uppercase tracking-wider"
              >
                + Agregar Muro
              </button>
              <button
                onClick={() =>
                  updateActiveMuros((prev) => {
                    const lastRow = prev[prev.length - 1];
                    let currentClave = lastRow ? lastRow.clave : "";
                    const newRows = [];
                    for (let i = 0; i < 5; i++) {
                      const next = getNextClaveValue(currentClave);
                      newRows.push({
                        id: `M-${Date.now()}-${i}`,
                        eje: lastRow?.eje || "",
                        clave: next,
                        largo: "",
                        ancho: "",
                        alto: "",
                        huecos: [],
                        castillos: [],
                        tipoAplanadoC1: "",
                        tipoAplanadoC2: "",
                        tipoRecubrimientoC1: "",
                        tipoRecubrimientoC2: "",
                      });
                      currentClave = next;
                    }
                    return [...prev, ...newRows];
                  })
                }
                className="px-6 py-2 border-2 border-dashed border-indigo-200 text-indigo-600 font-black rounded-lg hover:bg-indigo-50 text-[10px] uppercase tracking-wider"
              >
                + Agregar 5 Muros
              </button>
              <button
                onClick={handleCopySelectedWalls}
                className={`px-6 py-2 rounded-lg font-black text-[10px] uppercase flex items-center gap-2 transition-colors ${selectedWallRows.length > 0 ? "bg-indigo-600 text-white hover:bg-indigo-700 shadow-md" : "bg-slate-100 hover:bg-slate-200 text-slate-600"}`}
              >
                <Copy size={14} />{" "}
                {copyStatus === "copied-wall-list"
                  ? "¡Copiado!"
                  : selectedWallRows.length > 0
                    ? `Copiar (${selectedWallRows.length})`
                    : "Copiar Todas"}
              </button>
              <button
                onClick={handlePasteSelectedWalls}
                disabled={
                  !copiedWallRowsList || copiedWallRowsList.length === 0
                }
                className={`px-6 py-2 rounded-lg font-black text-[10px] uppercase flex items-center gap-2 transition-all ${copiedWallRowsList && copiedWallRowsList.length > 0 ? "bg-emerald-100 hover:bg-emerald-200 text-emerald-700 shadow-sm cursor-pointer" : "bg-slate-50 text-slate-300 cursor-not-allowed"}`}
              >
                <ClipboardPaste size={14} /> Pegar Datos
              </button>
              <button
                onClick={() => {
                  const levelNameExport =
                    niveles.find((n) => n.id === activeLevelId)?.nombre ||
                    "TODOS LOS NIVELES";
                  const renderExVal = (v) => {
                    const num = parseFloat(v);
                    if (isNaN(num) || num === 0) return "-";
                    return num.toFixed(2);
                  };

                  let html = `<table style="border-collapse: collapse; font-family: Arial, sans-serif;">
                      <colgroup>
                         <col width="30" style="width: 30px;" />
                         <col width="50" style="width: 50px;" />
                         <col width="100" style="width: 100px;" />
                         <col width="100" style="width: 100px;" />
                         <col width="80" style="width: 80px;" />
                         <col width="80" style="width: 80px;" />
                         <col width="80" style="width: 80px;" />
                         <col width="90" style="width: 90px;" />
                         <col width="90" style="width: 90px;" />
                         <col width="90" style="width: 90px;" />
                         <col width="145" style="width: 145px;" />
                         <col width="145" style="width: 145px;" />
                         <col width="145" style="width: 145px;" />
                         <col width="145" style="width: 145px;" />
                         <col width="145" style="width: 145px;" />
                      </colgroup>
                      <thead>
                      <tr style="height: 15pt;">
                         <td colspan="15" style="border: none; background-color: #ffffff;"></td>
                      </tr>
                      <tr style="height: 20pt;">
                         <td style="border: none; background-color: #ffffff;"></td>
                         <th colspan="14" style="background-color: #333f8f; color: #ffffff; font-size: 12pt; font-family: Arial, sans-serif; font-weight: bold; text-align: center; vertical-align: middle; border: 1px solid #333f8f; text-transform: uppercase;">
                            MATRIZ DE MUROS - NIVEL: ${levelNameExport}
                         </th>
                      </tr>
                      <tr style="height: 15pt;">
                         <td colspan="15" style="border: none; background-color: #ffffff;"></td>
                      </tr>
                      <tr style="background-color: #e6e6fa; color: #312eb5; font-size: 8pt; font-family: Arial, sans-serif; text-align: center; height: 15pt;">
                         <td style="border: none; background-color: #ffffff;"></td>
                         <th rowspan="2" style="border: 1px solid #c7d2fe; vertical-align: middle;">NO.</th>
                         <th rowspan="2" style="border: 1px solid #c7d2fe; vertical-align: middle;">EJE</th>
                         <th rowspan="2" style="border: 1px solid #c7d2fe; vertical-align: middle;">CLAVE</th>
                         <th rowspan="2" style="border: 1px solid #c7d2fe; vertical-align: middle;">LARGO (M)</th>
                         <th rowspan="2" style="border: 1px solid #c7d2fe; vertical-align: middle;">ANCHO (M)</th>
                         <th rowspan="2" style="border: 1px solid #c7d2fe; vertical-align: middle;">ALTO (M)</th>
                         <th rowspan="2" style="border: 1px solid #c7d2fe; vertical-align: middle;">BRUTO (M2)</th>
                         <th rowspan="2" style="border: 1px solid #c7d2fe; vertical-align: middle;">-HUECOS (M2)</th>
                         <th rowspan="2" style="border: 1px solid #c7d2fe; vertical-align: middle;">-CAST. (M2)</th>
                         <th rowspan="2" style="border: 1px solid #c7d2fe; vertical-align: middle;">NETA (M2)</th>
                         <th colspan="2" style="border: 1px solid #c7d2fe; vertical-align: middle;">APLANADO</th>
                         <th colspan="2" style="border: 1px solid #c7d2fe; vertical-align: middle;">RECUBRIMIENTO</th>
                      </tr>
                      <tr style="background-color: #e6e6fa; color: #312eb5; font-size: 8pt; font-family: Arial, sans-serif; height: 15pt; text-align: center;">
                         <td style="border: none; background-color: #ffffff;"></td>
                         <th style="border: 1px solid #c7d2fe; vertical-align: middle;">C1</th>
                         <th style="border: 1px solid #c7d2fe; vertical-align: middle;">C2</th>
                         <th style="border: 1px solid #c7d2fe; vertical-align: middle;">C1</th>
                         <th style="border: 1px solid #c7d2fe; vertical-align: middle;">C2</th>
                      </tr>
                   </thead><tbody>`;

                  muros.forEach((m, i) => {
                    const bruto = calcWallArea(m.largo, m.alto),
                      dedH = getWallHuecosTotal(m.huecos),
                      dedC = getWallCastillosTotal(m.castillos, m.alto),
                      neto = Math.max(0, bruto - dedH - dedC);
                    html += `<tr style="text-align: center; background-color: #ffffff; height: 18pt;">
                         <td style="border: none; background-color: #ffffff;"></td>
                         <td style="border: 1px solid #c7d2fe; color: #a6a6a6; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; text-align: center; mso-number-format:'0';">${i + 1}</td>
                         <td style="border: 1px solid #c7d2fe; color: #312eb5; font-weight: bold; text-transform: uppercase; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; text-align: center;">${m.eje || ""}</td>
                         <td style="border: 1px solid #c7d2fe; color: #000000; font-weight: bold; text-transform: uppercase; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; text-align: center;">${m.clave || ""}</td>
                         <td style="border: 1px solid #c7d2fe; color: #333f8f; font-weight: bold; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; mso-number-format:'0\.00'; text-align: center;">${renderExVal(m.largo)}</td>
                         <td style="border: 1px solid #c7d2fe; color: #333f8f; font-weight: bold; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; mso-number-format:'0\.00'; text-align: center;">${renderExVal(m.ancho)}</td>
                         <td style="border: 1px solid #c7d2fe; color: #333f8f; font-weight: bold; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; mso-number-format:'0\.00'; text-align: center;">${renderExVal(m.alto)}</td>
                         <td style="border: 1px solid #c7d2fe; color: #a6a6a6; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; mso-number-format:'0\.00'; text-align: center;">${renderExVal(bruto)}</td>
                         <td style="border: 1px solid #c7d2fe; color: #c00000; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; mso-number-format:'0\.00'; text-align: center;">${dedH > 0 ? dedH.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #c7d2fe; color: #c00000; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; mso-number-format:'0\.00'; text-align: center;">${dedC > 0 ? dedC.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #c7d2fe; color: #000000; font-weight: bold; font-size: 12pt; font-family: Arial, sans-serif; vertical-align: middle; mso-number-format:'0\.00'; text-align: center;">${renderExVal(neto)}</td>
                         <td style="border: 1px solid #c7d2fe; color: #a6a6a6; text-transform: uppercase; font-size: 10pt; font-family: Arial, sans-serif; vertical-align: middle; text-align: center;">${m.tipoAplanadoC1 || ""}</td>
                         <td style="border: 1px solid #c7d2fe; color: #a6a6a6; text-transform: uppercase; font-size: 10pt; font-family: Arial, sans-serif; vertical-align: middle; text-align: center;">${m.tipoAplanadoC2 || ""}</td>
                         <td style="border: 1px solid #c7d2fe; color: #a6a6a6; text-transform: uppercase; font-size: 10pt; font-family: Arial, sans-serif; vertical-align: middle; text-align: center;">${m.tipoRecubrimientoC1 || ""}</td>
                         <td style="border: 1px solid #c7d2fe; color: #a6a6a6; text-transform: uppercase; font-size: 10pt; font-family: Arial, sans-serif; vertical-align: middle; text-align: center;">${m.tipoRecubrimientoC2 || ""}</td>
                      </tr>`;
                  });
                  html += `</tbody></table><br/><br/><br/>`;

                  html += `<table style="width: 100%; border: none; font-family: Arial, sans-serif;"><tr>
                      <td style="width: 40px; border: none;"></td>
                      <td style="vertical-align: top; border: none; padding-right: 20px;">
                         <table style="border-collapse: collapse; text-align: center; width: 250px;">
                            <thead><tr style="height: 20pt;"><th colspan="2" style="background-color: #00b050; color: white; font-size: 12pt; font-family: Arial, sans-serif; padding: 8px; border: 1px solid #00b050; vertical-align: middle; text-transform: uppercase;">ÁREA NETA POR ESPESOR</th></tr></thead>
                            <tbody>`;
                  if (Object.keys(wallSummary.murosPorAncho).length === 0)
                    html += `<tr style="height: 18pt;"><td colspan="2" style="text-align: center; color: #94a3b8; font-size: 8pt; font-family: Arial, sans-serif; border: 1px solid #e2e8f0; vertical-align: middle;">Sin datos</td></tr>`;
                  Object.entries(wallSummary.murosPorAncho).forEach(
                    ([ancho, area]) => {
                      html += `<tr style="text-align: center; background-color: #ffffff; height: 18pt;">
                                  <td style="font-weight: bold; color: #000000; font-size: 12pt; font-family: Arial, sans-serif; border: none; border-bottom: 1px solid #e2e8f0; border-left: 1px solid #e2e8f0; vertical-align: middle; text-align: center;">Muro ${parseFloat(ancho).toFixed(2)}m</td>
                                  <td style="font-weight: bold; color: #000000; font-size: 12pt; font-family: Arial, sans-serif; border: none; border-bottom: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0; vertical-align: middle; mso-number-format:'0\.00'; text-align: center;">${area.toFixed(2)}</td>
                               </tr>`;
                    },
                  );
                  html += `</tbody></table></td>
                      <td style="vertical-align: top; border: none; padding-right: 20px;">
                         <table style="border-collapse: collapse; text-align: center; width: 250px;">
                            <thead><tr style="height: 20pt;"><th colspan="2" style="background-color: #7030a0; color: white; font-size: 12pt; font-family: Arial, sans-serif; padding: 8px; border: 1px solid #7030a0; vertical-align: middle; text-transform: uppercase;">DESGLOSE APLANADOS</th></tr></thead>
                            <tbody>`;
                  if (Object.keys(wallSummary.aplanados).length === 0)
                    html += `<tr style="height: 18pt;"><td colspan="2" style="text-align: center; color: #94a3b8; font-size: 8pt; font-family: Arial, sans-serif; border: 1px solid #e2e8f0; vertical-align: middle;">Sin datos</td></tr>`;
                  Object.entries(wallSummary.aplanados).forEach(
                    ([tipo, area]) => {
                      html += `<tr style="text-align: center; background-color: #ffffff; height: 18pt;">
                                  <td style="font-weight: bold; color: #000000; text-transform: uppercase; font-size: 12pt; font-family: Arial, sans-serif; border: none; border-bottom: 1px solid #e2e8f0; border-left: 1px solid #e2e8f0; vertical-align: middle; text-align: center;">${tipo}</td>
                                  <td style="font-weight: bold; color: #000000; font-size: 12pt; font-family: Arial, sans-serif; border: none; border-bottom: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0; vertical-align: middle; mso-number-format:'0\.00'; text-align: center;">${area.toFixed(2)}</td>
                               </tr>`;
                    },
                  );
                  html += `</tbody></table></td>
                      <td style="vertical-align: top; border: none;">
                         <table style="border-collapse: collapse; text-align: center; width: 250px;">
                            <thead><tr style="height: 20pt;"><th colspan="2" style="background-color: #000000; color: white; font-size: 12pt; font-family: Arial, sans-serif; padding: 8px; border: 1px solid #000000; vertical-align: middle; text-transform: uppercase;">DESGLOSE RECUBRIMIENTOS</th></tr></thead>
                            <tbody>`;
                  if (Object.keys(wallSummary.recubrimientos).length === 0)
                    html += `<tr style="height: 18pt;"><td colspan="2" style="text-align: center; color: #94a3b8; font-size: 8pt; font-family: Arial, sans-serif; border: 1px solid #e2e8f0; vertical-align: middle;">Sin datos</td></tr>`;
                  Object.entries(wallSummary.recubrimientos).forEach(
                    ([tipo, area]) => {
                      html += `<tr style="text-align: center; background-color: #ffffff; height: 18pt;">
                                  <td style="font-weight: bold; color: #000000; text-transform: uppercase; font-size: 12pt; font-family: Arial, sans-serif; border: none; border-bottom: 1px solid #e2e8f0; border-left: 1px solid #e2e8f0; vertical-align: middle; text-align: center;">${tipo}</td>
                                  <td style="font-weight: bold; color: #000000; font-size: 12pt; font-family: Arial, sans-serif; border: none; border-bottom: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0; vertical-align: middle; mso-number-format:'0\.00'; text-align: center;">${area.toFixed(2)}</td>
                               </tr>`;
                    },
                  );
                  html += `</tbody></table></td></tr></table>`;

                  exportFormattedExcel(html, `Matriz_Muros_${levelNameExport}`);
                }}
                className="px-6 py-2 bg-emerald-600 text-white font-black rounded-lg hover:bg-emerald-500 text-[10px] uppercase tracking-wider flex items-center gap-2 ml-auto"
              >
                <FileDown size={14} /> Exportar Excel
              </button>
            </div>
            <div className="bg-slate-50 p-6 border-t border-slate-200 shrink-0">
              <h3 className="text-sm font-black uppercase tracking-widest text-slate-800 mb-6 flex items-center gap-2">
                <LayoutDashboard size={18} className="text-indigo-600" />{" "}
                Resumen para Generadores
              </h3>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-100">
                  <span className="text-[10px] font-black text-emerald-600 uppercase tracking-widest block mb-3 border-b border-slate-100 pb-2">
                    Área Neta por Espesor
                  </span>
                  {Object.keys(wallSummary.murosPorAncho).length === 0 && (
                    <span className="text-xs font-bold text-slate-300">
                      Sin capturas
                    </span>
                  )}
                  {Object.entries(wallSummary.murosPorAncho).map(
                    ([ancho, area]) => (
                      <div
                        key={ancho}
                        className="flex justify-between items-center border-b border-slate-50 last:border-0 py-2"
                      >
                        <span className="truncate w-2/3 uppercase text-slate-700 text-2xl font-black">
                          Muro {ancho}m
                        </span>
                        <div className="flex items-center gap-2">
                          <span className="text-emerald-700 font-black text-xl">
                            {area.toFixed(2)} m2
                          </span>
                          <button
                            onClick={() => handleCopyWalls("TOTAL_NETO", ancho)}
                            className="p-1.5 bg-slate-100 text-slate-500 rounded-lg hover:bg-emerald-100 hover:text-emerald-700 transition-colors shadow-sm ml-1"
                            title={`Copiar muros de ${ancho}m`}
                          >
                            <Clipboard size={14} />
                          </button>
                        </div>
                      </div>
                    ),
                  )}
                </div>
                <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-100">
                  <span className="text-[10px] font-black text-purple-600 uppercase tracking-widest block mb-3 border-b border-slate-100 pb-2">
                    Desglose Aplanados
                  </span>
                  {Object.keys(wallSummary.aplanados).length === 0 && (
                    <span className="text-xs font-bold text-slate-300">
                      Sin capturas
                    </span>
                  )}
                  {Object.entries(wallSummary.aplanados).map(([tipo, area]) => (
                    <div
                      key={tipo}
                      className="flex justify-between items-center border-b border-slate-50 last:border-0 py-2"
                    >
                      <span className="truncate w-2/3 uppercase text-slate-700 text-sm font-black">
                        {tipo}
                      </span>
                      <div className="flex items-center gap-2">
                        <span className="text-purple-700 font-black text-lg">
                          {area.toFixed(2)} m2
                        </span>
                        <button
                          onClick={() => handleCopyWalls("APLANADO", tipo)}
                          className="p-1.5 bg-slate-100 text-slate-500 rounded-lg hover:bg-purple-100 hover:text-purple-700 transition-colors shadow-sm ml-1"
                          title={`Copiar detalles de muros con ${tipo}`}
                        >
                          <Clipboard size={14} />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
                <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-100">
                  <span className="text-[10px] font-black text-fuchsia-600 uppercase tracking-widest block mb-3 border-b border-slate-100 pb-2">
                    Desglose Recubrimientos
                  </span>
                  {Object.keys(wallSummary.recubrimientos).length === 0 && (
                    <span className="text-xs font-bold text-slate-300">
                      Sin capturas
                    </span>
                  )}
                  {Object.entries(wallSummary.recubrimientos).map(
                    ([tipo, area]) => (
                      <div
                        key={tipo}
                        className="flex justify-between items-center border-b border-slate-50 last:border-0 py-2"
                      >
                        <span className="truncate w-2/3 uppercase text-slate-700 text-sm font-black">
                          {tipo}
                        </span>
                        <div className="flex items-center gap-2">
                          <span className="text-fuchsia-700 font-black text-lg">
                            {area.toFixed(2)} m2
                          </span>
                          <button
                            onClick={() =>
                              handleCopyWalls("RECUBRIMIENTO", tipo)
                            }
                            className="p-1.5 bg-slate-100 text-slate-500 rounded-lg hover:bg-fuchsia-100 hover:text-fuchsia-700 transition-colors shadow-sm ml-1"
                            title={`Copiar detalles de muros con ${tipo}`}
                          >
                            <Clipboard size={14} />
                          </button>
                        </div>
                      </div>
                    ),
                  )}
                </div>
              </div>
            </div>
          </div>
        ) : showSettings ? (
          <div className="bg-white rounded-3xl p-10 shadow-xl border border-gray-100 animate-in fade-in zoom-in-95 duration-300">
            <div className="flex items-center justify-between mb-10">
              <h2 className="text-2xl font-black underline underline-offset-8 decoration-blue-500 decoration-4">
                {activePartidaId ? "Ajustes de Partida" : "Configuración del Proyecto"}
              </h2>
              <button
                onClick={() => setShowSettings(false)}
                className="p-2 hover:bg-slate-100 rounded-full transition-colors"
              >
                <X size={24} />
              </button>
            </div>

            <div className="grid md:grid-cols-2 gap-12">
              <div className="space-y-4">
                {activePartidaId && (
                  <>
                    <label className="block text-[10px] font-black text-blue-400 mb-3 uppercase tracking-[0.2em]">
                      Gestión de Niveles de la Partida
                    </label>
                    {niveles.map((n, i) => (
                      <div
                        key={n.id}
                        className="flex gap-3 bg-slate-50 p-2 rounded-xl border items-center shadow-sm"
                      >
                        <DebouncedCell
                          className="flex-1 bg-transparent p-2 font-black text-slate-700"
                          value={n.nombre || ""}
                          onChange={(v) => {
                            const x = [...niveles];
                            x[i].nombre = v;
                            setNiveles(x);
                          }}
                        />
                        <button
                          onClick={() =>
                            setNiveles(niveles.filter((l) => l.id !== n.id))
                          }
                          className="text-red-300 hover:text-red-500 p-2"
                        >
                          <Trash2 size={16} />
                        </button>
                      </div>
                    ))}
                    <button
                      onClick={() =>
                        setNiveles([
                          ...niveles,
                          {
                            id: `n${Date.now()}`,
                            nombre: `N${niveles.length + 1}`,
                          },
                        ])
                      }
                      className="w-full py-3 mt-4 border-2 border-dashed rounded-xl font-black text-blue-500 border-blue-200 hover:bg-blue-50 transition-all"
                    >
                      + Agregar Nivel a Partida
                    </button>
                  </>
                )}
              </div>

              <div className="bg-blue-50/50 p-8 rounded-[2.5rem] border h-fit">
                {!activePartidaId && (
                  <div className="space-y-6">
                    <div>
                      <label className="block text-[10px] font-black text-blue-400 mb-2 uppercase tracking-[0.2em]">
                        Nombre del Proyecto (Obra)
                      </label>
                      <DebouncedCell
                        className="w-full p-4 bg-white border border-blue-100 rounded-xl text-lg font-black shadow-sm outline-none"
                        value={obraInfo.nombre || ""}
                        onChange={(v) => {
                          const newInfo = { ...obraInfo, nombre: v };
                          setObraInfo(newInfo);
                          onUpdateMetadata(
                            projectId,
                            newInfo.nombre,
                            newInfo.ubicacion,
                          );
                        }}
                      />
                    </div>
                    <div>
                      <label className="block text-[10px] font-black text-blue-400 mb-2 uppercase tracking-[0.2em]">
                        Ubicación
                      </label>
                      <DebouncedCell
                        className="w-full p-4 bg-white border border-blue-100 rounded-xl text-sm font-bold shadow-sm outline-none"
                        value={obraInfo.ubicacion || ""}
                        onChange={(v) => {
                          const newInfo = { ...obraInfo, ubicacion: v };
                          setObraInfo(newInfo);
                          onUpdateMetadata(
                            projectId,
                            newInfo.nombre,
                            newInfo.ubicacion,
                          );
                        }}
                      />
                    </div>
                  </div>
                )}
                {activePartidaId && (
                  <div className={!activePartidaId ? "mt-6 pt-6 border-t border-blue-200/50" : ""}>
                    <label className="block text-[10px] font-black text-blue-400 mb-3 uppercase tracking-[0.2em]">
                      Nombre Partida Actual
                    </label>
                    <DebouncedCell
                      className="w-full p-5 bg-white border border-blue-100 rounded-2xl text-2xl font-black shadow-sm outline-none text-slate-800"
                      value={activePartida?.nombre || ""}
                      onChange={(v) =>
                        setPartidas((prev) =>
                          prev.map((p) =>
                            p.id === activePartidaId
                              ? { ...p, nombre: (v || "").toUpperCase() }
                              : p,
                          ),
                        )
                      }
                    />
                  </div>
                )}
              </div>
            </div>
          </div>
        ) : activePartidaId ? (
          <div onPaste={handlePasteProjectConceptos} className="overflow-x-auto bg-white rounded-lg shadow border border-gray-200 flex-1">
            <table className="text-sm text-left border-collapse table-fixed w-full">
              <thead className="bg-slate-800 text-white uppercase tracking-wider select-none sticky top-0 z-30 font-black">
                <tr>
                  <th
                    className="px-1 py-2 border border-slate-700 text-center sticky left-0 bg-slate-800 z-30"
                    style={{ width: colWidths.select }}
                  >
                    <button
                      onClick={() =>
                        setSelectedIds(
                          selectedIds.length === conceptos.length &&
                            conceptos.length > 0
                            ? []
                            : conceptos.map((c) => c.id),
                        )
                      }
                    >
                      {conceptos.length > 0 &&
                      selectedIds.length === conceptos.length ? (
                        <CheckSquare size={14} />
                      ) : (
                        <Square size={14} />
                      )}
                    </button>
                  </th>
                  {[
                    { id: "id", label: "Código" },
                    { id: "clave", label: "Clave" },
                    { id: "cc", label: "CC" },
                    { id: "justificacion", label: "Justif." },
                    { id: "descripcion", label: "Descripción" },
                    { id: "unidad", label: "Unid." },
                  ].map((col) => (
                    <th
                      key={col.id}
                      className="px-1.5 py-2 border border-slate-700 relative group text-[9px] font-black leading-tight"
                      style={{ width: colWidths[col.id] }}
                    >
                      <div className="truncate w-full">{col.label}</div>
                      <div
                        onMouseDown={(e) => startResizing(col.id, e, "summary")}
                        className="absolute right-0 top-0 bottom-0 w-2 cursor-col-resize hover:bg-blue-400 z-40 transition-colors"
                      />
                    </th>
                  ))}
                  {niveles.map((n) => (
                    <th
                      key={n.id}
                      className="px-1 py-1.5 border border-slate-700 text-center bg-blue-900 relative group"
                      style={{ width: colWidths.niveles }}
                    >
                      <div className="flex flex-col items-center gap-0.5 w-full overflow-hidden">
                        <span className="text-[8px] md:text-[9px] font-black truncate max-w-full px-1">
                          {n.nombre}
                        </span>
                        <button
                          onClick={() => setActiveLevelReport(n)}
                          title="Ver Cédula de Cuantificación"
                          className="bg-blue-700 hover:bg-blue-600 p-0.5 rounded transition-colors"
                        >
                          <ListPlus size={12} />
                        </button>
                      </div>
                    </th>
                  ))}
                  <th
                    className="px-1 py-2 border border-slate-700 text-center bg-black relative"
                    style={{ width: colWidths.totales }}
                  >
                    <span className="text-[9px] font-black">TOTAL</span>
                  </th>
                  <th
                    className="px-1 py-2 border border-slate-700"
                    style={{ width: 30 }}
                  ></th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {conceptos.map((c, idx) => (
                  <tr
                    key={c.id}
                    className={idx % 2 === 0 ? "bg-white" : "bg-gray-50"}
                  >
                    <td className="p-0 border text-center sticky left-0 bg-inherit z-10">
                      <button
                        onClick={() => toggleSelectConcept(c.id)}
                        className={`${selectedIds.includes(c.id) ? "text-blue-600" : "text-slate-300"}`}
                      >
                        {selectedIds.includes(c.id) ? (
                          <CheckSquare size={14} />
                        ) : (
                          <Square size={14} />
                        )}
                      </button>
                    </td>
                    <td className="p-0 border font-bold text-blue-700 leading-tight">
                      <DebouncedCell
                        value={c.id}
                        onChange={(v) => updateConceptoId(c.id, v)}
                        onDoubleClick={() => {
                          setConceptSelectorFor(c.id);
                          setConceptSearchTerm("");
                        }}
                        title="Doble clic para buscador"
                        className="w-full h-full p-1.5 bg-transparent text-left outline-none text-blue-700 font-black min-w-0 cursor-pointer hover:bg-blue-50/50 text-[13px]"
                      />
                    </td>
                    <td className="p-0 border text-center text-[13px]">
                      <DebouncedCell
                        value={c.clave}
                        onChange={(v) => updateConceptoField(c.id, "clave", v)}
                        className="w-full h-full p-1.5 bg-transparent text-center outline-none min-w-0"
                      />
                    </td>
                    <td className="p-0 border text-center text-[13px]">
                      <DebouncedCell
                        value={c.cc}
                        onChange={(v) => updateConceptoField(c.id, "cc", v)}
                        className="w-full h-full p-1.5 bg-transparent text-center outline-none min-w-0"
                      />
                    </td>
                    <td className="p-0 border italic text-gray-400 text-[9px]">
                      <DebouncedCell
                        isTextArea
                        rows={2}
                        value={c.justificacion}
                        onChange={(v) =>
                          updateConceptoField(c.id, "justificacion", v)
                        }
                        className="w-full h-full p-1 bg-transparent outline-none resize-none min-w-0 leading-tight"
                      />
                    </td>
                    <td className="p-0 border text-[10px] md:text-[11px] font-bold">
                      <DebouncedCell
                        isTextArea
                        rows={4}
                        value={c.descripcion}
                        onChange={(v) =>
                          updateConceptoField(c.id, "descripcion", v)
                        }
                        className="w-full h-auto min-h-[60px] p-2 bg-transparent outline-none min-w-0 leading-tight uppercase"
                      />
                    </td>
                    <td className="p-0 border text-center font-black">
                      <DebouncedCell
                        value={c.unidad}
                        onChange={(v) => updateConceptoField(c.id, "unidad", v)}
                        className="w-full h-full p-1.5 bg-transparent text-center outline-none text-blue-700 uppercase min-w-0 text-[15px]"
                      />
                    </td>
                    {niveles.map((n) => {
                      const vol = getVolumenNivel(c.id, n.id);
                      return (
                        <td
                          key={n.id}
                          onClick={() => {
                            setEditingModal({ concepto: c, nivel: n });
                            setSelectedGeneratorRows([]);
                          }}
                          className={`p-1 md:p-1.5 border text-center cursor-pointer font-black text-base md:text-lg leading-tight ${vol > 0 ? "bg-emerald-50 text-emerald-800" : "text-slate-200 hover:bg-blue-50"}`}
                        >
                          {vol.toFixed(2)}
                        </td>
                      );
                    })}
                    <td className="p-1 md:p-1.5 border text-center font-black bg-gray-100 text-base md:text-lg leading-tight text-slate-800">
                      {getTotalConcepto(c.id).toFixed(2)}
                    </td>
                    <td className="p-0 border text-center">
                      <button
                        onClick={() =>
                          updateActiveConceptos((prev) =>
                            prev.filter((x) => x.id !== c.id),
                          )
                        }
                        className="text-red-300 hover:text-red-500 transition-colors p-1"
                      >
                        <Trash2 size={14} />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div className="p-3 md:p-4 flex gap-4 sticky bottom-0 bg-white border-t">
              <button
                onClick={() =>
                  updateActiveConceptos((prev) => [
                    ...prev,
                    {
                      id: `C-${Date.now().toString(36).toUpperCase()}`,
                      clave: "",
                      cc: "",
                      justificacion: "",
                      descripcion: "",
                      unidad: "m2",
                    },
                  ])
                }
                className="px-5 py-2 bg-slate-700 text-white rounded-xl text-[10px] md:text-xs font-black shadow-md transition-transform active:scale-95 uppercase tracking-wider"
              >
                + Agregar Concepto
              </button>
              <button
                onClick={async () => {
                  try {
                    const clipText = await navigator.clipboard.readText();
                    if (clipText) {
                       handlePasteProjectConceptos({ clipboardData: { getData: () => clipText }, preventDefault: () => {} });
                    }
                  } catch (e) {
                    alert("Por favor, selecciona la tabla y presiona Ctrl+V para pegar (permisos del portapapeles denegados).");
                  }
                }}
                className="px-5 py-2 bg-emerald-600 text-white rounded-xl text-[10px] md:text-xs font-black shadow-md transition-transform active:scale-95 uppercase tracking-wider flex items-center gap-2 hover:bg-emerald-700"
              >
                <Clipboard size={14} /> Pegar desde Portapapeles
              </button>
              <button
                onClick={() => {
                   exportRealExcelElegante(activePartida, conceptos, niveles, generadores);
                }}
                className="px-5 py-2 bg-emerald-600 text-white rounded-xl text-[10px] md:text-xs font-black shadow-md transition-transform active:scale-95 uppercase tracking-wider hover:bg-emerald-500 flex items-center gap-2 ml-auto"
              >
                <FileDown size={16} /> Exportar Excel Elegante
              </button>
            </div>
          </div>
        ) : (
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6 animate-in fade-in duration-300 pb-10">
            {partidas.map((p) => (
              <div
                key={p.id}
                onClick={() => openPartida(p.id)}
                className="bg-white rounded-[2rem] p-8 border border-slate-200 shadow-sm hover:shadow-xl hover:border-blue-300 transition-all group flex flex-col cursor-pointer relative overflow-hidden h-[250px]"
              >
                <div className="absolute top-0 right-0 p-8 opacity-5 group-hover:opacity-10 transition-opacity transform group-hover:scale-110 group-hover:-rotate-12 duration-500">
                  <Layers size={120} />
                </div>
                <div className="bg-blue-50 w-16 h-16 rounded-2xl flex items-center justify-center text-blue-600 mb-6 group-hover:scale-110 transition-transform shadow-sm">
                  <Layers size={28} />
                </div>
                <h3 className="text-xl font-black text-slate-800 mb-2 uppercase leading-tight z-10">
                  {p.nombre}
                </h3>
                <p className="text-slate-400 text-xs font-bold uppercase tracking-widest mb-10 z-10">
                  {p.conceptos.length} Conceptos configurados
                </p>
                <div className="mt-auto flex justify-between items-center border-t border-slate-100 pt-6 z-10">
                  <span className="text-blue-600 font-black text-[10px] uppercase tracking-widest flex items-center gap-2">
                    Abrir Matriz <ChevronRight size={16} />
                  </span>
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      handleDeletePartida(p.id);
                    }}
                    className="text-slate-300 hover:text-red-500 p-2 rounded-full hover:bg-red-50 transition-colors"
                  >
                    <Trash2 size={18} />
                  </button>
                </div>
              </div>
            ))}
            <div
              onClick={handleCreatePartida}
              className="bg-slate-50/50 rounded-[2rem] p-8 border-2 border-dashed border-slate-300 hover:border-blue-400 hover:bg-blue-50 transition-all flex flex-col items-center justify-center cursor-pointer text-slate-400 hover:text-blue-600 min-h-[250px] group"
            >
              <div className="bg-white w-16 h-16 rounded-2xl flex items-center justify-center mb-6 shadow-sm group-hover:scale-110 transition-transform">
                <FolderPlus size={28} />
              </div>
              <span className="font-black text-sm uppercase tracking-widest">
                Nueva Partida
              </span>
              <span className="text-[10px] font-bold mt-2 opacity-60 uppercase text-center">
                Crear matriz en blanco
              </span>
            </div>
          </div>
        )}
      </main>

      {partidaToDelete && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[9999] flex items-center justify-center p-4">
          <div className="bg-white rounded-[2rem] p-8 max-w-md w-full shadow-2xl border border-slate-100 animate-in zoom-in-95 duration-200">
            <div className="w-16 h-16 bg-red-50 rounded-2xl flex items-center justify-center text-red-500 mb-6 font-black animate-bounce">
              <Trash2 size={32} />
            </div>
            <h3 className="text-2xl font-black text-slate-800 mb-4 uppercase tracking-tighter">
              ¿Eliminar partida?
            </h3>
            <p className="text-slate-500 font-bold text-sm mb-8 leading-relaxed uppercase tracking-wider">
              Esta acción borrará permanentemente todos los conceptos y datos de
              esta partida. No se puede deshacer.
            </p>
            <div className="grid grid-cols-2 gap-4">
              <button
                onClick={() => setPartidaToDelete(null)}
                className="px-6 py-4 rounded-2xl font-black text-slate-400 hover:text-slate-600 hover:bg-slate-50 transition-all uppercase tracking-widest text-xs"
              >
                Cancelar
              </button>
              <button
                onClick={() => confirmDeletePartida(partidaToDelete)}
                className="px-6 py-4 bg-red-600 text-white rounded-2xl font-black shadow-lg shadow-red-200 hover:bg-red-700 transition-all uppercase tracking-widest text-xs active:scale-95"
              >
                Sí, Eliminar
              </button>
            </div>
          </div>
        </div>
      )}

      {structAdjModal && (
        <StructAdjustmentModal
          structAdjModal={structAdjModal}
          onClose={() => setStructAdjModal(null)}
          onSave={saveStructAdjustment}
        />
      )}
    </div>
  );
}

const Dashboard = ({
  projects,
  onCreate,
  onSelect,
  onDelete,
  onCatMaestro,
}) => {
  return (
    <div className="min-h-screen bg-[#fcfdfe] font-sans text-slate-900 p-8 flex flex-col items-center">
      <div className="max-w-7xl w-full animate-in fade-in duration-500">
        <div className="flex justify-between items-end mb-12">
          <div className="flex flex-col gap-4">
            <div className="flex items-center gap-4">
              <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center text-white font-black text-sm">
                X
              </div>
              <div className="h-6 w-px bg-slate-200"></div>
              <span className="text-blue-600 font-black uppercase tracking-[0.4em] text-[8px]">
                Asset Management
              </span>
            </div>
            <div>
              <h1 className="text-4xl font-black tracking-tighter mb-2">
                Gestor de Obras
              </h1>
              <p className="text-slate-400 font-bold uppercase tracking-widest text-xs">
                Administración Global de Proyectos
              </p>
            </div>
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          <div
            onClick={onCreate}
            className="bg-slate-50/50 rounded-[2rem] p-8 border-2 border-dashed border-slate-300 hover:border-blue-400 hover:bg-blue-50 transition-all flex flex-col items-center justify-center cursor-pointer text-slate-400 hover:text-blue-600 min-h-[250px] group"
          >
            <div className="bg-white w-16 h-16 rounded-2xl flex items-center justify-center mb-6 shadow-sm group-hover:scale-110 transition-transform">
              <FolderPlus size={28} />
            </div>
            <span className="font-black text-sm uppercase tracking-widest">
              Nueva Obra
            </span>
            <span className="text-[10px] font-bold mt-2 opacity-60 uppercase text-center">
              Crear espacio de trabajo
            </span>
          </div>
          <div
            onClick={onCatMaestro}
            className="bg-teal-50/50 rounded-[2rem] p-8 border-2 border-dashed border-teal-300 hover:border-teal-400 hover:bg-teal-50 transition-all flex flex-col items-center justify-center cursor-pointer text-teal-400 hover:text-teal-600 min-h-[250px] group"
          >
            <div className="bg-white w-16 h-16 rounded-2xl flex items-center justify-center mb-6 shadow-sm group-hover:scale-110 transition-transform">
              <ListPlus size={28} />
            </div>
            <span className="font-black text-sm uppercase tracking-widest text-teal-600">
              Catálogo Maestro
            </span>
            <span className="text-[10px] font-bold mt-2 opacity-60 uppercase text-center text-teal-500">
              Gestionar base de datos central
            </span>
          </div>
          {projects.map((p) => (
            <div
              key={p.id}
              onClick={() => onSelect(p.id)}
              className="bg-white rounded-[2rem] p-8 border border-slate-200 shadow-sm hover:shadow-xl hover:border-blue-300 transition-all group flex flex-col cursor-pointer relative overflow-hidden h-[250px]"
            >
              <div className="absolute top-0 right-0 p-8 opacity-5 group-hover:opacity-10 transition-opacity transform group-hover:scale-110 group-hover:-rotate-12 duration-500">
                <Building2 size={120} />
              </div>
              <div className="bg-blue-50 w-16 h-16 rounded-2xl flex items-center justify-center text-blue-600 mb-6 group-hover:scale-110 transition-transform shadow-sm">
                <Building2 size={28} />
              </div>
              <h3 className="text-xl font-black text-slate-800 mb-2 uppercase leading-tight z-10">
                {p.name || "Obra Sin Nombre"}
              </h3>
              <p className="text-slate-400 text-xs font-bold uppercase tracking-widest mb-10 z-10">
                {p.location || "Ubicación no definida"}
              </p>
              <div className="mt-auto flex justify-between items-center border-t border-slate-100 pt-6 z-10">
                <span className="text-blue-600 font-black text-[10px] uppercase tracking-widest flex items-center gap-2">
                  Abrir Obra <ChevronRight size={16} />
                </span>
                <button
                  type="button"
                  onClick={(e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    console.log("Click en borrar obra:", p.id);
                    onDelete(p.id);
                  }}
                  className="bg-slate-50 text-slate-400 hover:text-red-600 px-4 py-2 rounded-xl hover:bg-red-50 transition-all flex items-center justify-center gap-2 group/del border border-slate-100 hover:border-red-100 relative z-50 ml-auto"
                >
                  <Trash2
                    size={16}
                    className="group-hover/del:scale-110 transition-transform"
                  />
                  <span className="text-[10px] font-black uppercase tracking-widest">
                    Borrar
                  </span>
                </button>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

export default function App() {
  const [user, setUser] = useState(null);
  const [authInit, setAuthInit] = useState(false);

  useEffect(() => {
    if (!auth) {
      setAuthInit(true);
      return;
    }
    const initAuth = async () => {
      try {
        if (
          typeof __initial_auth_token !== "undefined" &&
          __initial_auth_token
        ) {
          await signInWithCustomToken(auth, __initial_auth_token);
        } else {
          await signInAnonymously(auth);
        }
      } catch (e) {
        console.error("Error de autenticación", e);
      }
    };
    initAuth();
    const unsubscribe = onAuthStateChanged(auth, (u) => {
      setUser(u);
      setAuthInit(true);
    });
    return () => unsubscribe();
  }, []);

  const [projects, setProjects, loadedProjects, savingProjects] =
    usePersistentState("xdifica_projects_list", [], user);
  const [
    catalogoConceptos,
    setCatalogoConceptos,
    loadedCatalogo,
    savingCatalogo,
  ] = usePersistentState("xdifica_global_catalogo_v3", [], user);
  const [
    catalogoPartidasGlobal,
    setCatalogoPartidasGlobal,
    loadedCatPartidas,
    savingCatPartidas,
  ] = usePersistentState("xdifica_global_catalogo_partidas_v4", [], user);
  const [activeProjectId, setActiveProjectId] = useState(null);
  const [showGlobalCatalogo, setShowGlobalCatalogo] = useState(false);

  useEffect(() => {
    const handleBeforeUnload = (e) => {
      if (savingProjects || savingCatalogo || savingCatPartidas) {
        e.preventDefault();
        e.returnValue = "";
      }
    };
    window.addEventListener("beforeunload", handleBeforeUnload);
    return () => window.removeEventListener("beforeunload", handleBeforeUnload);
  }, [savingProjects, savingCatalogo, savingCatPartidas]);

  useEffect(() => {
    if (!loadedProjects || !authInit || !user) return;
    const oldData = localStorage.getItem("xdifica_partidas");
    if (oldData && projects.length === 0) {
      const migId = `PROJ-${Date.now()}`;
      let projectName = "Proyecto Migrado Local";
      let projectLocation = "Automático";
      const oldInfo = localStorage.getItem("xdifica_obraInfo");
      if (oldInfo) {
        try {
          const parsed = JSON.parse(oldInfo);
          if (parsed.nombre) projectName = parsed.nombre;
          if (parsed.ubicacion) projectLocation = parsed.ubicacion;
        } catch (e) {}
      }
      setProjects([
        { id: migId, name: projectName, location: projectLocation },
      ]);
      const payload = { partidas: JSON.parse(oldData) };
      if (oldInfo) payload.obraInfo = JSON.parse(oldInfo);
      const oldNiveles = localStorage.getItem("xdifica_niveles");
      if (oldNiveles) payload.niveles = JSON.parse(oldNiveles);
      if (user && db) {
        const docRef = doc(
          db,
          "artifacts",
          appId,
          "users",
          user.uid,
          "projects",
          migId,
        );
        setDoc(docRef, payload, { merge: true });
      } else {
        localStorage.setItem(
          `xdifica_proj_${migId}_partidas`,
          JSON.stringify(payload.partidas),
        );
        if (payload.obraInfo)
          localStorage.setItem(
            `xdifica_proj_${migId}_obraInfo`,
            JSON.stringify(payload.obraInfo),
          );
        if (payload.niveles)
          localStorage.setItem(
            `xdifica_proj_${migId}_niveles`,
            JSON.stringify(payload.niveles),
          );
      }
      localStorage.removeItem("xdifica_partidas");
      localStorage.removeItem("xdifica_obraInfo");
      localStorage.removeItem("xdifica_niveles");
    }
  }, [projects.length, setProjects, loadedProjects, authInit, user]);

  const handleCreateProject = () => {
    const newId = `PROJ-${Date.now().toString(36).toUpperCase()}`;
    setProjects([
      { id: newId, name: "NUEVA OBRA", location: "Sin definir" },
      ...projects,
    ]);
    setActiveProjectId(newId);
  };

  const [projectToDelete, setProjectToDelete] = useState(null);

  const handleDeleteProject = useCallback(
    async (id) => {
      console.log("Eliminando obra:", id);
      setProjects((prev) => prev.filter((p) => p.id !== id));
      if (user && db) {
        const docRef = doc(
          db,
          "artifacts",
          appId,
          "users",
          user.uid,
          "projects",
          id,
        );
        await deleteDoc(docRef).catch((e) =>
          console.error("Error al borrar en Firestore:", e),
        );
      } else {
        ["obraInfo", "niveles", "partidas", "catalogoConceptos"].forEach(
          (key) => localStorage.removeItem(`xdifica_proj_${id}_${key}`),
        );
      }
      setProjectToDelete(null);
    },
    [user, db, setProjects],
  );

  const handleUpdateMetadata = (id, name, location) => {
    setProjects((prev) =>
      prev.map((p) => (p.id === id ? { ...p, name, location } : p)),
    );
  };

  if (!authInit) return <LoadingScreen text="Conectando con la nube..." />;
  if (!loadedProjects || !loadedCatalogo || !loadedCatPartidas)
    return <LoadingScreen text="Sincronizando obras..." />;

  if (activeProjectId) {
    return (
      <ProjectWorkspace
        projectId={activeProjectId}
        onBack={() => setActiveProjectId(null)}
        onUpdateMetadata={handleUpdateMetadata}
        user={user}
        catalogoConceptos={catalogoConceptos}
        setCatalogoConceptos={setCatalogoConceptos}
        isCatalogoLoaded={loadedCatalogo}
        isSavingGlobal={savingProjects || savingCatalogo}
      />
    );
  }

  return (
    <>
      <Dashboard
        projects={projects}
        onCreate={handleCreateProject}
        onSelect={setActiveProjectId}
        onDelete={setProjectToDelete}
        onCatMaestro={() => setShowGlobalCatalogo(true)}
      />

      {projectToDelete && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[9999] flex items-center justify-center p-4">
          <div className="bg-white rounded-[2rem] p-8 max-w-md w-full shadow-2xl border border-slate-100 animate-in zoom-in-95 duration-200">
            <div className="w-16 h-16 bg-red-50 rounded-2xl flex items-center justify-center text-red-500 mb-6 font-black animate-bounce">
              <AlertCircle size={32} />
            </div>
            <h3 className="text-2xl font-black text-slate-800 mb-4 uppercase tracking-tighter">
              ¿Eliminar obra definitiva?
            </h3>
            <p className="text-slate-500 font-bold text-sm mb-8 leading-relaxed uppercase tracking-wider">
              Esta acción borrará permanentemente todos los conceptos,
              generadores y datos asociados a esta obra. No se puede deshacer.
            </p>
            <div className="grid grid-cols-2 gap-4">
              <button
                onClick={() => setProjectToDelete(null)}
                className="px-6 py-4 rounded-2xl font-black text-slate-400 hover:text-slate-600 hover:bg-slate-50 transition-all uppercase tracking-widest text-xs"
              >
                Cancelar
              </button>
              <button
                onClick={() => handleDeleteProject(projectToDelete)}
                className="px-6 py-4 bg-red-600 text-white rounded-2xl font-black shadow-lg shadow-red-200 hover:bg-red-700 transition-all uppercase tracking-widest text-xs active:scale-95"
              >
                Sí, Eliminar
              </button>
            </div>
          </div>
        </div>
      )}
      {showGlobalCatalogo && (
        <GlobalCatalogoModal
          onClose={() => setShowGlobalCatalogo(false)}
          user={user}
          catalogoConceptos={catalogoConceptos}
          setCatalogoConceptos={setCatalogoConceptos}
          catalogoPartidasGlobal={catalogoPartidasGlobal}
          setCatalogoPartidasGlobal={setCatalogoPartidasGlobal}
        />
      )}
    </>
  );
}

function GlobalCatalogoModal({
  onClose,
  user,
  catalogoConceptos,
  setCatalogoConceptos,
  catalogoPartidasGlobal = [],
  setCatalogoPartidasGlobal = () => {},
}) {
  const [activeGlobalPartidaId, setActiveGlobalPartidaId] = useState(null);
  const [selectedIds, setSelectedIds] = useState([]);
  const [partidaToDelete, setPartidaToDelete] = useState(null);
  const [catColWidths, setCatColWidths] = usePersistentState(
    "xdifica_catColWidths_v2",
    {
      selection: 40,
      id: 100,
      clave: 80,
      cc: 60,
      justificacion: 200,
      descripcion: 300,
      unidad: 70,
      action: 50,
    },
    user,
  );

  const resizingRef = useRef(null);
  const startResizing = (colId, e) => {
    e.preventDefault();
    resizingRef.current = {
      colId,
      startX: e.pageX,
      startWidth: catColWidths[colId] || 100,
    };
    document.addEventListener("mousemove", handleMouseMove);
    document.addEventListener("mouseup", stopResizing);
  };
  const handleMouseMove = useCallback(
    (e) => {
      if (!resizingRef.current) return;
      const { colId, startX, startWidth } = resizingRef.current;
      const newWidth = Math.max(30, startWidth + (e.pageX - startX));
      setCatColWidths((p) => ({ ...p, [colId]: newWidth }));
    },
    [setCatColWidths],
  );
  const stopResizing = useCallback(() => {
    document.removeEventListener("mousemove", handleMouseMove);
    document.removeEventListener("mouseup", stopResizing);
    resizingRef.current = null;
  }, [handleMouseMove]);

  const updateCatalogoField = useCallback(
    (id, field, value) => {
      setCatalogoConceptos((prev) =>
        prev.map((c) => (c.id === id ? { ...c, [field]: value } : c)),
      );
    },
    [setCatalogoConceptos],
  );

  const updateCatalogoId = useCallback(
    (oldId, newCodigo) => {
      if (!newCodigo.trim()) return;
      setCatalogoConceptos((prev) => {
        const matched = prev.find((cat) => (cat.codigo || cat.id) === newCodigo && cat.id !== oldId);
        return prev.map((c) =>
          c.id === oldId
            ? {
                ...c,
                codigo: newCodigo,
                ...(matched
                  ? {
                      clave: matched.clave || "",
                      cc: matched.cc || "",
                      justificacion: matched.justificacion || "",
                      descripcion: matched.descripcion || "",
                      unidad: matched.unidad || "",
                    }
                  : {}),
              }
            : c
        );
      });
    },
    [setCatalogoConceptos],
  );

  const toggleSelectConcept = (id) => {
    setSelectedIds((prev) =>
      prev.includes(id) ? prev.filter((i) => i !== id) : [...prev, id],
    );
  };

  const handleCreatePartidaGlobal = () => {
    const newId = `CAT-PARTIDA-${Date.now().toString(36).toUpperCase()}`;
    setCatalogoPartidasGlobal(prev => [
      ...prev,
      {
        id: newId,
        nombre: "NUEVA PARTIDA",
      }
    ]);
  };

  const handleDeletePartidaGlobal = (pid) => {
    setPartidaToDelete(pid);
  };
  
  const confirmDeletePartida = () => {
    if (partidaToDelete) {
      setCatalogoPartidasGlobal(prev => prev.filter(p => p.id !== partidaToDelete));
      setCatalogoConceptos(prev => prev.map(c => c.partidaId === partidaToDelete ? { ...c, partidaId: null } : c));
      if (activeGlobalPartidaId === partidaToDelete) setActiveGlobalPartidaId(null);
      setPartidaToDelete(null);
    }
  };

  const handlePasteFromExcel = useCallback(
    (e) => {
      if (!activeGlobalPartidaId) return; // Only paste inside a partida
      const clipboardData = e.clipboardData;
      if (!clipboardData) return;
      
      const text = clipboardData.getData("text/plain");
      if (!text) return;
      
      const lines = text.split(/\r?\n/).filter(line => line.trim() !== "");
      if (lines.length > 0 && text.includes("\t")) {
        e.preventDefault();
        const newRows = lines.map(line => {
          const cols = line.split("\t");
          if (cols.length >= 6) {
            return {
              id: `NEW-${Date.now()}-${Math.floor(Math.random()*100000)}`,
              codigo: cols[0]?.trim() || "",
              partidaId: activeGlobalPartidaId,
              clave: cols[1]?.trim() || "",
              cc: cols[2]?.trim() || "",
              justificacion: cols[3]?.trim() || "",
              descripcion: cols[4]?.trim() || "",
              unidad: cols[5]?.trim() || "",
            };
          } else {
            return {
              id: `NEW-${Date.now()}-${Math.floor(Math.random()*100000)}`,
              codigo: cols[0]?.trim() || "",
              partidaId: activeGlobalPartidaId,
              clave: "",
              cc: "",
              justificacion: "",
              descripcion: cols[1]?.trim() || "",
              unidad: cols[2]?.trim() || "",
            };
          }
        });
        
        setCatalogoConceptos(prev => {
          const newState = [...prev];
          newRows.forEach((r) => {
              newState.push(r);
          });
          return newState;
        });
      }
    },
    [setCatalogoConceptos, activeGlobalPartidaId]
  );

  return (
    <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[500] flex items-center justify-center p-4 md:p-8 animate-in fade-in duration-200">
      <div 
        className="bg-white rounded-3xl shadow-xl w-full max-w-6xl flex flex-col h-[80vh] overflow-hidden"
        onPaste={handlePasteFromExcel}
      >
        <div className="bg-teal-800 text-white p-6 flex justify-between items-center shrink-0 border-b border-teal-900">
          <div className="flex items-center gap-4">
            <div className="p-3 bg-teal-700 rounded-xl shadow-inner">
              <ListPlus size={24} className="text-teal-100" />
            </div>
            <div>
              <h2 className="text-xl font-black uppercase tracking-wider leading-tight">
                Catálogo Maestro de Conceptos (Global)
              </h2>
              <p className="text-xs text-teal-200 font-bold uppercase tracking-widest mt-1">
                Base de datos central disponible en todos los proyectos
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-2 bg-teal-700 hover:bg-teal-600 rounded-full transition-colors shadow-sm"
          >
            <X size={24} />
          </button>
        </div>
        <div className="flex-1 overflow-hidden bg-slate-50 p-4 md:p-6 flex flex-col min-h-0">
          {activeGlobalPartidaId ? (
            <div className="flex flex-col flex-1 animate-in fade-in duration-300 min-h-0">
              <div className="flex items-center justify-between gap-3 mb-4 shrink-0">
                <div className="flex items-center gap-3">
                  <button
                    onClick={() => setActiveGlobalPartidaId(null)}
                    className="flex items-center gap-2 text-teal-600 hover:text-teal-800 font-bold text-xs uppercase tracking-widest transition-colors"
                  >
                    <ChevronLeft size={16} /> Volver a Partidas
                  </button>
                  <div className="h-4 w-px bg-slate-300"></div>
                  <h3 className="font-black text-slate-700 uppercase tracking-widest text-sm">
                    {catalogoPartidasGlobal.find(p => p.id === activeGlobalPartidaId)?.nombre || "NUEVA PARTIDA"}
                  </h3>
                </div>
                <div className="flex items-center gap-3">
                  {selectedIds.length > 0 && (
                    <>
                      <button
                        onClick={async () => {
                          const itemsToCopy = catalogoConceptos.filter(c => selectedIds.includes(c.id));
                          const tsvStr = itemsToCopy.map(c => 
                            `${c.codigo || c.id}\t${c.clave || ""}\t${c.cc || ""}\t${(c.justificacion || "").replace(/\n/g, " ")}\t${(c.descripcion || "").replace(/\n/g, " ")}\t${c.unidad || ""}`
                          ).join("\n");
                          try {
                            await navigator.clipboard.writeText(tsvStr);
                            // Visual feedback
                            const btn = document.getElementById("btn-copy-cat");
                            if (btn) {
                              const originalText = btn.innerHTML;
                              btn.innerHTML = `<span class="flex items-center gap-2"><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg> Copiados!</span>`;
                              setTimeout(() => btn.innerHTML = originalText, 2000);
                            }
                          } catch (e) {}
                        }}
                        id="btn-copy-cat"
                        className="px-4 py-2 bg-emerald-600 text-white rounded-lg text-[10px] font-black shadow-md transition-all active:scale-95 uppercase tracking-wider hover:bg-emerald-700 flex items-center gap-2"
                      >
                        <Copy size={14} /> Copiar ({selectedIds.length})
                      </button>
                      <button
                        onClick={() => {
                          const existingIds = new Set(catalogoConceptos.map(x => x.id));
                          const newItems = [];
                          catalogoConceptos.forEach((c) => {
                            if (selectedIds.includes(c.id)) {
                              let newId = `${c.id}-COPIA-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
                              while (existingIds.has(newId)) {
                                  newId = `${c.id}-COPIA-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
                              }
                              existingIds.add(newId);
                              newItems.push({
                                ...c,
                                id: newId,
                              });
                            }
                          });
                          setCatalogoConceptos((prev) => [...prev, ...newItems]);
                          setSelectedIds([]);
                        }}
                        className="px-4 py-2 bg-indigo-600 text-white rounded-lg text-[10px] font-black shadow-md transition-all active:scale-95 uppercase tracking-wider hover:bg-indigo-700 flex items-center gap-2"
                      >
                        <Copy size={14} /> Duplicar ({selectedIds.length})
                      </button>
                      <button
                        onClick={() => {
                          setCatalogoConceptos((prev) =>
                            prev.filter((c) => !selectedIds.includes(c.id)),
                          );
                          setSelectedIds([]);
                        }}
                        className="px-4 py-2 bg-red-600 text-white rounded-lg text-[10px] font-black shadow-md transition-all active:scale-95 uppercase tracking-wider hover:bg-red-700 flex items-center gap-2"
                      >
                        <Trash2 size={14} /> Eliminar ({selectedIds.length})
                      </button>
                    </>
                  )}
                  <button
                    onClick={() => {
                      const newId = `CAT-${Date.now().toString(36).toUpperCase()}`;
                      setCatalogoConceptos((prev) => [
                        ...prev,
                        {
                          id: newId,
                          partidaId: activeGlobalPartidaId,
                          clave: "",
                          cc: "",
                          justificacion: "",
                          descripcion: "",
                          unidad: "m2",
                        },
                      ]);
                    }}
                    className="px-4 py-2 bg-teal-700 text-white rounded-lg text-[10px] font-black shadow-md transition-transform active:scale-95 uppercase tracking-wider hover:bg-teal-600 flex items-center gap-2"
                  >
                    <Plus size={14} /> Agregar Concepto
                  </button>
                </div>
              </div>
              
              <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-auto flex-1 min-h-0">
                <table className="text-sm text-left border-collapse table-fixed w-full min-w-[900px]">
                  <thead className="bg-slate-800 text-white uppercase tracking-wider select-none font-black sticky top-0 z-30">
                    <tr>
                      <th
                        className="p-3 border-r border-slate-700 text-center"
                        style={{ width: catColWidths.selection || 40 }}
                      >
                        <input
                          type="checkbox"
                          className="w-4 h-4 rounded border-slate-400 text-teal-600 focus:ring-teal-500 cursor-pointer"
                          checked={
                            catalogoConceptos.filter(c => c.partidaId === activeGlobalPartidaId).length > 0 &&
                            selectedIds.length === catalogoConceptos.filter(c => c.partidaId === activeGlobalPartidaId).length
                          }
                          onChange={(e) => {
                            if (e.target.checked)
                              setSelectedIds(catalogoConceptos.filter(c => c.partidaId === activeGlobalPartidaId).map((c) => c.id));
                            else setSelectedIds([]);
                          }}
                        />
                      </th>
                      {[
                        { id: "id", label: "Código" },
                        { id: "clave", label: "Clave" },
                        { id: "cc", label: "CC" },
                        { id: "justificacion", label: "Justificación" },
                        { id: "descripcion", label: "Descripción del Concepto" },
                        { id: "unidad", label: "Unidad" },
                      ].map((col) => (
                        <th
                          key={col.id}
                          className="px-3 py-3 border-r border-slate-700 relative group text-[10px] font-black leading-tight"
                          style={{ width: catColWidths[col.id] }}
                        >
                          <div className="truncate w-full">{col.label}</div>
                          <div
                            onMouseDown={(e) => startResizing(col.id, e)}
                            className="absolute right-0 top-0 bottom-0 w-2 cursor-col-resize hover:bg-teal-400 z-40 transition-colors"
                          />
                        </th>
                      ))}
                      <th
                        className="px-3 py-3 text-center"
                        style={{ width: catColWidths.action || 50 }}
                      ></th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 text-xs">
                    {(() => {
                      const groupConcepts = catalogoConceptos.filter(c => c.partidaId === activeGlobalPartidaId);
                      if (groupConcepts.length === 0) {
                        return (
                          <tr>
                            <td colSpan={8} className="p-12 text-center text-slate-400 font-bold uppercase tracking-widest">
                              <div className="flex flex-col items-center justify-center gap-3">
                                <AlertCircle size={32} className="opacity-30" />
                                <span>No hay conceptos en esta partida.</span>
                              </div>
                            </td>
                          </tr>
                        );
                      }
                      return groupConcepts.map((c, idx) => (
                        <tr
                          key={c.id}
                          className={`transition-colors ${idx % 2 === 0 ? "bg-white" : "bg-slate-50/50"} hover:bg-teal-50/30 ${selectedIds.includes(c.id) ? "bg-teal-50/70" : ""}`}
                        >
                          <td className="p-0 border-r border-slate-200 text-center">
                            <input
                              type="checkbox"
                              className="w-4 h-4 rounded border-slate-300 text-teal-600 focus:ring-teal-500 cursor-pointer"
                              checked={selectedIds.includes(c.id)}
                              onChange={(e) => {
                                if (e.target.checked)
                                  setSelectedIds((p) => [...p, c.id]);
                                else
                                  setSelectedIds((p) =>
                                    p.filter((sid) => sid !== c.id),
                                  );
                              }}
                            />
                          </td>
                          <td className="p-0 border-r border-slate-200">
                            <DebouncedCell
                              value={c.codigo || c.id}
                              onChange={(v) => updateCatalogoId(c.id, v)}
                              className="w-full h-full p-3 bg-transparent text-left outline-none font-black text-teal-800 min-w-0 text-[13px]"
                              placeholder="Ej. PRE-01"
                            />
                          </td>
                          <td className="p-0 border-r border-slate-200">
                            <DebouncedCell
                              value={c.clave}
                              onChange={(v) => updateCatalogoField(c.id, "clave", v)}
                              className="w-full h-full p-3 bg-transparent text-center outline-none min-w-0 font-bold text-slate-600 text-[13px]"
                            />
                          </td>
                          <td className="p-0 border-r border-slate-200">
                            <DebouncedCell
                              value={c.cc}
                              onChange={(v) => updateCatalogoField(c.id, "cc", v)}
                              className="w-full h-full p-3 bg-transparent text-center outline-none min-w-0 font-bold text-slate-600 text-[13px]"
                            />
                          </td>
                          <td className="p-0 border-r border-slate-200">
                            <DebouncedCell
                              isTextArea
                              rows={2}
                              value={c.justificacion}
                              onChange={(v) =>
                                updateCatalogoField(c.id, "justificacion", v)
                              }
                              className="w-full h-full p-2 bg-transparent outline-none resize-none min-w-0 italic text-slate-500"
                            />
                          </td>
                          <td className="p-0 border-r border-slate-200">
                            <DebouncedCell
                              isTextArea
                              rows={4}
                              value={c.descripcion}
                              onChange={(v) =>
                                updateCatalogoField(c.id, "descripcion", v)
                              }
                              className="w-full h-auto min-h-[60px] p-2 bg-transparent outline-none min-w-0 text-slate-700 font-bold uppercase transition-all focus:bg-white"
                            />
                          </td>
                          <td className="p-0 border-r border-slate-200">
                            <DebouncedCell
                              value={c.unidad}
                              onChange={(v) => updateCatalogoField(c.id, "unidad", v)}
                              className="w-full h-full p-3 bg-transparent text-center outline-none font-black text-teal-700 uppercase min-w-0 text-[15px]"
                            />
                          </td>
                          <td className="p-0 text-center">
                            <button
                              onClick={() =>
                                setCatalogoConceptos((prev) =>
                                  prev.filter((x) => x.id !== c.id),
                                )
                              }
                              className="text-slate-300 hover:text-red-500 transition-colors p-2"
                            >
                              <Trash2 size={16} />
                            </button>
                          </td>
                        </tr>
                      ));
                    })()}
                  </tbody>
                </table>
              </div>
            </div>
          ) : (
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6 animate-in fade-in duration-300 w-full">
              {catalogoPartidasGlobal.map((p) => {
                const conceptsCount = catalogoConceptos.filter(c => c.partidaId === p.id).length;
                return (
                  <div
                    key={p.id}
                    onClick={() => setActiveGlobalPartidaId(p.id)}
                    className="bg-white rounded-3xl p-5 border border-slate-200 shadow-sm hover:shadow-xl hover:border-teal-300 transition-all group flex flex-col cursor-pointer relative overflow-hidden h-[160px]"
                  >
                    <div className="absolute top-0 right-0 p-4 opacity-5 group-hover:opacity-10 transition-opacity transform group-hover:scale-110 group-hover:-rotate-12 duration-500">
                      <Layers size={80} />
                    </div>
                    <div className="flex items-center gap-3 mb-4">
                      <div className="bg-teal-50 w-10 h-10 rounded-xl flex items-center justify-center text-teal-600 group-hover:scale-110 transition-transform shadow-sm">
                        <Layers size={20} />
                      </div>
                      <input
                        type="text"
                        value={p.nombre}
                        onClick={(e) => e.stopPropagation()}
                        onChange={(e) => {
                          const val = e.target.value;
                          setCatalogoPartidasGlobal(prev => prev.map(p2 => p2.id === p.id ? { ...p2, nombre: val } : p2));
                        }}
                        className="bg-transparent font-black text-slate-800 uppercase leading-tight outline-none w-full text-sm placeholder:text-slate-400 z-10"
                        placeholder="NOMBRE DE LA PARTIDA"
                      />
                    </div>

                    <p className="text-slate-400 text-[10px] font-bold uppercase tracking-widest mb-2 z-10">
                      {conceptsCount} Conceptos configurados
                    </p>
                    <div className="mt-auto flex justify-between items-center border-t border-slate-100 pt-3 z-10">
                      <span className="text-teal-600 font-black text-[10px] uppercase tracking-widest flex items-center gap-2">
                        Abrir Matriz <ChevronRight size={14} />
                      </span>
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          handleDeletePartidaGlobal(p.id);
                        }}
                        className="text-slate-300 hover:text-red-500 p-2 rounded-full hover:bg-red-50 transition-colors"
                      >
                        <Trash2 size={18} />
                      </button>
                    </div>
                  </div>
                );
              })}
              
              <div
                onClick={handleCreatePartidaGlobal}
                className="bg-slate-50/50 rounded-3xl p-5 border-2 border-dashed border-teal-300 hover:border-teal-400 hover:bg-teal-50 transition-all flex flex-col items-center justify-center cursor-pointer text-slate-400 hover:text-teal-600 h-[160px] group"
              >
                <div className="bg-white w-10 h-10 rounded-xl flex items-center justify-center mb-3 shadow-sm group-hover:scale-110 transition-transform">
                  <FolderPlus size={20} />
                </div>
                <span className="font-black text-[11px] uppercase tracking-widest text-center">
                  Nueva Partida
                </span>
              </div>
            </div>
          )}
        </div>
        {partidaToDelete && (
          <div className="absolute inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center">
            <div className="bg-white p-6 rounded-3xl shadow-xl w-full max-w-sm flex flex-col items-center">
              <AlertCircle size={40} className="text-red-500 mb-4" />
              <h4 className="font-black text-slate-800 text-lg uppercase mb-2 text-center">¿Eliminar Partida?</h4>
              <p className="text-sm text-slate-500 text-center mb-6">Esta acción eliminará la organización de esta partida en el catálogo maestro.</p>
              <div className="flex w-full gap-3">
                <button onClick={() => setPartidaToDelete(null)} className="flex-1 py-3 bg-slate-100 uppercase tracking-widest text-xs font-bold text-slate-700 rounded-xl hover:bg-slate-200 transition-colors">Cancelar</button>
                <button onClick={confirmDeletePartida} className="flex-1 py-3 bg-red-600 text-white uppercase tracking-widest text-xs font-black rounded-xl shadow-md hover:bg-red-700 transition-transform active:scale-95">Eliminar</button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function StructAdjustmentModal({ structAdjModal, onClose, onSave }) {
  const [tipoLosa, setTipoLosa] = useState(structAdjModal.tipoLosa);
  const [esquinaLosa, setEsquinaLosa] = useState(structAdjModal.esquinaLosa);
  const [espesorLosa, setEspesorLosa] = useState(structAdjModal.espesorLosa);
  const [espesorZapata, setEspesorZapata] = useState(structAdjModal.espesorZapata || "0");

  return (
    <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[2000] flex items-center justify-center p-4">
      <div className="bg-white rounded-3xl shadow-2xl w-full max-w-sm overflow-hidden animate-in zoom-in-95 duration-200">
        <div className="bg-blue-800 text-white p-5 flex justify-between items-center">
          <div className="flex items-center gap-3">
            <Layout size={20} className="text-blue-300" />
            <div className="flex flex-col">
              <h3 className="text-sm font-black uppercase tracking-widest leading-none">
                Ajuste de Elemento
              </h3>
              <p className="text-[9px] font-bold text-blue-300 uppercase mt-1">
                Descuento por Losa y Zapata
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1 hover:bg-white/20 rounded-full transition-colors"
          >
            <X size={20} />
          </button>
        </div>
        <div className="p-6 space-y-6 max-h-[70vh] overflow-y-auto">
          <div>
            <label className="block text-[10px] font-black text-slate-400 uppercase mb-3 tracking-widest">
              Contacto Superior (¿Losa arriba?)
            </label>
            <div className="grid grid-cols-2 gap-3">
              <button
                onClick={() => setTipoLosa("maciza")}
                className={`py-3 px-4 rounded-2xl text-[11px] font-black uppercase tracking-wider transition-all border-2 ${tipoLosa === "maciza" ? "border-blue-600 bg-blue-50 text-blue-700 shadow-md" : "border-slate-100 bg-slate-50 text-slate-400 hover:border-slate-200"}`}
              >
                Maciza / Losa
              </button>
              <button
                onClick={() => setTipoLosa("vigueta")}
                className={`py-3 px-4 rounded-2xl text-[11px] font-black uppercase tracking-wider transition-all border-2 ${tipoLosa === "vigueta" ? "border-blue-600 bg-blue-50 text-blue-700 shadow-md" : "border-slate-100 bg-slate-50 text-slate-400 hover:border-slate-200"}`}
              >
                Vigueta
              </button>
            </div>
          </div>

          <div className="flex items-center justify-between p-4 bg-slate-50 rounded-2xl border border-slate-100">
            <div>
              <span className="block text-[10px] font-black text-slate-500 uppercase tracking-widest">
                ¿Es Esquina?
              </span>
              <span className="text-[9px] text-slate-400 font-bold uppercase">
                Afecta 1 cara lateral (Cimbra)
              </span>
            </div>
            <button
              onClick={() => setEsquinaLosa(!esquinaLosa)}
              className={`w-12 h-7 rounded-full transition-all relative ${esquinaLosa ? "bg-blue-600 shadow-inner" : "bg-slate-200"}`}
            >
              <div
                className={`absolute top-1 w-5 h-5 rounded-full bg-white transition-all shadow-sm ${esquinaLosa ? "left-6" : "left-1"}`}
              />
            </button>
          </div>

          <div className="grid grid-cols-1 gap-4">
            <div>
              <label className="block text-[10px] font-black text-slate-400 uppercase mb-2 tracking-widest">
                Espesor de Losa (m)
              </label>
              <div className="relative">
                <input
                  type="number"
                  step="0.01"
                  value={espesorLosa}
                  onChange={(e) => setEspesorLosa(e.target.value)}
                  className="w-full p-3 bg-slate-50 rounded-xl border border-slate-100 font-black text-slate-700 outline-none focus:border-blue-500 focus:bg-white transition-all text-sm pl-10"
                />
                <Ruler
                  className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-300"
                  size={16}
                />
              </div>
            </div>

            {((structAdjModal.tipo || "").toLowerCase().trim() !== "trabe") && (
              <div>
                <label className="block text-[10px] font-black text-slate-400 uppercase mb-2 tracking-widest">
                  Espesor de Zapata (m)
                </label>
                <div className="relative">
                  <input
                    type="number"
                    step="0.01"
                    value={espesorZapata}
                    onChange={(e) => setEspesorZapata(e.target.value)}
                    className="w-full p-3 bg-slate-50 rounded-xl border border-slate-100 font-black text-slate-700 outline-none focus:border-blue-500 focus:bg-white transition-all text-sm pl-10"
                  />
                  <Layers
                    className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-300"
                    size={16}
                  />
                </div>
              </div>
            )}
          </div>
        </div>
        <div className="p-4 bg-slate-50 flex gap-3">
          <button
            onClick={onClose}
            className="flex-1 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest hover:text-slate-600 transition-colors"
          >
            Cancelar
          </button>
          <button
            onClick={() => onSave({ tipoLosa, esquinaLosa, espesorLosa, espesorZapata })}
            className="flex-1 py-4 bg-blue-700 text-white rounded-2xl text-[10px] font-black uppercase tracking-widest shadow-xl shadow-blue-200 active:scale-95 transition-all hover:bg-blue-800"
          >
            Aplicar Ajuste
          </button>
        </div>
      </div>
    </div>
  );
}
