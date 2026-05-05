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
} from "lucide-react";
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

const getSteelTypesForElement = (t) => {
  t = (t || "").toLowerCase();
  if (t === "muro" || t === "muro curvo") return ["Vertical", "Horizontal", "Refuerzo Adicional"];
  if (t === "losa" || t === "losa nervada" || t.includes("zapata"))
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
  const template = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><style> table { border-collapse: collapse; font-family: Arial, sans-serif; font-size: 10px; margin-bottom: 20px; } th, td { border: 1px solid #cbd5e1; padding: 6px; vertical-align: middle; } th { font-weight: bold; text-align: center; } </style></head><body>${tableHtml}</body></html>`;
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
    localStorage.setItem(key, JSON.stringify(state));
    if (!user || !db) return;
    setIsSaving(true);
    const timeout = setTimeout(() => {
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
    }, 1000);
    return () => clearTimeout(timeout);
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
    localStorage.setItem(localKey, JSON.stringify(state));
    if (!user || !db) return;
    setIsSaving(true);
    const timeout = setTimeout(() => {
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
    }, 1000);
    return () => clearTimeout(timeout);
  }, [state, isLoaded, projectId, baseKey, user]);

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
}) => {
  const formatValue = (val) => {
    if (val === null || val === undefined) return "";
    if (type === "number" && step === "0.01" && val !== "") {
      const parsed = parseFloat(val);
      return isNaN(parsed) ? "" : parsed.toFixed(2);
    }
    return val;
  };
  const [localValue, setLocalValue] = useState(formatValue(value));
  useEffect(() => {
    setLocalValue(formatValue(value));
  }, [value, type, step]);
  const handleBlur = () => {
    let finalValue = localValue;
    if (type === "number" && step === "0.01" && finalValue !== "") {
      const parsed = parseFloat(finalValue);
      if (!isNaN(parsed)) {
        finalValue = parsed.toFixed(2);
        setLocalValue(finalValue);
      }
    }
    if (finalValue !== value) onChange(finalValue);
  };
  if (isTextArea)
    return (
      <textarea
        className={className}
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
      type={type}
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
            step="0.01"
            className="w-full p-1 bg-transparent text-center outline-none text-[11px]"
            value={row.largo}
            onChange={(v) => updateRow(row.id, "largo", v)}
          />
        </td>
        <td className="px-2 py-1.5 border-r border-slate-300 text-center font-black">
          <DebouncedCell
            type="number"
            step="0.01"
            className="w-full p-1 bg-transparent text-center outline-none text-[11px]"
            value={row.ancho}
            onChange={(v) => updateRow(row.id, "ancho", v)}
          />
        </td>
        <td className="px-2 py-1.5 border-r border-slate-300 text-center font-black">
          <DebouncedCell
            type="number"
            step="0.01"
            className="w-full p-1 bg-transparent text-center outline-none text-[11px]"
            value={row.kg_ml}
            onChange={(v) => updateRow(row.id, "kg_ml", v)}
          />
        </td>
        <td className="px-2 py-1.5 border-r border-slate-300 text-center font-black">
          <DebouncedCell
            type="number"
            step="0.01"
            className="w-full p-1 bg-transparent text-center outline-none text-[11px]"
            value={row.alto}
            onChange={(v) => updateRow(row.id, "alto", v)}
          />
        </td>
        <td className="px-2 py-1.5 border-r border-slate-300 text-center font-black bg-slate-50/50 text-blue-800 text-[11px]">
          {(() => {
            const l = parseFloat(row.largo); const lAbs = isNaN(l) || l === 0 ? 0 : Math.abs(l);
            const a = parseFloat(row.ancho); const aAbs = isNaN(a) || a === 0 ? 0 : Math.abs(a);
            const h = parseFloat(row.alto);  const hAbs = isNaN(h) || h === 0 ? 0 : Math.abs(h);
            const k = parseFloat(row.kg_ml); const kAbs = isNaN(k) || k === 0 ? 0 : Math.abs(k);
            const isActuallyZero = lAbs === 0 && aAbs === 0 && hAbs === 0 && kAbs === 0;
            const vol = isActuallyZero ? 0 : ((lAbs > 0 ? lAbs : 1) * (aAbs > 0 ? aAbs : 1) * (hAbs > 0 ? hAbs : 1) * (kAbs > 0 ? kAbs : 1));
            return vol.toLocaleString(undefined, {
              minimumFractionDigits: 2,
              maximumFractionDigits: 2,
            });
          })()}
        </td>
        <td className="px-2 py-1.5 border-r border-slate-300 text-center font-black">
          <DebouncedCell
            type="number"
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
  const [niveles, setNiveles, l2, s2] = useProjectState(
    projectId,
    "niveles",
    [{ id: "n1", nombre: "N1 +1.00" }],
    user,
  );
  const [partidas, setPartidas, l3, s3] = useProjectState(
    projectId,
    "partidas",
    [],
    user,
  );
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
  
  const [activeWallSubmodal, setActiveWallSubmodal] = useState(null);
  const [activeSteelSubmodal, setActiveSteelSubmodal] = useState(null);
  const [activeCasetonSubmodal, setActiveCasetonSubmodal] = useState(null);
  const [conceptSelectorFor, setConceptSelectorFor] = useState(null);
  const [conceptSearchTerm, setConceptSearchTerm] = useState("");

  const [selectedIds, setSelectedIds] = useState([]);
  const [clipboardRows, setClipboardRows] = useState(null);
  const [copyStatus, setCopyStatus] = useState("");
  const [copiedStructureRow, setCopiedStructureRow] = useState(null);
  const [selectedGeneratorRows, setSelectedGeneratorRows] = useState([]);
  const [selectedStructureRows, setSelectedStructureRows] = useState([]);
  const [copiedStructureRowsList, setCopiedStructureRowsList] = useState(null);
  const [selectedWallRows, setSelectedWallRows] = useState([]);
  const [copiedWallRowsList, setCopiedWallRowsList] = useState(null);
  const [copiedWallRow, setCopiedWallRow] = useState(null);
  const [showGlobalSteelReport, setShowGlobalSteelReport] = useState(false);
  const [globalReportLevel, setGlobalReportLevel] = useState("todos");

  const isFullyLoaded =
    l1 && l2 && l3 && l4 && l5 && l6 && l7 && l8 && isCatalogoLoaded && l10;

  const isSavingAny = 
    s1 || s2 || s3 || s4 || s5 || s6 || s7 || s8 || s10 || isSavingGlobal;

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

  useEffect(() => {
    if (
      l2 &&
      niveles.length > 0 &&
      !niveles.find((n) => n.id === activeLevelId)
    ) {
      setActiveLevelId(niveles[0].id);
    }
  }, [niveles, activeLevelId, l2]);

  const activePartida = useMemo(
    () => partidas.find((p) => p.id === activePartidaId),
    [partidas, activePartidaId],
  );
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
  const updateActiveConceptos = useCallback(
    (upd) =>
      setPartidas((prev) =>
        prev.map((p) =>
          p.id === activePartidaId
            ? {
                ...p,
                conceptos: typeof upd === "function" ? upd(p.conceptos) : upd,
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

  const handleCreatePartida = () =>
    setPartidas((prev) => [
      ...prev,
      {
        id: `P-${Date.now().toString(36).toUpperCase()}`,
        nombre: "NUEVA PARTIDA",
        conceptos: [],
        generadores: {},
        muros: {},
        estructuras: {},
      },
    ]);
  const handleDeletePartida = (id) => {
    if (window.confirm("¿Eliminar partida?"))
      setPartidas((prev) => prev.filter((p) => p.id !== id));
  };
  const openPartida = (id) => {
    setActivePartidaId(id);
    setSelectedIds([]);
    setShowSettings(false);
    setShowWallGenerator(false);
    setShowStructureGenerator(false);
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
    return (
      Math.round(
        (l > 0 ? l : 1) * (a > 0 ? a : 1) * (h > 0 ? h : 1) * (k > 0 ? k : 1) * p * 100,
      ) / 100
    );
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
        const matched = catalogoConceptos.find((cat) => cat.id === newId);
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
        const reader = new FileReader();
        reader.onload = (event) => {
          if (!editingModal) return;
          updateActiveGeneradores((prev) => {
            const currentConcept = prev[editingModal.concepto.id] || {},
              currentLevel = currentConcept[editingModal.nivel.id] || {
                rows: [],
                image: null,
              };
            return {
              ...prev,
              [editingModal.concepto.id]: {
                ...currentConcept,
                [editingModal.nivel.id]: {
                  ...currentLevel,
                  image: event.target.result,
                },
              },
            };
          });
        };
        reader.readAsDataURL(items[i].getAsFile());
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
        
        const volPza = (L > 0 ? L : 1) * (A > 0 ? A : 1) * (H > 0 ? H : 1) * (K > 0 ? K : 1);
        const isActuallyZero = L === 0 && A === 0 && H === 0 && K === 0;
        const finalVolPza = isActuallyZero ? 0 : volPza;
        const volTotal = finalVolPza * pzas;

        const formatDim = (val) => {
          if (isNaN(parseFloat(val)) || parseFloat(val) === 0) return "";
          return parseFloat(val).toFixed(2);
        };
        const formatPzas = (val) => {
          if (isNaN(parseFloat(val)) || parseFloat(val) === 0) return "1.00";
          return parseFloat(val).toFixed(2);
        };
        const claveCompleta = [r.eje, r.claveLoc].filter(Boolean).join(" - ");
        return `${claveCompleta}\\t${formatDim(r.largo)}\\t${formatDim(r.ancho)}\\t${formatDim(r.kg_ml)}\\t${formatDim(r.alto)}\\t${finalVolPza.toFixed(2)}\\t${formatPzas(r.piezas)}\\t${volTotal.toFixed(2)}`;
      })
      .join("\\n");
    const textArea = document.createElement("textarea");
    textArea.value = tsvString;
    textArea.style.position = "fixed";
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    try {
      document.execCommand("copy");
    } catch (err) {}
    document.body.removeChild(textArea);
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
              piezas:
                clipData.piezas !== undefined ? clipData.piezas : row.piezas,
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
      const anchoStr = parseFloat(w.ancho || 0).toFixed(2);
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
    if (tipo === "nervadura" || tipo === "losa de vigueta") return 0;
    
    if (tipo === "columna circular" || tipo === "pilas" || tipo === "pila") {
      const diam = getEffectiveLargo(r);
      const h = parseFloat(r.alto) || 0;
      const p = parseFloat(r.piezas) || 1;
      const radio = diam / 2;
      return Math.round((Math.PI * radio * radio * h * p) * 100) / 100;
    }

    const volBruto =
      getEffectiveLargo(r) *
      (parseFloat(r.ancho) || 0) *
      (parseFloat(r.alto) || 0) *
      (parseFloat(r.piezas) || 1);
    if (tipo === "losa nervada") {
      const volCasetones = getCasetonesTotalVol(r.casetones || [], r.piezas);
      return Math.max(0, Math.round((volBruto - volCasetones) * 100) / 100);
    }
    return Math.round(volBruto * 100) / 100;
  };
  const calcCimbra = (r) => {
    const l = getEffectiveLargo(r),
      a = parseFloat(r.ancho) || 0,
      h = parseFloat(r.alto) || 0,
      p = parseFloat(r.piezas) || 1,
      t = (r.tipo || "").toLowerCase();
    let area = 0;
    if (t === "columna" || t === "dado") area = (l + a) * 2 * h * p;
    else if (t === "columna circular") area = Math.PI * l * h * p;
    else if (t === "trabe" || t === "contratrabe")
      area = (l * a + l * h * 2) * p;
    else if (t === "losa" || t === "losa nervada" || t === "losa de vigueta") area = l * a * p;
    else if (t === "muro" || t === "muro curvo") area = l * h * 2 * p;
    else if (t === "zapata aislada" || t === "zapata corrida")
      area = (l + a) * 2 * p;
    return Math.round(area * 100) / 100;
  };
  const calcCimbraFrontera = (r) => {
    const l = getEffectiveLargo(r),
      a = parseFloat(r.ancho) || 0,
      p = parseFloat(r.piezas) || 1,
      t = (r.tipo || "").toLowerCase();
    let front = 0;
    if (t === "losa" || t === "losa nervada" || t === "losa de vigueta") front = (l + a) * 2 * p;
    return Math.round(front * 100) / 100;
  };
  const calcAceroItem = (acero) => {
    const p = parseFloat(acero.piezas) || 0,
      l = parseFloat(acero.longitud) || 0,
      g = parseFloat(acero.ganchos) || 0,
      t = parseFloat(acero.traslapes) || 0;
    const mlPorPieza = l + g * 0.5 + t;
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
      totalCasetonesGeneral = 0;
    let totalCimbraFrontera = 0;
    const breakdown = {},
      aceroCimentacionDetalle = {},
      aceroEstructuraDetalle = {},
      aceroEstructuraGlobalDetalle = {};
    estructuras.forEach((e) => {
      const c = calcConcreto(e),
        cim = calcCimbra(e),
        cimFrontera = calcCimbraFrontera(e),
        aceroKg = calcAceroTotalKg(e.aceros, e.piezas);
      const tipo = (e.tipo || "Sin Seleccionar").toLowerCase();
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
        "trabe",
        "nervadura",
        "columna",
        "columna circular",
      ].includes(tipo);
      const isEstructuraGpo = [
        "losa",
        "losa nervada",
        "losa de vigueta",
        "trabe",
        "nervadura",
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
        else if (tipo === "columna" || tipo === "columna circular") totalConcretoColumnas += c;
        else if (tipo === "muro" || tipo === "muro curvo") totalConcretoMuros += c;
      }

      totalCimbra += cim;
      totalCimbraFrontera += cimFrontera;

      const isConcEstGroup = ["losa", "losa nervada", "losa de vigueta", "trabe"].includes(tipo);

      let groupOtros = tipo;
      if (tipo === "columna") groupOtros = "Columnas";
      else if (tipo === "columna circular") groupOtros = "Columnas Circulares";
      else if (tipo === "pilas" || tipo === "pila") groupOtros = "Pilas";
      else if (tipo === "muro" || tipo === "muro curvo") groupOtros = "Muros";
      else if (tipo === "losa") groupOtros = "Losas";
      else if (tipo === "losa nervada") groupOtros = "Losas Nervadas";
      else if (tipo === "losa de vigueta") groupOtros = "Losas de Vigueta";
      else if (tipo === "trabe") groupOtros = "Trabes";
      else if (tipo === "nervadura") groupOtros = "Nervaduras";
      else if (tipo === "zapata aislada")
        groupOtros = `Zapatas Aisladas (Esp: ${parseFloat(e.alto || 0).toFixed(2)}m)`;
      else if (tipo === "zapata corrida")
        groupOtros = `Zapatas Corridas (Esp: ${parseFloat(e.alto || 0).toFixed(2)}m)`;
      else if (tipo === "contratrabe") groupOtros = "Contratrabes";
      else if (tipo === "dado") groupOtros = "Dados";

      const ensureBucket = (key, customTipo) => {
        if (!breakdown[key]) {
          breakdown[key] = {
            tipoOriginal: customTipo || tipo,
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
        const volCasetones = getCasetonesTotalVol(e.casetones || [], e.piezas);
        breakdown[groupOtros].casetones += volCasetones;
        totalCasetonesGeneral += volCasetones;
      }

      breakdown[groupOtros].cimbra += cim;
      breakdown[groupOtros].cimbraFrontera += cimFrontera;

      const p = parseFloat(e.piezas) || 1;
      const h = parseFloat(e.alto) || 0;
      const l = getEffectiveLargo(e);
      if (tipo === "muro" || tipo === "muro curvo") {
        breakdown[groupOtros].obturacionMuros += cim;
        breakdown[groupOtros].mallaRefuerzo += cim;
        breakdown[groupOtros].m2MuroTotal += l * h * p;
        
        if (tipo === "muro curvo") {
          if (h <= 3) {
            breakdown[groupOtros].cimbraMuroCurvo_0_3 += cim;
          } else if (h <= 6) {
            breakdown[groupOtros].cimbraMuroCurvo_0_3 += (cim / h) * 3;
            breakdown[groupOtros].cimbraMuroCurvo_3_6 += (cim / h) * (h - 3);
          } else {
            breakdown[groupOtros].cimbraMuroCurvo_0_3 += (cim / h) * 3;
            breakdown[groupOtros].cimbraMuroCurvo_3_6 += (cim / h) * 3;
            breakdown[groupOtros].cimbraMuroCurvo_6_9 += (cim / h) * (h - 6);
          }
        } else {
          if (h <= 3) {
            breakdown[groupOtros].cimbraMuros_0_3 += cim;
          } else if (h <= 6) {
            breakdown[groupOtros].cimbraMuros_0_3 += (cim / h) * 3;
            breakdown[groupOtros].cimbraMuros_3_6 += (cim / h) * (h - 3);
          } else {
            breakdown[groupOtros].cimbraMuros_0_3 += (cim / h) * 3;
            breakdown[groupOtros].cimbraMuros_3_6 += (cim / h) * 3;
            breakdown[groupOtros].cimbraMuros_6_9 += (cim / h) * (h - 6);
          }
        }
      } else if (tipo === "columna" || tipo === "columna circular") {
        breakdown[groupOtros].emplayerColumnas += cim;
        
        if (h <= 3) {
          breakdown[groupOtros].cimbraColumnas_0_3 += cim;
        } else if (h <= 6) {
          breakdown[groupOtros].cimbraColumnas_0_3 += (cim / h) * 3;
          breakdown[groupOtros].cimbraColumnas_3_6 += (cim / h) * (h - 3);
        } else {
          breakdown[groupOtros].cimbraColumnas_0_3 += (cim / h) * 3;
          breakdown[groupOtros].cimbraColumnas_3_6 += (cim / h) * 3;
          breakdown[groupOtros].cimbraColumnas_6_9 += (cim / h) * (h - 6);
        }
        
        if (h > 3) breakdown[groupOtros].andamiajeColumnas += p;
      } else if (tipo === "trabe") {
        breakdown[groupOtros].andamiajeTrabes += l * p;
      }

      if (["zapata aislada", "zapata corrida", "contratrabe"].includes(tipo)) {
        const a = parseFloat(e.ancho) || 0;
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
        const h = parseFloat(e.alto) || 0;
        const p = parseFloat(e.piezas) || 1;
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

      (e.aceros || []).forEach((a) => {
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
      });
    });

    Object.values(breakdown).forEach(group => {
      if (group.tipoOriginal === "muro" || group.tipoOriginal === "muro curvo") {
        group.pasosMuros = Math.ceil(group.m2MuroTotal / 50);
      }
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
    (oldId, newId) => {
      if (oldId === newId || !newId.trim()) return;
      setCatalogoConceptos((prev) =>
        prev.some((c) => c.id === newId)
          ? prev
          : prev.map((c) => (c.id === oldId ? { ...c, id: newId } : c)),
      );
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
        ["zapata aislada", "zapata corrida", "contratrabe", "dado", "pilas", "pila"].includes(
          (e.tipo || "Sin Seleccionar").toLowerCase(),
        ),
      );
    else if (tipoKey === "GLOBAL_ESTRUCTURA_ACERO")
      filtered = estructuras.filter((e) =>
        ["losa", "losa nervada", "losa de vigueta", "trabe", "nervadura", "columna", "columna circular"].includes(
          (e.tipo || "Sin Seleccionar").toLowerCase(),
        ),
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
        ["losa", "losa nervada", "losa de vigueta", "trabe", "nervadura"].includes(
          (e.tipo || "").toLowerCase(),
        ),
      );
    else if (tipoKey === "Columnas")
      filtered = estructuras.filter((e) =>
        ["columna", "columna circular"].includes(
          (e.tipo || "").toLowerCase(),
        ),
      );
    else if (tipoKey === "Acero en Estructuras")
      filtered = estructuras.filter((e) =>
        ["losa", "losa nervada", "losa de vigueta", "trabe", "nervadura", "columna", "columna circular"].includes(
          (e.tipo || "").toLowerCase(),
        ),
      );
    else if (tipoKey === "Cimentación")
      filtered = estructuras.filter((e) =>
        ["zapata aislada", "zapata corrida", "contratrabe", "dado", "pilas", "pila"].includes(
          (e.tipo || "").toLowerCase(),
        ),
      );
    else
      filtered = estructuras.filter((e) => {
        const t = (e.tipo || "Sin Seleccionar").toLowerCase();
        const tk = tipoKey.toLowerCase();

        let eGroup = t;
        if (t === "columna") eGroup = "columnas";
        else if (t === "columna circular") eGroup = "columnas circulares";
        else if (t === "muro" || t === "muro curvo") eGroup = "muros";
        else if (t === "losa") eGroup = "losas";
        else if (t === "losa nervada") eGroup = "losas nervadas";
        else if (t === "trabe") eGroup = "trabes";
        else if (t === "nervadura") eGroup = "nervaduras";
        else if (t === "zapata aislada")
          eGroup = `zapatas aisladas (esp: ${parseFloat(e.alto || 0).toFixed(2)}m)`;
        else if (t === "zapata corrida")
          eGroup = `zapatas corridas (esp: ${parseFloat(e.alto || 0).toFixed(2)}m)`;
        else if (t === "contratrabe") eGroup = "contratrabes";
        else if (t === "dado") eGroup = "dados";

        return eGroup === tk;
      });

    let rowsToCopy = [];
    if (material === "concreto") {
      rowsToCopy = filtered.flatMap((e) => {
        const base = { eje: e.eje || "", claveLoc: e.clave || "" };
        const pMain = parseFloat(e.piezas) || 1;
        let rows = [
          {
            ...base,
            largo: getEffectiveLargo(e),
            ancho: e.ancho,
            alto: e.alto,
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
      ["cimbra", "obturacionMuros", "mallaRefuerzo", "emplayerColumnas", "andamiajeColumnas", "andamiajeTrabes"].includes(material) ||
      material.startsWith("cimbraMuros_") ||
      material.startsWith("cimbraColumnas_") ||
      material.startsWith("cimbraMuroCurvo_")
    ) {
      rowsToCopy = filtered.flatMap((e) => {
        const l = getEffectiveLargo(e),
          a = e.ancho,
          h = parseFloat(e.alto) || 0,
          p = parseFloat(e.piezas) || 1,
          t = (e.tipo || "").toLowerCase(),
          base = { eje: e.eje || "", claveLoc: e.clave || "" };
        
        let shouldInclude = true;
        let outputAlto = h;

        if (material.startsWith("cimbraMuros_") || material.startsWith("cimbraColumnas_")) {
          if (material.endsWith("0_3")) {
            outputAlto = Math.min(3, h);
            shouldInclude = outputAlto > 0;
          } else if (material.endsWith("3_6")) {
            if (h <= 3) return [];
            outputAlto = Math.min(3, h - 3);
          } else if (material.endsWith("6_9")) {
            if (h <= 6) return [];
            outputAlto = h - 6;
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
              claveLoc: base.claveLoc ? base.claveLoc + labelSuffix : labelSuffix.trim(),
              largo: l,
              ancho: "",
              alto: outputAlto,
              piezas: p * 2,
            },
          ];
        if (t === "losa" || t === "losa nervada" || t === "losa de vigueta")
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
              claveLoc: base.claveLoc ? base.claveLoc + labelSuffix : labelSuffix.trim(),
              largo: Math.round(Math.PI * l * 100) / 100,
              ancho: "",
              alto: outputAlto,
              piezas: p,
            }
          ];
        }
        if (t === "columna" || t === "dado") {
          return [
            {
              ...base,
              claveLoc: base.claveLoc
                ? base.claveLoc + " (Lados X" + (labelSuffix ? labelSuffix : "") + ")"
                : "Lados X" + (labelSuffix ? labelSuffix : ""),
              largo: l,
              ancho: "",
              alto: outputAlto,
              piezas: p * 2,
            },
            {
              ...base,
              claveLoc: base.claveLoc
                ? base.claveLoc + " (Lados Y" + (labelSuffix ? labelSuffix : "") + ")"
                : "Lados Y" + (labelSuffix ? labelSuffix : ""),
              largo: a,
              ancho: "",
              alto: outputAlto,
              piezas: p * 2,
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
              ancho: "",
              alto: h,
              piezas: p * 2,
            },
            {
              ...base,
              claveLoc: base.claveLoc
                ? base.claveLoc + " (Lados Y" + labelSuffix + ")"
                : "Lados Y" + labelSuffix,
              largo: a,
              ancho: "",
              alto: h,
              piezas: p * 2,
            },
          ];
        if (t === "trabe" || t === "contratrabe")
          return [
            {
              ...base,
              claveLoc: base.claveLoc ? base.claveLoc + " (Fondo" + labelSuffix + ")" : "Fondo" + labelSuffix,
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
              ancho: "",
              alto: h,
              piezas: p * 2,
            },
          ];
        return [];
      });
    } else if (material === "pasosMuros") {
      const areaTotal = filtered.reduce((acc, e) => {
        if (((e.tipo || "").toLowerCase() === "muro" || (e.tipo || "").toLowerCase() === "muro curvo")) {
          const l = getEffectiveLargo(e);
          const h = parseFloat(e.alto) || 0;
          const p = parseFloat(e.piezas) || 1;
          return acc + (l * h * p);
        }
        return acc;
      }, 0);
      const pasos = Math.ceil(areaTotal / 50);
      rowsToCopy = [
        {
          eje: "",
          claveLoc: "Pasos en Muros",
          largo: 1,
          ancho: "",
          alto: "",
          piezas: pasos
        }
      ];
    } else if (material === "andamiajeColumnas") {
      rowsToCopy = filtered.flatMap((e) => {
        const h = parseFloat(e.alto) || 0;
        const p = parseFloat(e.piezas) || 1;
        const t = (e.tipo || "").toLowerCase();
        if (t === "columna" && h > 3) {
          return [{
            eje: e.eje || "",
            claveLoc: (e.clave || "") + " (Andamiaje)",
            largo: 1,
            ancho: "",
            alto: "",
            piezas: p
          }];
        }
        return [];
      });
    } else if (material === "andamiajeTrabes") {
      rowsToCopy = filtered.flatMap((e) => {
        const t = (e.tipo || "").toLowerCase();
        if (t === "trabe") {
          const p = parseFloat(e.piezas) || 1;
          const l = getEffectiveLargo(e);
          return [{
            eje: e.eje || "",
            claveLoc: (e.clave || "") + " (Andamiaje)",
            largo: l,
            ancho: "",
            alto: "",
            piezas: p
          }];
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
        if (t === "losa" || t === "losa nervada" || t === "losa de vigueta")
          return [
            {
              ...base,
              claveLoc: base.claveLoc
                ? base.claveLoc + " (Frontera X)"
                : "Frontera X",
              largo: l,
              ancho: "",
              alto: "",
              piezas: p * 2,
            },
            {
              ...base,
              claveLoc: base.claveLoc
                ? base.claveLoc + " (Frontera Y)"
                : "Frontera Y",
              largo: a,
              ancho: "",
              alto: "",
              piezas: p * 2,
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
        let pLargo = 0, pAncho = 0, eLargo = 0, eAncho = 0, eAlto = 0;
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
              largo: tipo === "pilas" || tipo === "pila" ? excavacionVol.toFixed(2) : eLargo,
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
      rowsToCopy = filtered.flatMap((e) =>
        (e.aceros || [])
          .filter((a) => numVar === "total" || a.numVarilla === numVar)
          .map((a) => ({
            eje: e.eje || "",
            claveLoc: (
              (e.clave ? e.clave + " " : "") +
              `${a.tipo || "Acero"} #${a.numVarilla || "-"}`
            ).trim(),
            largo: calcAceroItem(a).mlPorPieza,
            ancho: "",
            kg_ml: PESOS_VARILLA[a.numVarilla || "-"] || 0,
            alto: "",
            piezas: (parseFloat(a.piezas) || 1) * (parseFloat(e.piezas) || 1),
          })),
      );
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
              const updatedA = { ...a },
                tAcero = (updatedA.tipo || "").toLowerCase(),
                tEst = (newS.tipo || "").toLowerCase(),
                l = parseFloat(newS.largo) || 0,
                an = parseFloat(newS.ancho) || 0,
                al = parseFloat(newS.alto) || 0;
              if (tAcero === "longitudinal") {
                if (["columna", "columna circular", "dado", "pilas", "pila"].includes(tEst))
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
                updatedA.traslapes = "0";
                let perimetro = 0;
                if (["columna", "dado"].includes(tEst)) {
                  perimetro = (l + an) * 2;
                  if (
                    !updatedA.anchoCalc ||
                    parseFloat(updatedA.anchoCalc) === parseFloat(s.alto)
                  )
                    updatedA.anchoCalc = al > 0 ? al.toString() : "";
                } else if (tEst === "columna circular" || tEst === "pilas" || tEst === "pila") {
                  perimetro = l * Math.PI;
                  if (!updatedA.anchoCalc || parseFloat(updatedA.anchoCalc) === parseFloat(s.alto))
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
              if (field === "tipo") {
                const l = parseFloat(s.largo) || 0,
                  an = parseFloat(s.ancho) || 0,
                  al = parseFloat(s.alto) || 0;
                if (tipoAcero === "estribos" || tipoAcero === "zuncho") {
                  updatedA.ganchos = "0";
                  updatedA.traslapes = "0";
                  let perimetro = 0;
                  if (["columna", "dado"].includes(tipoEstructura)) {
                    perimetro = (l + an) * 2;
                    if (!updatedA.anchoCalc && al > 0)
                      updatedA.anchoCalc = al.toString();
                  } else if (tipoEstructura === "columna circular" || tipoEstructura === "pilas" || tipoEstructura === "pila") {
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
                  if (["columna", "columna circular", "pilas", "pila", "dado"].includes(tipoEstructura))
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
                  ].includes(tipoEstructura)
                ) {
                  updatedA.anchoCalc = l > 0 ? l.toString() : "";
                  updatedA.longitud = an > 0 ? an.toString() : "";
                } else if (tipoEstructura === "muro") {
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
        if (tipoAcero === "estribos" || tipoAcero === "zuncho") {
          if (["columna", "columna circular", "pilas", "pila", "dado"].includes(tipoEst))
            defVal = al > 0 ? al.toString() : "";
          else if (["trabe", "nervadura", "contratrabe"].includes(tipoEst))
            defVal = l > 0 ? l.toString() : "";
        } else if (tipoEst === "muro") {
          if (tipoAcero === "vertical") defVal = l > 0 ? l.toString() : "";
          else if (tipoAcero === "horizontal")
            defVal = al > 0 ? al.toString() : "";
        } else if (tipoAcero === "longitudinal") {
          if (["columna", "columna circular", "pilas", "pila", "dado"].includes(tipoEst))
            defVal = al > 0 ? al.toString() : "";
          else if (["trabe", "nervadura", "contratrabe"].includes(tipoEst))
            defVal = l > 0 ? l.toString() : "";
          else if (
            [
              "zapata aislada",
              "zapata corrida",
              "losa",
              "losa nervada",
            ].includes(tipoEst)
          )
            defVal = an > 0 ? an.toString() : "";
        } else if (
          tipoAcero === "transversal" &&
          ["zapata aislada", "zapata corrida", "losa", "losa nervada"].includes(
            tipoEst,
          )
        )
          defVal = l > 0 ? l.toString() : "";
        else if (tipoAcero === "bastones" || tipoAcero === "refuerzo adicional")
          defVal = l > 0 ? l.toString() : "";
      } else if (field === "longitud") {
        if (tipoAcero === "estribos" || tipoAcero === "zuncho") {
          let perimetro = 0;
          if (["columna", "dado"].includes(tipoEst)) perimetro = (l + an) * 2;
          else if (tipoEst === "columna circular" || tipoEst === "pilas" || tipoEst === "pila") perimetro = l * Math.PI;
          else if (["trabe", "nervadura", "contratrabe"].includes(tipoEst))
            perimetro = (an + al) * 2;
          if (perimetro > 0) defVal = (perimetro - 0.06 + 0.12).toFixed(2);
        } else if (tipoAcero === "longitudinal") {
          if (["columna", "columna circular", "pilas", "pila", "dado"].includes(tipoEst))
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
            ].includes(tipoEst)
          )
            defVal = l > 0 ? l.toString() : "";
        } else if (
          tipoAcero === "transversal" &&
          ["zapata aislada", "zapata corrida", "losa", "losa nervada"].includes(
            tipoEst,
          )
        )
          defVal = an > 0 ? an.toString() : "";
        else if (tipoEst === "muro") {
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
            ["columna", "columna circular", "pilas", "pila", "dado", "trabe", "nervadura", "contratrabe"].includes(
              tipoEst,
            ));
        if (!isPorPiezas) {
          const anchoC = parseFloat(a.anchoCalc),
            sep = parseFloat(a.separacion);
          if (!isNaN(anchoC) && anchoC > 0 && !isNaN(sep) && sep > 0)
            defVal = (Math.floor(anchoC / (sep / 100)) + 2).toString();
        }
      } else if (field === "traslapes") {
        const longCalc = parseFloat(a.longitud) || 0;
        defVal = longCalc >= 12 ? Math.floor(longCalc / 12).toString() : "0";
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
          if (["columna", "columna circular", "pilas", "pila", "dado"].includes(tipo) && parseFloat(s.alto) > 0)
            tramoSugerido = s.alto.toString();
          else if (
            ["trabe", "nervadura"].includes(tipo) &&
            parseFloat(s.largo) > 0
          )
            tramoSugerido = s.largo.toString();
          return {
            ...s,
            aceros: [
              ...(s.aceros || []),
              {
                id: `AC-${Date.now()}`,
                tipo: "",
                numVarilla: "4",
                anchoCalc: tramoSugerido,
                separacion: "",
                piezas: "1",
                longitud: "",
                ganchos: "",
                traslapes: "",
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
    const aceros = structure.aceros || [];
    const tiposAceroOpciones = getSteelTypesForElement(structure.tipo);
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
                        isPorPiezas = isVinculadoLargo || isLongitudinalManual;
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
                            {isLongitudinalManual ? (
                              <div className="text-center font-black text-slate-400 uppercase cursor-not-allowed">
                                N/A
                              </div>
                            ) : (
                              <DebouncedCell
                                type="number"
                                step="0.01"
                                value={item.anchoCalc}
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
                                className="w-full p-1 bg-transparent outline-none text-center font-black text-indigo-900"
                                placeholder="0.00"
                              />
                            )}
                          </td>
                          <td className="p-0 border-r bg-indigo-50/30">
                            {isPorPiezas ? (
                              <div className="text-center font-black text-slate-400 uppercase cursor-not-allowed">
                                N/A
                              </div>
                            ) : (
                              <DebouncedCell
                                type="number"
                                step="0.01"
                                value={item.separacion}
                                onChange={(v) =>
                                  updateSteelItem(
                                    structure.id,
                                    item.id,
                                    "separacion",
                                    v,
                                  )
                                }
                                className="w-full p-1 bg-transparent outline-none text-center font-black text-indigo-900"
                                placeholder="-"
                              />
                            )}
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
                            {isLongitudinalManual || isVinculadoLargo ? (
                              <div className="text-center font-black text-blue-800/60 bg-blue-50/50">
                                {item.longitud
                                  ? parseFloat(item.longitud).toFixed(2)
                                  : "0.00"}
                              </div>
                            ) : (
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
                                className="w-full p-1 bg-transparent outline-none text-center font-black text-blue-800"
                                placeholder="0.00"
                              />
                            )}
                          </td>
                          <td className="p-0 border-r bg-emerald-50/30">
                            <DebouncedCell
                              type="number"
                              step="1"
                              value={item.ganchos}
                              onChange={(v) =>
                                updateSteelItem(
                                  structure.id,
                                  item.id,
                                  "ganchos",
                                  v,
                                )
                              }
                              className="w-full p-1 bg-transparent outline-none text-center font-black text-emerald-800"
                              placeholder="0"
                            />
                          </td>
                          <td className="p-0 border-r bg-emerald-50/30">
                            <DebouncedCell
                              type="number"
                              step="0.01"
                              value={item.traslapes}
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
                              className="w-full p-1 bg-transparent outline-none text-center font-black text-emerald-800"
                              placeholder="0.00"
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
                              ({PESOS_VARILLA[num] || 0} kg/m)
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
                      ></th>
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
              <h4 className="font-black text-slate-600 uppercase text-xs mb-4 flex items-center gap-2">
                <FileDown size={16} /> Croquis / Referencia
              </h4>
              <div className="flex-1 bg-white border-2 border-dashed border-slate-300 rounded-xl flex items-center justify-center relative overflow-hidden group">
                {currentLevel.image ? (
                  <>
                    <img
                      src={currentLevel.image}
                      alt="Referencia"
                      className="w-full h-full object-contain p-2"
                    />
                    <div className="absolute inset-0 bg-black/50 opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-center gap-4">
                      <button
                        onClick={() =>
                          updateActiveGeneradores((prev) => {
                            const c = prev[concepto.id];
                            return {
                              ...prev,
                              [concepto.id]: {
                                ...c,
                                [nivel.id]: { ...c[nivel.id], image: null },
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
                      size={32}
                      className="mx-auto text-slate-300 mb-3"
                    />
                    <p className="text-[10px] font-bold text-slate-400 uppercase">
                      Haz Ctrl+V para pegar una imagen de referencia aquí
                    </p>
                  </div>
                )}
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
              const croquis = dataNivel.image;
              const totalVol = getVolumenNivel(c.id, n.id);

              return (
                <div
                  key={c.id}
                  className="mb-8 bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden break-inside-avoid"
                >
                  <div className="bg-slate-800 text-white p-3 flex justify-between items-center rounded-t-xl">
                    <div className="flex-1 flex items-center justify-between">
                      <h4 className="font-black text-[12px] tracking-widest w-1/3 text-left">
                        {c.clave || c.id}
                      </h4>
                      <p className="text-[15px] text-slate-300 font-bold uppercase text-center w-1/3 px-2">
                        {c.descripcion}
                      </p>
                      <div className="text-right w-1/3">
                        <span className="block text-[12px] font-black text-slate-400 uppercase">
                          Unidad
                        </span>
                        <span className="font-black text-blue-400 text-[18px]">
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
                            const l = parseFloat(r.largo); const lAbs = isNaN(l) || l === 0 ? 0 : Math.abs(l);
                            const a = parseFloat(r.ancho); const aAbs = isNaN(a) || a === 0 ? 0 : Math.abs(a);
                            const h = parseFloat(r.alto); const hAbs = isNaN(h) || h === 0 ? 0 : Math.abs(h);
                            const k = parseFloat(r.kg_ml); const kAbs = isNaN(k) || k === 0 ? 0 : Math.abs(k);
                            const isActuallyZero = lAbs === 0 && aAbs === 0 && hAbs === 0 && kAbs === 0;
                            const volPza = isActuallyZero ? 0 : ((lAbs > 0 ? lAbs : 1) * (aAbs > 0 ? aAbs : 1) * (hAbs > 0 ? hAbs : 1) * (kAbs > 0 ? kAbs : 1));
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
                                  ? parseFloat(r.kg_ml).toFixed(2)
                                  : "0.00"}
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
                    {croquis && (
                      <div className="w-full bg-slate-50 border-t border-slate-200 p-4 flex flex-col">
                        <span className="text-[12px] font-black uppercase tracking-widest text-slate-400 mb-2 text-center border-b border-slate-200 pb-1">
                          Croquis de Referencia
                        </span>
                        <div className="w-full flex items-center justify-center bg-white border-2 border-dashed border-slate-200 rounded-xl overflow-hidden p-2">
                          <img
                            src={croquis}
                            alt="Croquis"
                            className="w-full h-auto object-contain max-h-[300px]"
                          />
                        </div>
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
                            <td colspan="5" style="background-color: #000000; color: #ffffff; padding: 4px 8px; text-align: center; font-size: 10pt; font-weight: bold; border-top: 2px solid #000000; border-bottom: 1px solid #000000; border-right: 1px solid #ffffff;">
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
                      const l = parseFloat(r.largo); const lAbs = isNaN(l) || l === 0 ? 0 : Math.abs(l);
                      const a = parseFloat(r.ancho); const aAbs = isNaN(a) || a === 0 ? 0 : Math.abs(a);
                      const h = parseFloat(r.alto); const hAbs = isNaN(h) || h === 0 ? 0 : Math.abs(h);
                      const k = parseFloat(r.kg_ml); const kAbs = isNaN(k) || k === 0 ? 0 : Math.abs(k);
                      const isActuallyZero = lAbs === 0 && aAbs === 0 && hAbs === 0 && kAbs === 0;
                      const volPza = isActuallyZero ? 0 : ((lAbs > 0 ? lAbs : 1) * (aAbs > 0 ? aAbs : 1) * (hAbs > 0 ? hAbs : 1) * (kAbs > 0 ? kAbs : 1));
                      html += `<tr>
                            <td style="border: none;"></td>
                            <td style="padding: 4px; border: 1px solid #D9D9D9; border-left: 2px solid #000000; text-transform: uppercase; text-align: left; color: #000000; font-size: 10pt;">${r.eje} ${r.claveLoc}</td>
                            <td style="padding: 4px; border: 1px solid #D9D9D9; text-align: center; color: #000000; font-size: 10pt; mso-number-format:'0\\.00';">${renderExVal(r.largo)}</td>
                            <td style="padding: 4px; border: 1px solid #D9D9D9; text-align: center; color: #000000; font-size: 10pt; mso-number-format:'0\\.00';">${renderExVal(r.ancho)}</td>
                            <td style="padding: 4px; border: 1px solid #D9D9D9; text-align: center; color: #000000; font-size: 10pt; mso-number-format:'0\\.00';">${renderExVal(r.kg_ml)}</td>
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
                         </tr>
                         <tr><td colspan="9" style="height: 10px; border: none;"></td></tr>
                         <tr>
                            <td style="border: none;"></td>
                            <td colspan="8" style="height: 250px; border: 1.5px dashed #A6A6A6; text-align: center; vertical-align: middle; color: #BFBFBF; font-size: 10pt; text-transform: uppercase;">
                               [ ESPACIO PARA INSERTAR CROQUIS / IMAGEN DE REFERENCIA ]
                            </td>
                         </tr>
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
      const tipo = (e.tipo || "").toLowerCase();
      const isCimentacion = [
        "zapata aislada",
        "zapata corrida",
        "contratrabe",
        "dado",
        "pilas",
        "pila",
      ].includes(tipo);
      const isMuro = tipo === "muro";

      e.aceros.forEach((a) => {
        const calc = calcAceroItem(a);
        const kgTotalItem = Math.round(calc.kg * ePiezas * 100) / 100;
        totalGlobalAcero += kgTotalItem;
        const numVar = a.numVarilla || "-";

        if (isCimentacion) {
          resumenCimentacion[numVar] =
            (resumenCimentacion[numVar] || 0) + kgTotalItem;
        } else if (isMuro) {
          const grupoMuro = e.clave || (tipo.charAt(0).toUpperCase() + tipo.slice(1)) || "Muro";
          if (!resumenMuros[grupoMuro]) resumenMuros[grupoMuro] = {};
          resumenMuros[grupoMuro][numVar] = (resumenMuros[grupoMuro][numVar] || 0) + kgTotalItem;
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
                    if (isNaN(num) || num === 0) return "-";
                    return num.toFixed(2);
                  };

                  let html = `<table style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">
                      <colgroup>
                         <col width="30" />
                         <col width="70" />
                         <col width="160" />
                         <col width="120" />
                         <col width="50" />
                         <col width="100" />
                         <col width="60" />
                         <col width="60" />
                         <col width="50" />
                         <col width="70" />
                         <col width="60" />
                         <col width="50" />
                         <col width="50" />
                         <col width="60" />
                         <col width="70" />
                         <col width="70" />
                         <col width="110" />
                      </colgroup>
                      <thead>
                      <tr style="height: 15pt;">
                         <td style="border: none; background-color: #ffffff;"></td>
                         <th colspan="16" style="background-color: #0b1a30; color: #ffffff; font-size: 12pt; font-weight: bold; text-align: center; vertical-align: middle; border: 1px solid #000000; text-transform: uppercase;">
                            REPORTE GLOBAL DE ACERO (DESGLOSE POR ELEMENTO) - NIVEL: ${titleLevelName}
                         </th>
                      </tr>
                      <tr style="height: 15pt;">
                         <td colspan="17" style="border: none; background-color: #ffffff;"></td>
                      </tr>
                      <tr style="background-color: #e2e8f0; color: #000000; font-size: 8pt; font-weight: bold; text-align: center; height: 15pt;">
                         <td style="border: none; background-color: #ffffff;"></td>
                         <th style="border: 1px solid #000000; vertical-align: middle;">NIVEL</th>
                         <th style="border: 1px solid #000000; vertical-align: middle;">UBICACIÓN (EJE / CLAVE)</th>
                         <th style="border: 1px solid #000000; vertical-align: middle;">ELEMENTO</th>
                         <th style="border: 1px solid #000000; vertical-align: middle;">PZAS</th>
                         <th style="border: 1px solid #000000; vertical-align: middle;">TIPO REF.</th>
                         <th style="border: 1px solid #000000; vertical-align: middle;"># VAR</th>
                         <th style="border: 1px solid #000000; vertical-align: middle;">TRAMO</th>
                         <th style="border: 1px solid #000000; vertical-align: middle;">SEP.</th>
                         <th style="border: 1px solid #000000; vertical-align: middle;">PZAS REF.</th>
                         <th style="border: 1px solid #000000; vertical-align: middle;">LONG.</th>
                         <th style="border: 1px solid #000000; vertical-align: middle;">GAN.</th>
                         <th style="border: 1px solid #000000; vertical-align: middle;">TRAS.</th>
                         <th style="border: 1px solid #000000; vertical-align: middle;">ML/PZA</th>
                         <th style="border: 1px solid #000000; color: #0070c0; vertical-align: middle;">TOT.(ML)</th>
                         <th style="border: 1px solid #000000; color: #002060; vertical-align: middle;">TOT.(KG)</th>
                         <th style="border: 1px solid #000000; color: #002060; vertical-align: middle;">TOTAL ELEM.(KG)</th>
                      </tr>
                   </thead><tbody>`;

                  estructurasConAcero.forEach((e) => {
                    const ePiezas = parseFloat(e.piezas) || 1;
                    const elemTotalKg = calcAceroTotalKg(e.aceros, e.piezas);
                    const tipoRowSpan =
                      e.aceros.length > 0 ? e.aceros.length : 1;

                    if (e.aceros.length === 0) {
                      html += `<tr style="background-color: #ffffff; text-align: center; height: 15pt;">
                            <td style="border: none; background-color: #ffffff;"></td>
                            <td style="border: 1px solid #000000; font-size: 8pt; font-weight: bold; color: #000000; vertical-align: middle; text-align: center;">${e._nivelInfo}</td>
                            <td style="border: 1px solid #000000; font-size: 8pt; font-weight: bold; color: #000000; text-transform: uppercase; vertical-align: middle; text-align: center;">${e.eje} ${e.clave}</td>
                            <td style="border: 1px solid #000000; font-size: 8pt; color: #595959; text-transform: uppercase; vertical-align: middle; text-align: center;">${e.tipo}</td>
                            <td style="border: 1px solid #000000; font-size: 8pt; font-weight: bold; color: #1d4ed8; vertical-align: middle; text-align: center; mso-number-format:'0\\.00';">${ePiezas.toFixed(2)}</td>
                            <td colspan="11" style="border: 1px solid #000000; font-size: 8pt; color: #808080; vertical-align: middle; text-align: center;">Sin acero</td>
                            <td style="border: 1px solid #000000; font-size: 8pt; font-weight: bold; color: #1d4ed8; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${elemTotalKg.toFixed(2)}</td>
                         </tr>`;
                    } else {
                      e.aceros.forEach((a, idx) => {
                        const calc = calcAceroItem(a);
                        const kgTotal =
                          Math.round(calc.kg * ePiezas * 100) / 100;
                        html += `<tr style="background-color: #ffffff; text-align: center; height: 15pt;">
                               <td style="border: none; background-color: #ffffff;"></td>`;

                        if (idx === 0) {
                          html += `
                               <td rowspan="${tipoRowSpan}" style="border: 1px solid #000000; font-size: 8pt; font-weight: bold; color: #000000; vertical-align: middle; text-align: center;">${e._nivelInfo}</td>
                               <td rowspan="${tipoRowSpan}" style="border: 1px solid #000000; font-size: 8pt; font-weight: bold; color: #000000; text-transform: uppercase; vertical-align: middle; text-align: center;">${e.eje} ${e.clave}</td>
                               <td rowspan="${tipoRowSpan}" style="border: 1px solid #000000; font-size: 8pt; color: #595959; text-transform: uppercase; vertical-align: middle; text-align: center;">${e.tipo}</td>
                               <td rowspan="${tipoRowSpan}" style="border: 1px solid #000000; font-size: 8pt; font-weight: bold; color: #1d4ed8; vertical-align: middle; text-align: center; mso-number-format:'0\\.00';">${ePiezas.toFixed(2)}</td>`;
                        }

                        html += `
                               <td style="border: 1px solid #000000; font-size: 8pt; color: #000000; text-align: center; text-transform: uppercase; vertical-align: middle;">${a.tipo}</td>
                               <td style="border: 1px solid #000000; font-size: 8pt; color: #000000; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">#${parseFloat(a.numVarilla || 0).toFixed(2)}</td>
                               <td style="border: 1px solid #000000; font-size: 8pt; color: #1d4ed8; font-weight: bold; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${renderExVal(a.anchoCalc)}</td>
                               <td style="border: 1px solid #000000; font-size: 8pt; color: #1d4ed8; font-weight: bold; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${renderExVal(a.separacion)}</td>
                               <td style="border: 1px solid #000000; font-size: 8pt; color: #1d4ed8; font-weight: bold; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${parseFloat(a.piezas || 1).toFixed(2)}</td>
                               <td style="border: 1px solid #000000; font-size: 8pt; color: #1d4ed8; font-weight: bold; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${renderExVal(a.longitud)}</td>
                               <td style="border: 1px solid #000000; font-size: 8pt; color: #059669; font-weight: bold; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${renderExVal(a.ganchos)}</td>
                               <td style="border: 1px solid #000000; font-size: 8pt; color: #059669; font-weight: bold; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${renderExVal(a.traslapes)}</td>
                               <td style="border: 1px solid #000000; font-size: 8pt; color: #64748b; font-weight: bold; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${calc.mlPorPieza.toFixed(2)}</td>
                               <td style="border: 1px solid #000000; font-size: 8pt; color: #1d4ed8; font-weight: bold; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${(calc.ml * ePiezas).toFixed(2)}</td>
                               <td style="border: 1px solid #000000; font-size: 8pt; color: #1d4ed8; font-weight: bold; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${kgTotal.toFixed(2)}</td>`;

                        if (idx === 0) {
                          html += `<td rowspan="${tipoRowSpan}" style="border: 1px solid #000000; font-size: 8pt; color: #1d4ed8; font-weight: bold; text-align: center; vertical-align: middle; mso-number-format:'0\\.00';">${elemTotalKg.toFixed(2)}</td>`;
                        }
                        html += `</tr>`;
                      });
                    }
                  });
                  html += `</tbody></table><br/><br/>`;

                  html += `<table style="width: 100%; border: none; font-family: Arial, sans-serif;"><tr>
                     <td style="width: 30px; border: none;"></td>
                     <td style="vertical-align: top; border: none; padding-right: 15px;">
                        <table style="border-collapse: collapse; text-align: center;">
                           <thead>
                              <tr style="height: 15pt;"><th colspan="2" style="background-color: #0f766e; color: #ffffff; font-size: 12pt; border: 1px solid #bfbfbf; vertical-align: middle;">CIMENTACIÓN</th></tr>
                              <tr style="background-color: #ccfbf1; font-size: 12pt; color: #0f766e; height: 15pt;">
                                 <th style="border: 1px solid #bfbfbf; width: 100px; vertical-align: middle; font-weight: normal; font-size: 10pt;">CALIBRE</th>
                                 <th style="border: 1px solid #bfbfbf; width: 120px; vertical-align: middle; font-weight: normal; font-size: 10pt;">PESO (KG)</th>
                              </tr>
                           </thead>
                           <tbody>
                              ${
                                Object.keys(resumenCimentacion).length === 0
                                  ? `<tr style="height: 15pt;"><td colspan="2" style="border: 1px solid #bfbfbf; color: #94a3b8; font-size: 12pt; vertical-align: middle;">Sin datos</td></tr>`
                                  : Object.entries(resumenCimentacion)
                                      .sort(
                                        (a, b) => Number(a[0]) - Number(b[0]),
                                      )
                                      .map(
                                        ([num, kg]) => `
                                <tr style="background-color: #ffffff; height: 15pt;">
                                   <td style="border: 1px solid #bfbfbf; font-weight: bold; color: #0f766e; font-size: 12pt; vertical-align: middle; text-align: center; mso-number-format:'0\\.00';">#${parseFloat(num).toFixed(2)}</td>
                                   <td style="border: 1px solid #bfbfbf; font-weight: bold; color: #000000; text-align: center; font-size: 12pt; vertical-align: middle; mso-number-format:'0\\.00';">${kg.toFixed(2)}</td>
                                </tr>`,
                                      )
                                      .join("")
                              }
                           </tbody>
                        </table>
                     </td>
                     <td style="vertical-align: top; border: none; padding-right: 15px;">
                        <table style="border-collapse: collapse; text-align: center;">
                           <thead>
                              <tr style="height: 15pt;"><th colspan="2" style="background-color: #1d4ed8; color: #ffffff; font-size: 12pt; border: 1px solid #bfbfbf; vertical-align: middle;">ESTRUCTURA</th></tr>
                              <tr style="background-color: #dbeafe; font-size: 12pt; color: #1d4ed8; height: 15pt;">
                                 <th style="border: 1px solid #bfbfbf; width: 100px; vertical-align: middle; font-weight: normal; font-size: 10pt;">CALIBRE</th>
                                 <th style="border: 1px solid #bfbfbf; width: 120px; vertical-align: middle; font-weight: normal; font-size: 10pt;">PESO (KG)</th>
                              </tr>
                           </thead>
                           <tbody>
                              ${
                                Object.keys(resumenEstructura).length === 0
                                  ? `<tr style="height: 15pt;"><td colspan="2" style="border: 1px solid #bfbfbf; color: #94a3b8; font-size: 12pt; vertical-align: middle;">Sin datos</td></tr>`
                                  : Object.entries(resumenEstructura)
                                      .sort(
                                        (a, b) => Number(a[0]) - Number(b[0]),
                                      )
                                      .map(
                                        ([num, kg]) => `
                                <tr style="background-color: #ffffff; height: 15pt;">
                                   <td style="border: 1px solid #bfbfbf; font-weight: bold; color: #1d4ed8; font-size: 12pt; vertical-align: middle; text-align: center; mso-number-format:'0\\.00';">#${parseFloat(num).toFixed(2)}</td>
                                   <td style="border: 1px solid #bfbfbf; font-weight: bold; color: #000000; text-align: center; font-size: 12pt; vertical-align: middle; mso-number-format:'0\\.00';">${kg.toFixed(2)}</td>
                                </tr>`,
                                      )
                                      .join("")
                              }
                           </tbody>
                        </table>
                     </td>
                     <td style="vertical-align: top; border: none; padding-right: 15px;">
                        <table style="border-collapse: collapse; text-align: center;">
                           <thead>
                              <tr style="height: 15pt;"><th colspan="2" style="background-color: #4338ca; color: #ffffff; font-size: 12pt; border: 1px solid #bfbfbf; vertical-align: middle;">MUROS</th></tr>
                              <tr style="background-color: #e0e7ff; font-size: 12pt; color: #4338ca; height: 15pt;">
                                 <th style="border: 1px solid #bfbfbf; width: 100px; vertical-align: middle; font-weight: normal; font-size: 10pt;">CALIBRE</th>
                                 <th style="border: 1px solid #bfbfbf; width: 120px; vertical-align: middle; font-weight: normal; font-size: 10pt;">PESO (KG)</th>
                              </tr>
                           </thead>
                           <tbody>
                              ${
                                Object.keys(resumenMuros).length === 0
                                  ? `<tr style="height: 15pt;"><td colspan="2" style="border: 1px solid #bfbfbf; color: #94a3b8; font-size: 12pt; vertical-align: middle;">Sin datos</td></tr>`
                                  : Object.entries(resumenMuros)
                                      .map(([grupo, vars]) => {
                                         let rows = `<tr style="background-color: #e0e7ff; height: 15pt;"><td colspan="2" style="border: 1px solid #bfbfbf; font-weight: bold; color: #4338ca; font-size: 11pt; vertical-align: middle; text-align: center;">${grupo}</td></tr>`;
                                         rows += Object.entries(vars)
                                           .sort((a, b) => Number(a[0]) - Number(b[0]))
                                           .map(([num, kg]) => `
                                <tr style="background-color: #ffffff; height: 15pt;">
                                   <td style="border: 1px solid #bfbfbf; font-weight: bold; color: #4338ca; font-size: 12pt; vertical-align: middle; text-align: center; mso-number-format:'0\\.00';">#${parseFloat(num).toFixed(2)}</td>
                                   <td style="border: 1px solid #bfbfbf; font-weight: bold; color: #000000; text-align: center; font-size: 12pt; vertical-align: middle; mso-number-format:'0\\.00';">${kg.toFixed(2)}</td>
                                </tr>`)
                                           .join("");
                                         return rows;
                                      })
                                      .join("")
                              }
                           </tbody>
                        </table>
                     </td>
                     <td style="vertical-align: top; border: none;">
                        <table style="border-collapse: collapse; text-align: center;">
                           <thead>
                              <tr style="height: 15pt;"><th style="background-color: #0b1a30; color: #93c5fd; font-size: 12pt; border: 1px solid #bfbfbf; vertical-align: middle; width: 250px;">GRAN TOTAL ACERO (KG)</th></tr>
                           </thead>
                           <tbody>
                              <tr style="height: 30pt;"><td style="border: 1px solid #bfbfbf; font-size: 20pt; font-weight: bold; color: #1d4ed8; padding: 10px; background-color: #ffffff; vertical-align: middle; text-align: center; mso-number-format:'0\\.00';">${totalGlobalAcero.toFixed(2)}</td></tr>
                           </tbody>
                        </table>
                     </td>
                   </tr>
                   </table>`;

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
                    <th className="p-3 border-b border-slate-300">Nivel</th>
                    <th className="p-3 border-b border-slate-300">
                      Ubicación (Eje / Clave)
                    </th>
                    <th className="p-3 border-b border-slate-300">Elemento</th>
                    <th className="p-3 border-b border-slate-300 text-center">
                      Pzas
                    </th>
                    <th className="p-0 border-b border-slate-300">
                      <table className="w-full">
                        <thead>
                          <tr>
                            <th className="w-[11%] p-2 text-center border-r border-slate-300">
                              Tipo Ref.
                            </th>
                            <th className="w-[7%] p-2 text-center border-r border-slate-300">
                              # Var
                            </th>
                            <th className="w-[9%] p-2 text-center border-r border-slate-300">
                              Tramo
                            </th>
                            <th className="w-[9%] p-2 text-center border-r border-slate-300">
                              Sep.
                            </th>
                            <th className="w-[9%] p-2 text-center border-r border-slate-300">
                              Pzas
                            </th>
                            <th className="w-[9%] p-2 text-center border-r border-slate-300">
                              Long.
                            </th>
                            <th className="w-[8%] p-2 text-center border-r border-slate-300">
                              Gan.
                            </th>
                            <th className="w-[8%] p-2 text-center border-r border-slate-300">
                              Tras.
                            </th>
                            <th className="w-[10%] p-2 text-center border-r border-slate-300">
                              ML/Pza
                            </th>
                            <th className="w-[10%] p-2 text-right border-r border-slate-300">
                              Tot.(ML)
                            </th>
                            <th className="w-[10%] p-2 text-right">Tot.(KG)</th>
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
                                    <td className="w-[11%] p-2 text-center border-r border-slate-100 font-bold text-slate-600 uppercase">
                                      {a.tipo}
                                    </td>
                                    <td className="w-[7%] p-2 text-center border-r border-slate-100 font-black text-slate-800">
                                      #{a.numVarilla}
                                    </td>
                                    <td className="w-[9%] p-2 text-center border-r border-slate-100 font-bold text-indigo-900">
                                      {a.anchoCalc || "-"}
                                    </td>
                                    <td className="w-[9%] p-2 text-center border-r border-slate-100 font-bold text-indigo-900">
                                      {a.separacion || "-"}
                                    </td>
                                    <td className="w-[9%] p-2 text-center border-r border-slate-100 font-black text-blue-900">
                                      {a.piezas || "1"}
                                    </td>
                                    <td className="w-[9%] p-2 text-center border-r border-slate-100 font-black text-blue-800">
                                      {parseFloat(a.longitud || 0).toFixed(2)}
                                    </td>
                                    <td className="w-[8%] p-2 text-center border-r border-slate-100 font-black text-emerald-800">
                                      {a.ganchos || "0"}
                                    </td>
                                    <td className="w-[8%] p-2 text-center border-r border-slate-100 font-black text-emerald-800">
                                      {parseFloat(a.traslapes || 0).toFixed(2)}
                                    </td>
                                    <td className="w-[10%] p-2 text-center border-r border-slate-100 font-black text-slate-500">
                                      {calc.mlPorPieza.toFixed(2)}
                                    </td>
                                    <td className="w-[10%] p-2 text-right border-r border-slate-100 font-black text-slate-600">
                                      {(calc.ml * ePiezas).toFixed(2)}
                                    </td>
                                    <td className="w-[10%] p-2 text-right font-black text-blue-700">
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
                      .map(([grupo, vars]) => (
                        <div key={grupo} className="flex flex-col gap-1.5">
                          <span className="text-[9px] font-black text-slate-400 uppercase tracking-wider">{grupo}</span>
                          <div className="grid grid-cols-2 gap-2">
                             {Object.entries(vars).sort((a,b)=>Number(a[0])-Number(b[0])).map(([num, kg]) => (
                                <div key={num} className="bg-indigo-50/50 border border-indigo-100 rounded-lg p-2">
                                  <span className="block text-[8px] font-black text-indigo-500 uppercase">
                                    Varilla #{num}
                                  </span>
                                  <span className="font-black text-indigo-700 text-sm">
                                    {kg.toLocaleString(undefined, { minimumFractionDigits: 2 })} <span className="text-[8px]">kg</span>
                                  </span>
                                </div>
                             ))}
                          </div>
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
                <div className={`w-1.5 h-1.5 rounded-full ${isSavingAny ? "bg-amber-400 animate-pulse shadow-[0_0_8px_rgba(251,191,36,0.5)]" : "bg-teal-500 shadow-[0_0_8px_rgba(20,184,166,0.3)]"}`}></div>
                <span className={`text-[8px] font-black uppercase tracking-widest ${isSavingAny ? "text-amber-600" : "text-teal-600"}`}>
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
              <button
                onClick={() => {
                  setShowSettings(!showSettings);
                  setShowWallGenerator(false);
                  setShowStructureGenerator(false);
                }}
                className={`px-4 py-2.5 rounded-xl font-black uppercase text-[10px] shadow-lg flex items-center gap-2 ${showSettings && !showWallGenerator && !showStructureGenerator ? "bg-blue-600 text-white" : "bg-white border text-slate-700 hover:bg-slate-50"}`}
              >
                <Settings size={16} /> Ajustes
              </button>
            {activePartidaId && (
              <>
                <button
                  onClick={() => {
                    setShowStructureGenerator(!showStructureGenerator);
                    setShowWallGenerator(false);
                    setShowSettings(false);
                  }}
                  className={`px-4 py-2.5 rounded-xl font-black uppercase text-[10px] shadow-lg transition-all flex items-center gap-2 ${showStructureGenerator ? "bg-amber-600 text-white" : "bg-white border text-slate-700 hover:bg-amber-50"}`}
                >
                  <Box size={16} /> Estructura
                </button>
                <button
                  onClick={() => {
                    setShowWallGenerator(!showWallGenerator);
                    setShowStructureGenerator(false);
                    setShowSettings(false);
                  }}
                  className={`px-4 py-2.5 rounded-xl font-black uppercase text-[10px] shadow-lg transition-all flex items-center gap-2 ${showWallGenerator ? "bg-indigo-700 text-white" : "bg-white border text-slate-700 hover:bg-indigo-50"}`}
                >
                  <Building2 size={16} /> Muros
                </button>
              </>
            )}
          </div>
        </div>
      </header>
      <main className="max-w-7xl mx-auto w-full flex-1 relative">
        {showStructureGenerator ? (
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
            <div className="overflow-x-auto bg-white">
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
                    ></th>
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
                        <td className="p-0 border-r">
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
                            <option value="losa de vigueta">Losa de Vigueta</option>
                            <option value="columna">Columna</option>
                            <option value="columna circular">Columna Circular</option>
                            <option value="pilas">Pilas</option>
                            <option value="trabe">Trabe</option>
                            <option value="nervadura">Nervadura</option>
                          </select>
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
                          <DebouncedCell
                            value={e.alto}
                            onChange={(v) =>
                              updateEstructuraField(e.id, "alto", v)
                            }
                            className="w-full p-1.5 bg-transparent text-center text-blue-900 font-bold text-xs"
                          />
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
                          className="p-0 border-r bg-sky-50 transition-colors hover:bg-sky-100 cursor-pointer"
                          onClick={() =>
                            ["losa nervada", "nervadura"].includes(
                              e.tipo?.toLowerCase(),
                            )
                              ? setActiveCasetonSubmodal(e.id)
                              : null
                          }
                        >
                          {["losa nervada", "nervadura"].includes(
                            e.tipo?.toLowerCase(),
                          ) ? (
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
                          ) : (
                            <div className="w-full h-full p-1 text-center font-black text-sky-700 flex justify-center items-center text-sm cursor-default">
                              {concreto > 0 ? concreto.toFixed(2) : "-"}
                            </div>
                          )}
                        </td>
                        <td className="p-1 border-r text-center font-black text-amber-700 bg-amber-50 text-sm">
                          {cimbra > 0 ? cimbra.toFixed(2) : "-"}
                        </td>
                        <td className="p-1 border-r text-center font-black text-orange-700 bg-orange-50 text-sm">
                          {cimbraFrontera > 0 ? cimbraFrontera.toFixed(2) : "-"}
                        </td>
                        <td
                          className={`p-0 border-r ${e.tipo?.toLowerCase() === "losa de vigueta" ? "bg-slate-50 cursor-not-allowed" : "bg-slate-100/50 hover:bg-slate-200 cursor-pointer"} transition-colors`}
                          onClick={() => {
                            if (e.tipo?.toLowerCase() !== "losa de vigueta") {
                              setActiveSteelSubmodal(e.id);
                            }
                          }}
                        >
                          <div className={`w-full h-full p-2 text-center font-black flex justify-center items-center gap-1 group ${e.tipo?.toLowerCase() === "losa de vigueta" ? "opacity-30" : "text-slate-700"}`}>
                            {aceroKg > 0 ? (
                              <span className="text-sm">
                                {aceroKg.toFixed(2)}
                              </span>
                            ) : (
                              <span className="text-slate-400 font-normal text-[10px]">
                                0.00
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
                         <th colspan="4" style="border: 1px solid #b4c6e7; background-color: #deebf7; color: #2f5496; vertical-align: middle;">VOLÚMENES / ÁREAS TOTALES</th>
                      </tr>
                      <tr style="font-size: 8pt; font-weight: bold; height: 15pt; text-align: center;">
                         <td style="border: none; background-color: #ffffff;"></td>
                         <th style="background-color: #deebf7; color: #2f5496; border: 1px solid #b4c6e7; vertical-align: middle;">CONCRETO (M3)</th>
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
                            <th colspan="17" style="background-color: #e36c09; color: white; font-size: 12pt; padding: 6px; border: 1px solid #e36c09; vertical-align: middle;">
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
                            <th style="border: 1px solid #f8cbad; width: 120px; vertical-align: middle;">OBTURACIÓN (M2)</th>
                            <th style="border: 1px solid #f8cbad; width: 120px; vertical-align: middle;">MALLA REF. (M2)</th>
                            <th style="border: 1px solid #f8cbad; width: 120px; vertical-align: middle;">PASOS MUROS (PZA)</th>
                            <th style="border: 1px solid #f8cbad; width: 120px; vertical-align: middle;">EMPLAYER (M2)</th>
                            <th style="border: 1px solid #f8cbad; width: 120px; vertical-align: middle;">ANDAMIAJE</th>
                            <th style="border: 1px solid #f8cbad; width: 140px; vertical-align: middle;">ACERO TOTAL (KG)</th>
                         </tr>
                      </thead>
                      <tbody>`;

                  Object.entries(estructuraSummary.breakdown).forEach(
                    ([tipo, data]) => {
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
                        data.cimbraColumnas_6_9 === 0
                      )
                        return;
                      const isCim = data.isCimentacion;
                      const concText = isCim
                        ? "-"
                        : data.concreto > 0
                          ? data.concreto.toFixed(2)
                          : "-";
                      
                      const isColumnGrp = tipo === "Columnas" || tipo === "Columnas Circulares";
                      const andamiajeVal = isColumnGrp 
                                            ? data.andamiajeColumnas 
                                            : tipo === "Trabes" 
                                              ? data.andamiajeTrabes 
                                              : 0;

                      html += `<tr style="text-align: center; background-color: #ffffff; height: 15pt;">
                         <td style="border: none;"></td>
                         <td style="border: 1px solid #f8cbad; font-weight: bold; text-align: left; padding-left: 8px; color: #000000; font-size: 10pt; vertical-align: middle;">${tipo}</td>
                         <td style="border: 1px solid #f8cbad; color: #2f5496; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${concText}</td>
                         <td style="border: 1px solid #f8cbad; color: #2f5496; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${data.casetones > 0 ? data.casetones.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #e36c09; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${data.cimbra > 0 && tipo !== "Muros" && !isColumnGrp ? data.cimbra.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #e36c09; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${(tipo === "Muros" && data.cimbraMuros_0_3 > 0) ? data.cimbraMuros_0_3.toFixed(2) : (isColumnGrp && data.cimbraColumnas_0_3 > 0) ? data.cimbraColumnas_0_3.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #e36c09; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${(tipo === "Muros" && data.cimbraMuros_3_6 > 0) ? data.cimbraMuros_3_6.toFixed(2) : (isColumnGrp && data.cimbraColumnas_3_6 > 0) ? data.cimbraColumnas_3_6.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #e36c09; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${(tipo === "Muros" && data.cimbraMuros_6_9 > 0) ? data.cimbraMuros_6_9.toFixed(2) : (isColumnGrp && data.cimbraColumnas_6_9 > 0) ? data.cimbraColumnas_6_9.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #e36c09; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${(tipo === "Muros" && data.cimbraMuroCurvo_0_3 > 0) ? data.cimbraMuroCurvo_0_3.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #e36c09; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${(tipo === "Muros" && data.cimbraMuroCurvo_3_6 > 0) ? data.cimbraMuroCurvo_3_6.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #e36c09; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${(tipo === "Muros" && data.cimbraMuroCurvo_6_9 > 0) ? data.cimbraMuroCurvo_6_9.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #e36c09; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${data.cimbraFrontera > 0 ? data.cimbraFrontera.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #548235; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${data.obturacionMuros > 0 ? data.obturacionMuros.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #548235; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${data.mallaRefuerzo > 0 ? data.mallaRefuerzo.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #548235; font-weight: bold; font-size: 10pt; vertical-align: middle;">${data.pasosMuros > 0 ? data.pasosMuros : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #548235; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${data.emplayerColumnas > 0 ? data.emplayerColumnas.toFixed(2) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #548235; font-weight: bold; font-size: 10pt; vertical-align: middle; ${isColumnGrp ? '' : `mso-number-format:'0\\.00';`}">${andamiajeVal > 0 ? (isColumnGrp ? andamiajeVal : andamiajeVal.toFixed(2)) : "-"}</td>
                         <td style="border: 1px solid #f8cbad; color: #2f5496; font-weight: bold; font-size: 10pt; vertical-align: middle; mso-number-format:'0\\.00';">${data.aceroKg > 0 ? data.aceroKg.toFixed(2) : "-"}</td>
                      </tr>`;
                    },
                  );
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
                    {estructuraSummary.totalAceroEstructuraGlobal > 0 && (
                      <div className="bg-blue-50/50 border border-blue-200 rounded-2xl p-4 flex flex-col xl:flex-row items-center gap-4 shadow-sm">
                        <div className="font-black text-blue-900 uppercase tracking-widest text-sm w-40 shrink-0 flex flex-col">
                          <span>Acero Estructura</span>
                          <span className="text-[8px] opacity-60">
                            (Losas, Trabes, Col, Col. Circ, Nerv)
                          </span>
                        </div>
                        <div className="flex flex-wrap gap-3 flex-1">
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
                    {Object.entries(estructuraSummary.breakdown).map(
                      ([tipo, data]) => {
                        if (
                          data.isCimentacion &&
                          data.cimbra === 0 &&
                          data.cimbraFrontera === 0 &&
                          data.excavacion === 0 &&
                          data.relleno === 0
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
                              {data.casetones > 0 && (
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
                              {data.cimbra > 0 && tipo !== "Muros" && tipo !== "Columnas" && tipo !== "Columnas Circulares" && (
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
                                  <button onClick={() => handleCopyStructure(tipo, "cimbraMuros_0_3")} className="text-amber-500 hover:text-amber-700">
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
                                  <button onClick={() => handleCopyStructure(tipo, "cimbraMuros_3_6")} className="text-orange-500 hover:text-orange-700">
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
                                  <button onClick={() => handleCopyStructure(tipo, "cimbraMuros_6_9")} className="text-red-500 hover:text-red-700">
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {tipo === "Muros" && data.cimbraMuroCurvo_0_3 > 0 && (
                                <div className="bg-amber-100 border border-amber-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-amber-700 uppercase">
                                      Cimbra Curva 0-3m
                                    </span>
                                    <span className="text-sm font-black text-amber-900">
                                      {data.cimbraMuroCurvo_0_3.toFixed(2)} m2/ml
                                    </span>
                                  </div>
                                  <button onClick={() => handleCopyStructure(tipo, "cimbraMuroCurvo_0_3")} className="text-amber-500 hover:text-amber-700">
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {tipo === "Muros" && data.cimbraMuroCurvo_3_6 > 0 && (
                                <div className="bg-orange-100 border border-orange-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-orange-700 uppercase">
                                      Cimbra Curva 3-6m
                                    </span>
                                    <span className="text-sm font-black text-orange-900">
                                      {data.cimbraMuroCurvo_3_6.toFixed(2)} m2/ml
                                    </span>
                                  </div>
                                  <button onClick={() => handleCopyStructure(tipo, "cimbraMuroCurvo_3_6")} className="text-orange-500 hover:text-orange-700">
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {tipo === "Muros" && data.cimbraMuroCurvo_6_9 > 0 && (
                                <div className="bg-red-100 border border-red-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-red-700 uppercase">
                                      Cimbra Curva {">"} 6m
                                    </span>
                                    <span className="text-sm font-black text-red-900">
                                      {data.cimbraMuroCurvo_6_9.toFixed(2)} m2/ml
                                    </span>
                                  </div>
                                  <button onClick={() => handleCopyStructure(tipo, "cimbraMuroCurvo_6_9")} className="text-red-500 hover:text-red-700">
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {(tipo === "Columnas" || tipo === "Columnas Circulares") && data.cimbraColumnas_0_3 > 0 && (
                                <div className="bg-amber-100 border border-amber-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-amber-700 uppercase">
                                      Cimbra 0-3m
                                    </span>
                                    <span className="text-sm font-black text-amber-900">
                                      {data.cimbraColumnas_0_3.toFixed(2)} m2/ml
                                    </span>
                                  </div>
                                  <button onClick={() => handleCopyStructure(tipo, "cimbraColumnas_0_3")} className="text-amber-500 hover:text-amber-700">
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {(tipo === "Columnas" || tipo === "Columnas Circulares") && data.cimbraColumnas_3_6 > 0 && (
                                <div className="bg-orange-100 border border-orange-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-orange-700 uppercase">
                                      Cimbra 3-6m
                                    </span>
                                    <span className="text-sm font-black text-orange-900">
                                      {data.cimbraColumnas_3_6.toFixed(2)} m2/ml
                                    </span>
                                  </div>
                                  <button onClick={() => handleCopyStructure(tipo, "cimbraColumnas_3_6")} className="text-orange-500 hover:text-orange-700">
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {(tipo === "Columnas" || tipo === "Columnas Circulares") && data.cimbraColumnas_6_9 > 0 && (
                                <div className="bg-red-100 border border-red-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-red-700 uppercase">
                                      Cimbra {">"} 6m
                                    </span>
                                    <span className="text-sm font-black text-red-900">
                                      {data.cimbraColumnas_6_9.toFixed(2)} m2/ml
                                    </span>
                                  </div>
                                  <button onClick={() => handleCopyStructure(tipo, "cimbraColumnas_6_9")} className="text-red-500 hover:text-red-700">
                                    <Clipboard size={14} />
                                  </button>
                                </div>
                              )}
                              {data.obturacionMuros > 0 && (
                                <div className="bg-white border border-slate-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-slate-500 uppercase">Obturación Muros</span>
                                    <span className="text-sm font-black text-slate-800">{data.obturacionMuros.toFixed(2)} m2/ml</span>
                                  </div>
                                  <button onClick={() => handleCopyStructure(tipo, "obturacionMuros")} className="text-slate-400 hover:text-slate-600"><Clipboard size={14} /></button>
                                </div>
                              )}
                              {data.mallaRefuerzo > 0 && (
                                <div className="bg-white border border-slate-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-slate-500 uppercase">Malla Refuerzo</span>
                                    <span className="text-sm font-black text-slate-800">{data.mallaRefuerzo.toFixed(2)} m2/ml</span>
                                  </div>
                                  <button onClick={() => handleCopyStructure(tipo, "mallaRefuerzo")} className="text-slate-400 hover:text-slate-600"><Clipboard size={14} /></button>
                                </div>
                              )}
                              {data.pasosMuros > 0 && (
                                <div className="bg-white border border-slate-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-slate-500 uppercase">Pasos Muros</span>
                                    <span className="text-sm font-black text-slate-800">{data.pasosMuros} pzas</span>
                                  </div>
                                  <button onClick={() => handleCopyStructure(tipo, "pasosMuros")} className="text-slate-400 hover:text-slate-600"><Clipboard size={14} /></button>
                                </div>
                              )}
                              {data.emplayerColumnas > 0 && (
                                <div className="bg-white border border-slate-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-slate-500 uppercase">Emplayer</span>
                                    <span className="text-sm font-black text-slate-800">{data.emplayerColumnas.toFixed(2)} m2</span>
                                  </div>
                                  <button onClick={() => handleCopyStructure(tipo, "emplayerColumnas")} className="text-slate-400 hover:text-slate-600"><Clipboard size={14} /></button>
                                </div>
                              )}
                              {data.andamiajeColumnas > 0 && (
                                <div className="bg-white border border-slate-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-slate-500 uppercase">Andamiaje Col ({">"}3m)</span>
                                    <span className="text-sm font-black text-slate-800">{data.andamiajeColumnas} pzas</span>
                                  </div>
                                  <button onClick={() => handleCopyStructure(tipo, "andamiajeColumnas")} className="text-slate-400 hover:text-slate-600"><Clipboard size={14} /></button>
                                </div>
                              )}
                              {data.andamiajeTrabes > 0 && (
                                <div className="bg-white border border-slate-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-slate-500 uppercase">Andamiaje Trabes</span>
                                    <span className="text-sm font-black text-slate-800">{data.andamiajeTrabes.toFixed(2)} ml</span>
                                  </div>
                                  <button onClick={() => handleCopyStructure(tipo, "andamiajeTrabes")} className="text-slate-400 hover:text-slate-600"><Clipboard size={14} /></button>
                                </div>
                              )}
                              {data.excavacion > 0 && (
                                <div className="bg-amber-100 border border-amber-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-amber-700 uppercase">Excavación</span>
                                    <span className="text-sm font-black text-amber-900">{data.excavacion.toFixed(2)} m3</span>
                                  </div>
                                  <button onClick={() => handleCopyStructure(tipo, "excavacion")} className="text-amber-500 hover:text-amber-700"><Clipboard size={14} /></button>
                                </div>
                              )}
                              {data.relleno > 0 && (
                                <div className="bg-orange-100 border border-orange-200 rounded-xl p-3 flex items-center justify-between gap-4 min-w-[140px] shadow-sm">
                                  <div>
                                    <span className="block text-[8px] font-black text-orange-700 uppercase">Relleno</span>
                                    <span className="text-sm font-black text-orange-900">{data.relleno.toFixed(2)} m3</span>
                                  </div>
                                  <button onClick={() => handleCopyStructure(tipo, "relleno")} className="text-orange-500 hover:text-orange-700"><Clipboard size={14} /></button>
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
                              {data.aceroKg > 0 &&
                                !data.isCimentacion &&
                                ![
                                  "losa",
                                  "losa nervada",
                                  "trabe",
                                  "nervadura",
                                  "columna",
                                  "columna circular",
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
                      },
                    )}
                  </div>
                </div>
              </div>
            </div>
          </div>
        ) : showWallGenerator ? (
           <div className="bg-white rounded-3xl shadow-xl border border-gray-100 animate-in fade-in zoom-in-95 duration-300 flex flex-col">
             <div className="bg-[#312e81] text-white p-4 flex justify-between items-center shrink-0">
               <div className="flex items-center gap-3"><div className="p-2 bg-indigo-800/50 rounded-lg border border-indigo-600/50"><Building2 size={20} className="text-indigo-200" /></div><div><h2 className="text-lg font-black uppercase tracking-wider leading-tight">Matriz de Muros</h2><p className="text-[9px] text-indigo-300 font-bold uppercase tracking-widest">{activePartida.nombre}</p></div></div>
               <div className="flex items-center gap-3"><select value={activeLevelId} onChange={(e) => setActiveLevelId(e.target.value)} className="bg-[#1e1b4b] text-white font-black text-[10px] uppercase p-2 px-4 rounded-lg border border-[#4338ca] outline-none cursor-pointer">{niveles.map(n => <option key={n.id} value={n.id} className="text-slate-800">{n.nombre}</option>)}</select><button onClick={() => setShowWallGenerator(false)} className="p-1.5 bg-[#4338ca] hover:bg-indigo-500 rounded-full transition-colors"><X size={20} /></button></div>
             </div>
             <div className="overflow-x-auto bg-white">
               <table className="w-full text-left border-collapse text-[10px] xl:text-xs">
                 <thead className="bg-slate-50 text-slate-500 font-black uppercase tracking-wider border-b-2 border-slate-200 sticky top-0 z-30 shadow-sm text-[9px] md:text-[10px]">
                   <tr>
                     <th className="p-2 text-center border-r border-slate-200 relative select-none bg-slate-50" rowSpan={2} style={{ width: wallColWidths.select || 40 }}><button onClick={toggleAllWallRows} className="hover:text-indigo-600 transition-colors">{muros.length > 0 && selectedWallRows.length === muros.length ? <CheckSquare size={14}/> : <Square size={14}/>}</button></th>
                     <th className="p-2 text-center border-r border-slate-200" rowSpan={2}>No.</th><th className="p-2 border-r border-slate-200 text-center" rowSpan={2}>Eje</th><th className="p-2 border-r border-slate-200 text-center" rowSpan={2}>Clave</th><th className="p-2 text-center border-r border-slate-200 text-blue-700 bg-slate-50" rowSpan={2}>Largo</th><th className="p-2 text-center border-r border-slate-200 text-blue-700 bg-slate-50" rowSpan={2}>Ancho</th><th className="p-2 text-center border-r border-slate-200 text-blue-700 bg-slate-50" rowSpan={2}>Alto</th><th className="p-2 text-center border-r border-slate-200 bg-slate-200 text-slate-700" rowSpan={2}>Bruto</th><th className="p-2 text-center border-r border-slate-200 bg-red-50 text-red-600" rowSpan={2}>-Huecos</th><th className="p-2 text-center border-r border-slate-200 bg-orange-50 text-orange-600" rowSpan={2}>-Cast.</th><th className="p-2 text-center border-r border-slate-200 bg-emerald-100 text-emerald-800 font-black" rowSpan={2}>Neta</th>
                     <th className="p-1 border-b border-r border-slate-200 text-center text-purple-700" colSpan={2}>Aplanado</th><th className="p-1 border-b border-r border-slate-200 text-center text-fuchsia-700" colSpan={2}>Recubrimiento</th><th className="p-2 text-center border-l border-slate-200" rowSpan={2}></th>
                   </tr>
                   <tr>
                     <th className="p-1 border-r border-slate-200 text-center text-purple-700">C1</th><th className="p-1 border-r border-slate-200 text-center text-purple-700">C2</th><th className="p-1 border-r border-slate-200 text-center text-fuchsia-700">C1</th><th className="p-1 border-r border-slate-200 text-center text-fuchsia-700">C2</th>
                   </tr>
                 </thead>
                 <tbody className="divide-y divide-slate-100">
                   {muros.map((m, i) => {
                     const bruto = calcWallArea(m.largo, m.alto), dedH = getWallHuecosTotal(m.huecos), dedC = getWallCastillosTotal(m.castillos, m.alto), neto = Math.max(0, bruto - dedH - dedC);
                     return (
                       <tr key={m.id} className={`transition-colors ${selectedWallRows.includes(m.id) ? 'bg-indigo-50/50' : 'hover:bg-slate-50'}`}>
                         <td className="p-1 text-center border-r border-slate-200"><button onClick={() => toggleWallRow(m.id)} className={`${selectedWallRows.includes(m.id) ? 'text-indigo-600' : 'text-slate-300 hover:text-indigo-500'} transition-colors`}>{selectedWallRows.includes(m.id) ? <CheckSquare size={14} /> : <Square size={14} />}</button></td>
                         <td className="p-1.5 text-center font-black text-slate-400 border-r">{i + 1}</td>
                         <td className="p-0 border-r"><DebouncedCell value={m.eje} onChange={v => updateActiveMuros(prev => prev.map(x => x.id === m.id ? {...x, eje: v} : x))} className="w-full p-2 bg-transparent font-bold uppercase outline-none text-[10px] md:text-xs text-center" /></td>
                         <td className="p-0 border-r"><DebouncedCell value={m.clave} onChange={v => updateActiveMuros(prev => prev.map(x => x.id === m.id ? {...x, clave: v} : x))} className="w-full p-2 bg-transparent font-bold uppercase outline-none text-[10px] md:text-xs text-center" /></td>
                         <td className="p-0 border-r bg-blue-50/10"><DebouncedCell value={m.largo} onChange={v => updateActiveMuros(prev => prev.map(x => x.id === m.id ? {...x, largo: v} : x))} className="w-full p-2 bg-transparent text-center text-blue-900 font-bold text-[10px] md:text-xs" /></td>
                         <td className="p-0 border-r bg-blue-50/10"><DebouncedCell value={m.ancho} onChange={v => updateActiveMuros(prev => prev.map(x => x.id === m.id ? {...x, ancho: v} : x))} className="w-full p-2 bg-transparent text-center text-blue-900 font-bold text-[10px] md:text-xs" /></td>
                         <td className="p-0 border-r bg-blue-50/10"><DebouncedCell value={m.alto} onChange={v => updateActiveMuros(prev => prev.map(x => x.id === m.id ? {...x, alto: v} : x))} className="w-full p-2 bg-transparent text-center text-blue-900 font-bold text-[10px] md:text-xs" /></td>
                         <td className="p-2 border-r text-center font-black text-slate-500 bg-slate-50/50 text-[10px] md:text-xs">{bruto > 0 ? bruto.toFixed(2) : '-'}</td>
                         <td className="p-0 border-r bg-red-50/30 hover:bg-red-100 transition-colors cursor-pointer" onClick={() => setActiveWallSubmodal({wallId: m.id, type: 'huecos'})}><div className="w-full h-full p-2 text-center font-black text-red-600 text-[10px] md:text-xs">{dedH > 0 ? dedH.toFixed(2) : '0.00'}</div></td>
                         <td className="p-0 border-r bg-orange-50/30 hover:bg-orange-100 transition-colors cursor-pointer" onClick={() => setActiveWallSubmodal({wallId: m.id, type: 'castillos'})}><div className="w-full h-full p-2 text-center font-black text-orange-600 text-[10px] md:text-xs">{dedC > 0 ? dedC.toFixed(2) : '0.00'}</div></td>
                         <td className="p-2 border-r text-center font-black text-emerald-700 bg-emerald-50 text-[10px] md:text-sm">{neto > 0 ? neto.toFixed(2) : '-'}</td>
                         <td className="p-0 border-r"><DebouncedCell value={m.tipoAplanadoC1} onChange={v => updateActiveMuros(prev => prev.map(x => x.id === m.id ? {...x, tipoAplanadoC1: v} : x))} className="w-full p-2 bg-transparent text-center font-bold text-slate-700 text-[9px] md:text-[10px] uppercase" /></td>
                         <td className="p-0 border-r"><DebouncedCell value={m.tipoAplanadoC2} onChange={v => updateActiveMuros(prev => prev.map(x => x.id === m.id ? {...x, tipoAplanadoC2: v} : x))} className="w-full p-2 bg-transparent text-center font-bold text-slate-700 text-[9px] md:text-[10px] uppercase" /></td>
                         <td className="p-0 border-r"><DebouncedCell value={m.tipoRecubrimientoC1} onChange={v => updateActiveMuros(prev => prev.map(x => x.id === m.id ? {...x, tipoRecubrimientoC1: v} : x))} className="w-full p-2 bg-transparent text-center font-bold text-slate-700 text-[9px] md:text-[10px] uppercase" /></td>
                         <td className="p-0 border-r"><DebouncedCell value={m.tipoRecubrimientoC2} onChange={v => updateActiveMuros(prev => prev.map(x => x.id === m.id ? {...x, tipoRecubrimientoC2: v} : x))} className="w-full p-2 bg-transparent text-center font-bold text-slate-700 text-[9px] md:text-[10px] uppercase" /></td>
                         <td className="p-1 text-center border-l border-slate-200"><div className="flex justify-center items-center gap-0.5"><button onClick={() => handleCopyWallRow(m)} className="p-1 rounded text-slate-400 hover:text-indigo-600"><Copy size={14} /></button><button onClick={() => handlePasteWallRow(m.id)} disabled={!copiedWallRow} className="p-1 rounded text-slate-400 disabled:opacity-10"><ClipboardPaste size={14} /></button><button onClick={() => updateActiveMuros(prev => prev.filter(x => x.id !== m.id))} className="text-slate-300 hover:text-red-500 p-1"><Trash2 size={14} /></button></div></td>
                       </tr>
                     );
                   })}
                 </tbody>
               </table>
             </div>
             <div className="p-4 bg-white border-t border-slate-100 flex gap-3 flex-wrap shrink-0">
               <button onClick={() => updateActiveMuros(prev => { const lastRow = prev[prev.length - 1]; const nextClave = lastRow ? getNextClaveValue(lastRow.clave) : ''; return [...prev, { id: `M-${Date.now()}`, eje: lastRow?.eje || '', clave: nextClave, largo: '', ancho: '', alto: '', huecos: [], castillos: [], tipoAplanadoC1: '', tipoAplanadoC2: '', tipoRecubrimientoC1: '', tipoRecubrimientoC2: '' }]; })} className="px-6 py-2 border-2 border-dashed border-indigo-200 text-indigo-600 font-black rounded-lg hover:bg-indigo-50 text-[10px] uppercase tracking-wider">+ Agregar Muro</button>
               <button onClick={handleCopySelectedWalls} className={`px-6 py-2 rounded-lg font-black text-[10px] uppercase flex items-center gap-2 transition-colors ${selectedWallRows.length > 0 ? 'bg-indigo-600 text-white hover:bg-indigo-700 shadow-md' : 'bg-slate-100 hover:bg-slate-200 text-slate-600'}`}><Copy size={14}/> {copyStatus === 'copied-wall-list' ? '¡Copiado!' : selectedWallRows.length > 0 ? `Copiar (${selectedWallRows.length})` : 'Copiar Todas'}</button>
               <button onClick={handlePasteSelectedWalls} disabled={!copiedWallRowsList || copiedWallRowsList.length === 0} className={`px-6 py-2 rounded-lg font-black text-[10px] uppercase flex items-center gap-2 transition-all ${copiedWallRowsList && copiedWallRowsList.length > 0 ? 'bg-emerald-100 hover:bg-emerald-200 text-emerald-700 shadow-sm cursor-pointer' : 'bg-slate-50 text-slate-300 cursor-not-allowed'}`}><ClipboardPaste size={14}/> Pegar Datos</button>
               <button onClick={() => {
                   const levelNameExport = niveles.find(n=>n.id===activeLevelId)?.nombre||'TODOS LOS NIVELES';
                   const renderExVal = (v) => {
                       const num = parseFloat(v);
                       if (isNaN(num) || num === 0) return '-';
                       return num.toFixed(2);
                   };

                   let html = `<table style="border-collapse: collapse; font-family: Arial, sans-serif;">
                      <colgroup>
                         <col width="40" style="width: 40px;" />
                         <col width="120" style="width: 120px;" />
                         <col width="120" style="width: 120px;" />
                         <col width="120" style="width: 120px;" />
                         <col width="120" style="width: 120px;" />
                         <col width="120" style="width: 120px;" />
                         <col width="120" style="width: 120px;" />
                         <col width="120" style="width: 120px;" />
                         <col width="120" style="width: 120px;" />
                         <col width="120" style="width: 120px;" />
                         <col width="120" style="width: 120px;" />
                         <col width="280" style="width: 280px;" />
                         <col width="280" style="width: 280px;" />
                         <col width="280" style="width: 280px;" />
                         <col width="280" style="width: 280px;" />
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
                         <th rowspan="2" style="border: 1px solid #c7d2fe; width: 120px; vertical-align: middle;">NO.</th>
                         <th rowspan="2" style="border: 1px solid #c7d2fe; width: 120px; vertical-align: middle;">EJE</th>
                         <th rowspan="2" style="border: 1px solid #c7d2fe; width: 120px; vertical-align: middle;">CLAVE</th>
                         <th rowspan="2" style="border: 1px solid #c7d2fe; width: 120px; vertical-align: middle;">LARGO (M)</th>
                         <th rowspan="2" style="border: 1px solid #c7d2fe; width: 120px; vertical-align: middle;">ANCHO (M)</th>
                         <th rowspan="2" style="border: 1px solid #c7d2fe; width: 120px; vertical-align: middle;">ALTO (M)</th>
                         <th rowspan="2" style="border: 1px solid #c7d2fe; width: 120px; vertical-align: middle;">BRUTO (M2)</th>
                         <th rowspan="2" style="border: 1px solid #c7d2fe; width: 120px; vertical-align: middle;">-HUECOS (M2)</th>
                         <th rowspan="2" style="border: 1px solid #c7d2fe; width: 120px; vertical-align: middle;">-CAST. (M2)</th>
                         <th rowspan="2" style="border: 1px solid #c7d2fe; width: 120px; vertical-align: middle;">NETA (M2)</th>
                         <th colspan="2" style="border: 1px solid #c7d2fe; width: 560px; vertical-align: middle;">APLANADO</th>
                         <th colspan="2" style="border: 1px solid #c7d2fe; width: 560px; vertical-align: middle;">RECUBRIMIENTO</th>
                      </tr>
                      <tr style="background-color: #e6e6fa; color: #312eb5; font-size: 8pt; font-family: Arial, sans-serif; height: 15pt; text-align: center;">
                         <td style="border: none; background-color: #ffffff;"></td>
                         <th style="border: 1px solid #c7d2fe; width: 280px; vertical-align: middle;">C1</th>
                         <th style="border: 1px solid #c7d2fe; width: 280px; vertical-align: middle;">C2</th>
                         <th style="border: 1px solid #c7d2fe; width: 280px; vertical-align: middle;">C1</th>
                         <th style="border: 1px solid #c7d2fe; width: 280px; vertical-align: middle;">C2</th>
                      </tr>
                   </thead><tbody>`;

                   muros.forEach((m, i) => {
                      const bruto = calcWallArea(m.largo, m.alto), dedH = getWallHuecosTotal(m.huecos), dedC = getWallCastillosTotal(m.castillos, m.alto), neto = Math.max(0, bruto - dedH - dedC);
                      html += `<tr style="text-align: center; background-color: #ffffff; height: 18pt;">
                         <td style="border: none; background-color: #ffffff;"></td>
                         <td style="border: 1px solid #c7d2fe; color: #a6a6a6; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; text-align: center; mso-number-format:'0';">${i+1}</td>
                         <td style="border: 1px solid #c7d2fe; color: #312eb5; font-weight: bold; text-transform: uppercase; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; text-align: center;">${m.eje || ''}</td>
                         <td style="border: 1px solid #c7d2fe; color: #000000; font-weight: bold; text-transform: uppercase; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; text-align: center;">${m.clave || ''}</td>
                         <td style="border: 1px solid #c7d2fe; color: #000066; font-weight: bold; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; mso-number-format:'0\.00'; text-align: center;">${renderExVal(m.largo)}</td>
                         <td style="border: 1px solid #c7d2fe; color: #000066; font-weight: bold; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; mso-number-format:'0\.00'; text-align: center;">${renderExVal(m.ancho)}</td>
                         <td style="border: 1px solid #c7d2fe; color: #000066; font-weight: bold; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; mso-number-format:'0\.00'; text-align: center;">${renderExVal(m.alto)}</td>
                         <td style="border: 1px solid #c7d2fe; color: #404040; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; mso-number-format:'0\.00'; text-align: center;">${renderExVal(bruto)}</td>
                         <td style="border: 1px solid #c7d2fe; color: #c00000; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; mso-number-format:'0\.00'; text-align: center;">${dedH > 0 ? dedH.toFixed(2) : '-'}</td>
                         <td style="border: 1px solid #c7d2fe; color: #e36c0a; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; mso-number-format:'0\.00'; text-align: center;">${dedC > 0 ? dedC.toFixed(2) : '-'}</td>
                         <td style="border: 1px solid #c7d2fe; color: #00b050; font-weight: bold; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; mso-number-format:'0\.00'; text-align: center;">${renderExVal(neto)}</td>
                         <td style="border: 1px solid #c7d2fe; color: #800080; font-weight: bold; text-transform: uppercase; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; text-align: center;">${m.tipoAplanadoC1 || ''}</td>
                         <td style="border: 1px solid #c7d2fe; color: #800080; font-weight: bold; text-transform: uppercase; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; text-align: center;">${m.tipoAplanadoC2 || ''}</td>
                         <td style="border: 1px solid #c7d2fe; color: #800080; font-weight: bold; text-transform: uppercase; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; text-align: center;">${m.tipoRecubrimientoC1 || ''}</td>
                         <td style="border: 1px solid #c7d2fe; color: #800080; font-weight: bold; text-transform: uppercase; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; text-align: center;">${m.tipoRecubrimientoC2 || ''}</td>
                      </tr>`;
                   });
                   html += `</tbody></table><br/><br/><br/>`;

                   html += `<table style="width: 100%; border: none; font-family: Arial, sans-serif;"><tr>
                      <td style="width: 40px; border: none;"></td>
                      <td style="vertical-align: top; border: none; padding-right: 20px;">
                         <table style="border-collapse: collapse; text-align: center; width: 250px;">
                            <thead><tr style="height: 20pt;"><th colspan="2" style="background-color: #00b050; color: white; font-size: 9pt; font-family: Arial, sans-serif; padding: 8px; border: 1px solid #00b050; vertical-align: middle; text-transform: uppercase;">ÁREA NETA POR ESPESOR</th></tr></thead>
                            <tbody>`;
                            if(Object.keys(wallSummary.murosPorAncho).length === 0) html += `<tr style="height: 18pt;"><td colspan="2" style="text-align: center; color: #94a3b8; font-size: 8pt; font-family: Arial, sans-serif; border: 1px solid #e2e8f0; vertical-align: middle;">Sin datos</td></tr>`;
                            Object.entries(wallSummary.murosPorAncho).forEach(([ancho, area]) => { 
                               html += `<tr style="text-align: center; background-color: #ffffff; height: 18pt;">
                                  <td style="font-weight: bold; color: #00b050; font-size: 8pt; font-family: Arial, sans-serif; border: none; border-bottom: 1px solid #e2e8f0; border-left: 1px solid #e2e8f0; vertical-align: middle; text-align: center;">Muro ${parseFloat(ancho).toFixed(2)}m</td>
                                  <td style="font-weight: bold; color: #00b050; font-size: 9pt; font-family: Arial, sans-serif; border: none; border-bottom: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0; vertical-align: middle; mso-number-format:'0\.00'; text-align: center;">${area.toFixed(2)}</td>
                               </tr>`; 
                            });
                   html += `</tbody></table></td>
                      <td style="vertical-align: top; border: none; padding-right: 20px;">
                         <table style="border-collapse: collapse; text-align: center; width: 250px;">
                            <thead><tr style="height: 20pt;"><th colspan="2" style="background-color: #7030a0; color: white; font-size: 9pt; font-family: Arial, sans-serif; padding: 8px; border: 1px solid #7030a0; vertical-align: middle; text-transform: uppercase;">DESGLOSE APLANADOS</th></tr></thead>
                            <tbody>`;
                            if(Object.keys(wallSummary.aplanados).length === 0) html += `<tr style="height: 18pt;"><td colspan="2" style="text-align: center; color: #94a3b8; font-size: 8pt; font-family: Arial, sans-serif; border: 1px solid #e2e8f0; vertical-align: middle;">Sin datos</td></tr>`;
                            Object.entries(wallSummary.aplanados).forEach(([tipo, area]) => { 
                               html += `<tr style="text-align: center; background-color: #ffffff; height: 18pt;">
                                  <td style="font-weight: bold; color: #7030a0; text-transform: uppercase; font-size: 8pt; font-family: Arial, sans-serif; border: none; border-bottom: 1px solid #e2e8f0; border-left: 1px solid #e2e8f0; vertical-align: middle; text-align: center;">${tipo}</td>
                                  <td style="font-weight: bold; color: #7030a0; font-size: 9pt; font-family: Arial, sans-serif; border: none; border-bottom: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0; vertical-align: middle; mso-number-format:'0\.00'; text-align: center;">${area.toFixed(2)}</td>
                               </tr>`; 
                            });
                   html += `</tbody></table></td>
                      <td style="vertical-align: top; border: none;">
                         <table style="border-collapse: collapse; text-align: center; width: 250px;">
                            <thead><tr style="height: 20pt;"><th colspan="2" style="background-color: #d200d2; color: white; font-size: 9pt; font-family: Arial, sans-serif; padding: 8px; border: 1px solid #d200d2; vertical-align: middle; text-transform: uppercase;">DESGLOSE RECUBRIMIENTOS</th></tr></thead>
                            <tbody>`;
                            if(Object.keys(wallSummary.recubrimientos).length === 0) html += `<tr style="height: 18pt;"><td colspan="2" style="text-align: center; color: #94a3b8; font-size: 8pt; font-family: Arial, sans-serif; border: 1px solid #e2e8f0; vertical-align: middle;">Sin datos</td></tr>`;
                            Object.entries(wallSummary.recubrimientos).forEach(([tipo, area]) => { 
                               html += `<tr style="text-align: center; background-color: #ffffff; height: 18pt;">
                                  <td style="font-weight: bold; color: #d200d2; text-transform: uppercase; font-size: 8pt; font-family: Arial, sans-serif; border: none; border-bottom: 1px solid #e2e8f0; border-left: 1px solid #e2e8f0; vertical-align: middle; text-align: center;">${tipo}</td>
                                  <td style="font-weight: bold; color: #d200d2; font-size: 9pt; font-family: Arial, sans-serif; border: none; border-bottom: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0; vertical-align: middle; mso-number-format:'0\.00'; text-align: center;">${area.toFixed(2)}</td>
                               </tr>`; 
                            });
                   html += `</tbody></table></td></tr></table>`;

                   exportFormattedExcel(html, `Matriz_Muros_${levelNameExport}`);
               }} className="px-6 py-2 bg-emerald-600 text-white font-black rounded-lg hover:bg-emerald-500 text-[10px] uppercase tracking-wider flex items-center gap-2 ml-auto"><FileDown size={14}/> Exportar Excel</button>
             </div>
             <div className="bg-slate-50 p-6 border-t border-slate-200 shrink-0">
                <h3 className="text-sm font-black uppercase tracking-widest text-slate-800 mb-6 flex items-center gap-2"><LayoutDashboard size={18} className="text-indigo-600" /> Resumen para Generadores</h3>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                  <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-100">
                     <span className="text-[10px] font-black text-emerald-600 uppercase tracking-widest block mb-3 border-b border-slate-100 pb-2">Área Neta por Espesor</span>
                     {Object.keys(wallSummary.murosPorAncho).length === 0 && <span className="text-xs font-bold text-slate-300">Sin capturas</span>}
                     {Object.entries(wallSummary.murosPorAncho).map(([ancho, area]) => ( <div key={ancho} className="flex justify-between items-center border-b border-slate-50 last:border-0 py-2"><span className="truncate w-2/3 uppercase text-slate-700 text-2xl font-black">Muro {ancho}m</span><div className="flex items-center gap-2"><span className="text-emerald-700 font-black text-xl">{area.toFixed(2)} m2</span><button onClick={() => handleCopyWalls('TOTAL_NETO', ancho)} className="p-1.5 bg-slate-100 text-slate-500 rounded-lg hover:bg-emerald-100 hover:text-emerald-700 transition-colors shadow-sm ml-1" title={`Copiar muros de ${ancho}m`}><Clipboard size={14}/></button></div></div> ))}
                  </div>
                  <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-100">
                     <span className="text-[10px] font-black text-purple-600 uppercase tracking-widest block mb-3 border-b border-slate-100 pb-2">Desglose Aplanados</span>
                     {Object.keys(wallSummary.aplanados).length === 0 && <span className="text-xs font-bold text-slate-300">Sin capturas</span>}
                     {Object.entries(wallSummary.aplanados).map(([tipo, area]) => ( <div key={tipo} className="flex justify-between items-center border-b border-slate-50 last:border-0 py-2"><span className="truncate w-2/3 uppercase text-slate-700 text-sm font-black">{tipo}</span><div className="flex items-center gap-2"><span className="text-purple-700 font-black text-lg">{area.toFixed(2)} m2</span><button onClick={() => handleCopyWalls('APLANADO', tipo)} className="p-1.5 bg-slate-100 text-slate-500 rounded-lg hover:bg-purple-100 hover:text-purple-700 transition-colors shadow-sm ml-1" title={`Copiar detalles de muros con ${tipo}`}><Clipboard size={14}/></button></div></div> ))}
                  </div>
                  <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-100">
                     <span className="text-[10px] font-black text-fuchsia-600 uppercase tracking-widest block mb-3 border-b border-slate-100 pb-2">Desglose Recubrimientos</span>
                     {Object.keys(wallSummary.recubrimientos).length === 0 && <span className="text-xs font-bold text-slate-300">Sin capturas</span>}
                     {Object.entries(wallSummary.recubrimientos).map(([tipo, area]) => ( <div key={tipo} className="flex justify-between items-center border-b border-slate-50 last:border-0 py-2"><span className="truncate w-2/3 uppercase text-slate-700 text-sm font-black">{tipo}</span><div className="flex items-center gap-2"><span className="text-fuchsia-700 font-black text-lg">{area.toFixed(2)} m2</span><button onClick={() => handleCopyWalls('RECUBRIMIENTO', tipo)} className="p-1.5 bg-slate-100 text-slate-500 rounded-lg hover:bg-fuchsia-100 hover:text-fuchsia-700 transition-colors shadow-sm ml-1" title={`Copiar detalles de muros con ${tipo}`}><Clipboard size={14}/></button></div></div> ))}
                  </div>
                </div>
             </div>
           </div>
        ) : showSettings ? (
          <div className="bg-white rounded-3xl p-10 shadow-xl border border-gray-100 animate-in fade-in zoom-in-95 duration-300">
            <div className="flex items-center justify-between mb-10">
              <h2 className="text-2xl font-black underline underline-offset-8 decoration-blue-500 decoration-4">
                Configuración del Proyecto
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
                <label className="block text-[10px] font-black text-blue-400 mb-3 uppercase tracking-[0.2em]">
                  Gestión de Niveles
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
                  + Agregar Nivel
                </button>
              </div>

              <div className="bg-blue-50/50 p-8 rounded-[2.5rem] border h-fit">
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
                {activePartidaId && (
                  <div className="mt-6 pt-6 border-t border-blue-200/50">
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
          <div className="overflow-x-auto bg-white rounded-lg shadow border border-gray-200 flex-1">
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
                        className="w-full h-full p-1.5 bg-transparent text-left outline-none text-blue-700 font-black min-w-0 cursor-pointer hover:bg-blue-50/50 text-[16px]"
                      />
                    </td>
                    <td className="p-0 border text-center text-[10px]">
                      <DebouncedCell
                        value={c.clave}
                        onChange={(v) => updateConceptoField(c.id, "clave", v)}
                        className="w-full h-full p-1.5 bg-transparent text-center outline-none min-w-0"
                      />
                    </td>
                    <td className="p-0 border text-center text-[10px]">
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
                    <td className="p-0 border text-xs md:text-sm font-bold">
                      <DebouncedCell
                        isTextArea
                        rows={2}
                        value={c.descripcion}
                        onChange={(v) =>
                          updateConceptoField(c.id, "descripcion", v)
                        }
                        className="w-full h-full p-1 bg-transparent outline-none resize-none min-w-0 leading-tight uppercase"
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
                          className={`p-1 md:p-1.5 border text-center cursor-pointer font-black text-[11px] md:text-xs leading-tight ${vol > 0 ? "bg-emerald-50 text-emerald-800" : "text-slate-200 hover:bg-blue-50"}`}
                        >
                          {vol.toFixed(2)}
                        </td>
                      );
                    })}
                    <td className="p-1 md:p-1.5 border text-center font-black bg-gray-100 text-[11px] md:text-xs leading-tight text-slate-800">
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
                onClick={() => {
                  let html = `<table style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">
                      <colgroup>
                         <col width="120" />
                         <col width="100" />
                         <col width="100" />
                         <col width="250" />
                         <col width="450" />
                         <col width="80" />
                         ${niveles.map(() => '<col width="100" />').join("")}
                         <col width="100" />
                      </colgroup>
                      <thead>
                      <tr style="height: 25pt;">
                         <td colspan="${6 + niveles.length + 1}" style="background-color: #0b1a30; color: #ffffff; font-size: 10pt; font-weight: bold; text-align: center; vertical-align: middle; padding: 8px; border: none; text-transform: uppercase;">
                            CATÁLOGO GENERAL - PARTIDA: ${activePartida.nombre}
                         </td>
                      </tr>
                      <tr><td colspan="${6 + niveles.length + 1}" style="height: 10px; border: none; background-color: #ffffff;"></td></tr>
                      <tr style="height: 20pt;">
                         <td style="background-color: #1f3864; color: #ffffff; text-align: center; vertical-align: middle; font-weight: bold; font-size: 10pt; border: 1px solid #ffffff;">CÓDIGO</td>
                         <td style="background-color: #1f3864; color: #ffffff; text-align: center; vertical-align: middle; font-weight: bold; font-size: 10pt; border: 1px solid #ffffff;">CLAVE</td>
                         <td style="background-color: #1f3864; color: #ffffff; text-align: center; vertical-align: middle; font-weight: bold; font-size: 10pt; border: 1px solid #ffffff;">CC</td>
                         <td style="background-color: #1f3864; color: #ffffff; text-align: center; vertical-align: middle; font-weight: bold; font-size: 10pt; border: 1px solid #ffffff;">JUSTIFICACIÓN</td>
                         <td style="background-color: #1f3864; color: #ffffff; text-align: center; vertical-align: middle; font-weight: bold; font-size: 10pt; border: 1px solid #ffffff;">DESCRIPCIÓN</td>
                         <td style="background-color: #1f3864; color: #ffffff; text-align: center; vertical-align: middle; font-weight: bold; font-size: 10pt; border: 1px solid #ffffff;">UNID.</td>
                         ${niveles.map((n) => `<td style="background-color: #1f3864; color: #ffffff; text-align: center; vertical-align: middle; font-weight: bold; font-size: 10pt; border: 1px solid #ffffff;">${n.nombre.toUpperCase()}</td>`).join("")}
                         <td style="background-color: #000000; color: #ffffff; text-align: center; vertical-align: middle; font-weight: bold; font-size: 10pt; border: 1px solid #ffffff;">TOTAL</td>
                      </tr>
                  </thead><tbody>`;

                  conceptos.forEach((c) => {
                    const just = (c.justificacion || "").replace(/\n/g, " ");
                    const desc = (c.descripcion || "").replace(/\n/g, " ");
                    const total = getTotalConcepto(c.id);

                    html += `<tr style="height: 25pt; background-color: #f4f6f9;">
                         <td style="font-weight: bold; color: #002060; text-align: center; vertical-align: middle; padding: 6px; border: 1px solid #b4c6e7; font-size: 12pt;">${c.id}</td>
                         <td style="font-weight: bold; color: #404040; text-align: center; vertical-align: middle; padding: 6px; border: 1px solid #b4c6e7; font-size: 10pt;">${c.clave || ""}</td>
                         <td style="font-weight: bold; color: #404040; text-align: center; vertical-align: middle; padding: 6px; border: 1px solid #b4c6e7; font-size: 10pt;">${c.cc || ""}</td>
                         <td style="color: #404040; text-align: center; vertical-align: middle; padding: 6px; border: 1px solid #b4c6e7; font-size: 10pt;">${just}</td>
                         <td style="font-weight: bold; color: #404040; text-align: center; vertical-align: middle; padding: 6px; border: 1px solid #b4c6e7; font-size: 10pt;">${desc}</td>
                         <td style="color: #0070c0; font-weight: bold; text-align: center; vertical-align: middle; padding: 6px; border: 1px solid #b4c6e7; font-size: 12pt;">${c.unidad || ""}</td>
                         ${niveles
                           .map((n) => {
                             const vol = getVolumenNivel(c.id, n.id);
                             if (vol > 0) {
                               return `<td style="font-weight: bold; background-color: #e2efda; color: #00b050; text-align: right; vertical-align: middle; padding: 6px; border: 1px solid #b4c6e7; font-size: 10pt; mso-number-format:'0\\.00';">${vol.toFixed(2)}</td>`;
                             } else {
                               return `<td style="color: #d9d9d9; text-align: right; vertical-align: middle; padding: 6px; border: 1px solid #b4c6e7; font-size: 10pt;">-</td>`;
                             }
                           })
                           .join("")}
                         <td style="font-weight: bold; color: #000000; text-align: right; vertical-align: middle; padding: 6px; border: 1px solid #b4c6e7; font-size: 10pt; mso-number-format:'0\\.00';">${total > 0 ? total.toFixed(2) : "-"}</td>
                      </tr>`;
                  });
                  html += "</tbody></table>";
                  exportFormattedExcel(
                    html,
                    `Matriz_Conceptos_${activePartida.nombre}`,
                  );
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
    </div>
  );
}

const Dashboard = ({ projects, onCreate, onSelect, onDelete, onCatMaestro }) => {
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
                  <Trash2 size={16} className="group-hover/del:scale-110 transition-transform" />
                  <span className="text-[10px] font-black uppercase tracking-widest">Borrar</span>
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

  const [projects, setProjects, loadedProjects, savingProjects] = usePersistentState(
    "xdifica_projects_list",
    [],
    user,
  );
  const [catalogoConceptos, setCatalogoConceptos, loadedCatalogo, savingCatalogo] =
    usePersistentState("xdifica_global_catalogo_v3", [], user);
  const [activeProjectId, setActiveProjectId] = useState(null);
  const [showGlobalCatalogo, setShowGlobalCatalogo] = useState(false);

  useEffect(() => {
    const handleBeforeUnload = (e) => {
      if (savingProjects || savingCatalogo) {
        e.preventDefault();
        e.returnValue = "";
      }
    };
    window.addEventListener("beforeunload", handleBeforeUnload);
    return () => window.removeEventListener("beforeunload", handleBeforeUnload);
  }, [savingProjects, savingCatalogo]);

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

  const handleDeleteProject = useCallback(async (id) => {
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
  }, [user, db, setProjects]);

  const handleUpdateMetadata = (id, name, location) => {
    setProjects((prev) =>
      prev.map((p) => (p.id === id ? { ...p, name, location } : p)),
    );
  };

  if (!authInit) return <LoadingScreen text="Conectando con la nube..." />;
  if (!loadedProjects || !loadedCatalogo) return <LoadingScreen text="Sincronizando obras..." />;

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
}) {
  const [catColWidths, setCatColWidths] = usePersistentState(
    "xdifica_catColWidths_v2",
    {
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
    resizingRef.current = { colId, startX: e.pageX, startWidth: catColWidths[colId] || 100 };
    document.addEventListener("mousemove", handleMouseMove);
    document.addEventListener("mouseup", stopResizing);
  };
  const handleMouseMove = useCallback((e) => {
    if (!resizingRef.current) return;
    const { colId, startX, startWidth } = resizingRef.current;
    const newWidth = Math.max(30, startWidth + (e.pageX - startX));
    setCatColWidths((p) => ({ ...p, [colId]: newWidth }));
  }, [setCatColWidths]);
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
    (oldId, newId) => {
      if (oldId === newId || !newId.trim()) return;
      setCatalogoConceptos((prev) =>
        prev.some((c) => c.id === newId)
          ? prev
          : prev.map((c) => (c.id === oldId ? { ...c, id: newId } : c)),
      );
    },
    [setCatalogoConceptos],
  );

  return (
    <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[500] flex items-center justify-center p-4 md:p-8 animate-in fade-in duration-200">
      <div className="bg-white rounded-3xl shadow-xl w-full max-w-6xl flex flex-col h-[80vh] overflow-hidden">
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
        <div className="flex-1 overflow-x-auto bg-slate-50 p-4 md:p-6">
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden min-w-[900px]">
            <table className="text-sm text-left border-collapse table-fixed w-full">
              <thead className="bg-slate-800 text-white uppercase tracking-wider select-none font-black sticky top-0 z-30">
                <tr>
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
                {catalogoConceptos.length === 0 && (
                  <tr>
                    <td
                      colSpan={7}
                      className="p-12 text-center text-slate-400 font-bold uppercase tracking-widest"
                    >
                      <div className="flex flex-col items-center justify-center gap-3">
                        <AlertCircle size={32} className="opacity-30" />
                        <span>El catálogo está vacío. Agrega conceptos base.</span>
                      </div>
                    </td>
                  </tr>
                )}
                {catalogoConceptos.map((c, idx) => (
                  <tr
                    key={c.id}
                    className={`transition-colors ${idx % 2 === 0 ? "bg-white" : "bg-slate-50/50"} hover:bg-teal-50/30`}
                  >
                    <td className="p-0 border-r border-slate-200">
                      <DebouncedCell
                        value={c.id}
                        onChange={(v) => updateCatalogoId(c.id, v)}
                        className="w-full h-full p-3 bg-transparent text-left outline-none font-black text-teal-800 min-w-0 text-[16px]"
                        placeholder="Ej. PRE-01"
                      />
                    </td>
                    <td className="p-0 border-r border-slate-200">
                      <DebouncedCell
                        value={c.clave}
                        onChange={(v) => updateCatalogoField(c.id, "clave", v)}
                        className="w-full h-full p-3 bg-transparent text-center outline-none min-w-0 font-bold text-slate-600"
                      />
                    </td>
                    <td className="p-0 border-r border-slate-200">
                      <DebouncedCell
                        value={c.cc}
                        onChange={(v) => updateCatalogoField(c.id, "cc", v)}
                        className="w-full h-full p-3 bg-transparent text-center outline-none min-w-0 font-bold text-slate-600"
                      />
                    </td>
                    <td className="p-0 border-r border-slate-200">
                      <DebouncedCell
                        isTextArea
                        rows={2}
                        value={c.justificacion}
                        onChange={(v) => updateCatalogoField(c.id, "justificacion", v)}
                        className="w-full h-full p-2 bg-transparent outline-none resize-none min-w-0 italic text-slate-500"
                      />
                    </td>
                    <td className="p-0 border-r border-slate-200">
                      <DebouncedCell
                        isTextArea
                        rows={2}
                        value={c.descripcion}
                        onChange={(v) => updateCatalogoField(c.id, "descripcion", v)}
                        className="w-full h-full p-2 bg-transparent outline-none resize-none min-w-0 text-slate-700 font-bold uppercase"
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
                        onClick={() => setCatalogoConceptos((prev) => prev.filter((x) => x.id !== c.id))}
                        className="text-slate-300 hover:text-red-500 transition-colors p-2"
                      >
                        <Trash2 size={16} />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
        <div className="p-4 md:p-6 bg-white border-t border-slate-100 flex gap-4 shrink-0">
          <button
            onClick={() => {
              const newId = `CAT-${Date.now().toString(36).toUpperCase()}`;
              setCatalogoConceptos((prev) => [
                ...prev,
                { id: newId, clave: "", cc: "", justificacion: "", descripcion: "", unidad: "m2" },
              ]);
            }}
            className="px-6 py-3 bg-teal-700 text-white rounded-xl text-xs font-black shadow-md transition-transform active:scale-95 uppercase tracking-wider hover:bg-teal-600 flex items-center gap-2"
          >
            <Plus size={16} /> Agregar Concepto
          </button>
        </div>
      </div>
    </div>
  );
}
