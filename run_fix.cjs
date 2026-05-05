const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf8');

// 1. Remove configuracion state and fix isFullyLoaded
code = code.replace(
  /const \[configuracion.*?} \], user\);/s,
  ''
);

code = code.replace(
  /const isFullyLoaded = l1 && l2 && l3 && l4 && l5 && l6 && l7 && l8 && l9 && l10 && l11;/,
  'const isFullyLoaded = l1 && l2 && l3 && l4 && l5 && l6 && l7 && l8 && l9 && l10;'
);

// 2. Replace the showSettings render block with the original one
const settingsRegex = /\) : showSettings \? \(\n\s*<div className="bg-white rounded-3xl p-8 shadow-xl.*?\) : activePartidaId \? \(/s;
const originalSettings = `) : showSettings ? (
          <div className="bg-white rounded-3xl p-10 shadow-xl border border-gray-100">
            <h2 className="text-2xl font-black mb-10 underline underline-offset-8 decoration-blue-500 decoration-4">Configuración del Proyecto</h2>
            <div className="grid md:grid-cols-2 gap-12">
               <div className="space-y-4">
                 <label className="block text-[10px] font-black text-blue-400 mb-3 uppercase tracking-[0.2em]">Gestión de Niveles</label>
                 {niveles.map((n, i) => ( <div key={n.id} className="flex gap-3 bg-slate-50 p-2 rounded-xl border items-center shadow-sm"><DebouncedCell className="flex-1 bg-transparent p-2 font-black text-slate-700" value={n.nombre || ''} onChange={v => { const x = [...niveles]; x[i].nombre = v; setNiveles(x); }} /><button onClick={() => setNiveles(niveles.filter(l => l.id !== n.id))} className="text-red-300 hover:text-red-500 p-2"><Trash2 size={16}/></button></div> ))}
                 <button onClick={() => setNiveles([...niveles, { id: \`n\${Date.now()}\`, nombre: \`N\${niveles.length+1}\` }])} className="w-full py-3 mt-4 border-2 border-dashed rounded-xl font-black text-blue-500 border-blue-200 hover:bg-blue-50 transition-all">+ Agregar Nivel</button>
               </div>
               <div className="bg-blue-50/50 p-8 rounded-[2.5rem] border h-fit">
                 <div className="space-y-6">
                   <div><label className="block text-[10px] font-black text-blue-400 mb-2 uppercase tracking-[0.2em]">Nombre del Proyecto (Obra)</label><DebouncedCell className="w-full p-4 bg-white border border-blue-100 rounded-xl text-lg font-black shadow-sm outline-none" value={obraInfo.nombre || ''} onChange={v => { const newInfo = {...obraInfo, nombre: v}; setObraInfo(newInfo); onUpdateMetadata(projectId, newInfo.nombre, newInfo.ubicacion); }} /></div>
                   <div><label className="block text-[10px] font-black text-blue-400 mb-2 uppercase tracking-[0.2em]">Ubicación</label><DebouncedCell className="w-full p-4 bg-white border border-blue-100 rounded-xl text-sm font-bold shadow-sm outline-none" value={obraInfo.ubicacion || ''} onChange={v => { const newInfo = {...obraInfo, ubicacion: v}; setObraInfo(newInfo); onUpdateMetadata(projectId, newInfo.nombre, newInfo.ubicacion); }} /></div>
                 </div>
                 {activePartidaId && ( <div className="mt-6 pt-6 border-t border-blue-200/50"><label className="block text-[10px] font-black text-blue-400 mb-3 uppercase tracking-[0.2em]">Nombre Partida Actual</label><DebouncedCell className="w-full p-5 bg-white border border-blue-100 rounded-2xl text-2xl font-black shadow-sm outline-none text-slate-800" value={activePartida?.nombre || ''} onChange={(v) => setPartidas(prev => prev.map(p => p.id === activePartidaId ? { ...p, nombre: (v || '').toUpperCase() } : p))} /></div> )}
               </div>
            </div>
          </div>
        ) : activePartidaId ? (`;

code = code.replace(settingsRegex, originalSettings);

// 3. Replace everything from function ObraManager... to the end with the Dashboard + App + old Dashboard code
const endRegex = /function ObraManager\(\{(.*?)\}\n/s;
const indexOfEnd = code.match(endRegex)?.index || code.indexOf("function ObraManager(");
if (indexOfEnd > -1) {
  code = code.substring(0, indexOfEnd);
}

const originalEnd = `
const Dashboard = ({ projects, onCreate, onSelect, onDelete }) => {
  return (
    <div className="min-h-screen bg-[#fcfdfe] font-sans text-slate-900 p-8 flex flex-col items-center">
       <div className="max-w-7xl w-full animate-in fade-in duration-500">
          <div className="flex justify-between items-end mb-12">
             <div className="flex flex-col gap-4">
                <div className="flex items-center gap-4"><div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center text-white font-black text-sm">X</div><div className="h-6 w-px bg-slate-200"></div><span className="text-blue-600 font-black uppercase tracking-[0.4em] text-[8px]">Asset Management</span></div>
                <div>
                   <h1 className="text-4xl font-black tracking-tighter mb-2">Gestor de Obras</h1>
                   <p className="text-slate-400 font-bold uppercase tracking-widest text-xs">Administración Global de Proyectos</p>
                </div>
             </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
             <div onClick={onCreate} className="bg-slate-50/50 rounded-[2rem] p-8 border-2 border-dashed border-slate-300 hover:border-blue-400 hover:bg-blue-50 transition-all flex flex-col items-center justify-center cursor-pointer text-slate-400 hover:text-blue-600 min-h-[250px] group">
                <div className="bg-white w-16 h-16 rounded-2xl flex items-center justify-center mb-6 shadow-sm group-hover:scale-110 transition-transform"><FolderPlus size={28} /></div>
                <span className="font-black text-sm uppercase tracking-widest">Nueva Obra</span>
                <span className="text-[10px] font-bold mt-2 opacity-60 uppercase text-center">Crear espacio de trabajo</span>
             </div>
             {projects.map(p => (
                <div key={p.id} onClick={() => onSelect(p.id)} className="bg-white rounded-[2rem] p-8 border border-slate-200 shadow-sm hover:shadow-xl hover:border-blue-300 transition-all group flex flex-col cursor-pointer relative overflow-hidden h-[250px]">
                   <div className="absolute top-0 right-0 p-8 opacity-5 group-hover:opacity-10 transition-opacity transform group-hover:scale-110 group-hover:-rotate-12 duration-500"><Building2 size={120} /></div>
                   <div className="bg-blue-50 w-16 h-16 rounded-2xl flex items-center justify-center text-blue-600 mb-6 group-hover:scale-110 transition-transform shadow-sm"><Building2 size={28} /></div>
                   <h3 className="text-xl font-black text-slate-800 mb-2 uppercase leading-tight z-10">{p.name || 'Obra Sin Nombre'}</h3>
                   <p className="text-slate-400 text-xs font-bold uppercase tracking-widest mb-10 z-10">{p.location || 'Ubicación no definida'}</p>
                   <div className="mt-auto flex justify-between items-center border-t border-slate-100 pt-6 z-10">
                      <span className="text-blue-600 font-black text-[10px] uppercase tracking-widest flex items-center gap-2">Abrir Obra <ChevronRight size={16}/></span>
                      <button onClick={(e) => { e.stopPropagation(); onDelete(p.id); }} className="text-slate-300 hover:text-red-500 p-2 rounded-full hover:bg-red-50 transition-colors"><Trash2 size={18} /></button>
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
      if (!auth) { setAuthInit(true); return; }
      const initAuth = async () => {
         try {
            if (typeof __initial_auth_token !== 'undefined' && __initial_auth_token) {
               await signInWithCustomToken(auth, __initial_auth_token);
            } else {
               await signInAnonymously(auth);
            }
         } catch (e) { console.error("Error de autenticación", e); }
      };
      initAuth();
      const unsubscribe = onAuthStateChanged(auth, (u) => {
         setUser(u);
         setAuthInit(true);
      });
      return () => unsubscribe();
   }, []);

   const [projects, setProjects, loadedProjects] = usePersistentState('xdifica_projects_list', [], user);
   const [activeProjectId, setActiveProjectId] = useState(null);

   useEffect(() => {
      if (!loadedProjects || !authInit || !user) return;
      const oldData = localStorage.getItem('xdifica_partidas');
      if (oldData && projects.length === 0) {
          const migId = \`PROJ-\${Date.now()}\`;
          let projectName = 'Proyecto Migrado Local';
          let projectLocation = 'Automático';
          const oldInfo = localStorage.getItem('xdifica_obraInfo');
          if(oldInfo) {
             try {
               const parsed = JSON.parse(oldInfo);
               if (parsed.nombre) projectName = parsed.nombre;
               if (parsed.ubicacion) projectLocation = parsed.ubicacion;
             } catch(e) {}
          }
          setProjects([{ id: migId, name: projectName, location: projectLocation }]);
          const payload = { partidas: JSON.parse(oldData) };
          if(oldInfo) payload.obraInfo = JSON.parse(oldInfo);
          const oldNiveles = localStorage.getItem('xdifica_niveles');
          if(oldNiveles) payload.niveles = JSON.parse(oldNiveles);
          if (user && db) {
              const docRef = doc(db, 'artifacts', appId, 'users', user.uid, 'projects', migId);
              setDoc(docRef, payload, { merge: true });
          } else {
              localStorage.setItem(\`xdifica_proj_\${migId}_partidas\`, JSON.stringify(payload.partidas));
              if (payload.obraInfo) localStorage.setItem(\`xdifica_proj_\${migId}_obraInfo\`, JSON.stringify(payload.obraInfo));
              if (payload.niveles) localStorage.setItem(\`xdifica_proj_\${migId}_niveles\`, JSON.stringify(payload.niveles));
          }
          localStorage.removeItem('xdifica_partidas');
          localStorage.removeItem('xdifica_obraInfo');
          localStorage.removeItem('xdifica_niveles');
      }
   }, [projects.length, setProjects, loadedProjects, authInit, user]);

   const handleCreateProject = () => {
      const newId = \`PROJ-\${Date.now().toString(36).toUpperCase()}\`;
      setProjects([{ id: newId, name: 'NUEVA OBRA', location: 'Sin definir' }, ...projects]);
      setActiveProjectId(newId);
   };

   const handleDeleteProject = async (id) => {
      if (window.confirm("¿Estás seguro de que deseas eliminar esta obra y todos sus datos?")) {
          setProjects(prev => prev.filter(p => p.id !== id));
          if (user && db) {
             const docRef = doc(db, 'artifacts', appId, 'users', user.uid, 'projects', id);
             await deleteDoc(docRef);
          } else {
             ['obraInfo', 'niveles', 'partidas', 'catalogoConceptos'].forEach(key => localStorage.removeItem(\`xdifica_proj_\${id}_\${key}\`));
          }
      }
   };

   const handleUpdateMetadata = (id, name, location) => {
      setProjects(prev => prev.map(p => p.id === id ? { ...p, name, location } : p));
   };

   if (!authInit) return <LoadingScreen text="Conectando con la nube..." />;
   if (!loadedProjects) return <LoadingScreen text="Sincronizando obras..." />;

   if (activeProjectId) {
      return <ProjectWorkspace projectId={activeProjectId} onBack={() => setActiveProjectId(null)} onUpdateMetadata={handleUpdateMetadata} user={user} />;
   }

   return <Dashboard projects={projects} onCreate={handleCreateProject} onSelect={setActiveProjectId} onDelete={handleDeleteProject} />;
}
`;

code += originalEnd;
fs.writeFileSync('src/App.tsx', code);
console.log('Restored App.tsx correctly.');
