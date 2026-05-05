const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf8');

const newSettingsBlock = `) : showSettings ? (
          <div className="bg-white rounded-3xl p-10 shadow-xl border border-gray-100 animate-in fade-in zoom-in-95 duration-300">
            <div className="flex items-center justify-between mb-10">
              <h2 className="text-2xl font-black underline underline-offset-8 decoration-blue-500 decoration-4">Configuración del Proyecto</h2>
              <button onClick={() => setShowSettings(false)} className="p-2 hover:bg-slate-100 rounded-full transition-colors"><X size={24} /></button>
            </div>
            
            <div className="grid md:grid-cols-2 gap-12">
               <div className="space-y-8">
                 <div className="space-y-4">
                   <label className="block text-[10px] font-black text-blue-400 mb-3 uppercase tracking-[0.2em]">Gestión de Niveles</label>
                   {niveles.map((n, i) => ( <div key={n.id} className="flex gap-3 bg-slate-50 p-2 rounded-xl border items-center shadow-sm"><DebouncedCell className="flex-1 bg-transparent p-2 font-black text-slate-700" value={n.nombre || ''} onChange={v => { const x = [...niveles]; x[i].nombre = v; setNiveles(x); }} /><button onClick={() => setNiveles(niveles.filter(l => l.id !== n.id))} className="text-red-300 hover:text-red-500 p-2"><Trash2 size={16}/></button></div> ))}
                   <button onClick={() => setNiveles([...niveles, { id: \`n\${Date.now()}\`, nombre: \`N\${niveles.length+1}\` }])} className="w-full py-3 mt-4 border-2 border-dashed rounded-xl font-black text-blue-500 border-blue-200 hover:bg-blue-50 transition-all">+ Agregar Nivel</button>
                 </div>
                 
                 <div className="space-y-4">
                   <label className="block text-[10px] font-black text-blue-400 mb-3 uppercase tracking-[0.2em]">Configuración de Aceros (Pesos Kg/ml)</label>
                   <div className="space-y-2 max-h-60 overflow-y-auto pr-2">
                      {configuracion?.aceros?.map((acero, i) => (
                        <div key={i} className="flex items-center gap-4 bg-slate-50 p-4 rounded-xl border border-slate-100">
                           <div className="w-12 h-12 bg-white rounded-xl shadow-sm flex items-center justify-center font-black text-slate-400 border border-slate-100">#{acero.numero}</div>
                           <div className="flex-1">
                              <input type="number" step="0.01" value={acero.kg_ml} onChange={(e) => { const newConf = {...configuracion}; newConf.aceros[i].kg_ml = parseFloat(e.target.value) || 0; setConfiguracion(newConf); }} className="w-full bg-white border border-slate-200 rounded-lg p-2 font-bold text-slate-700 outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100 transition-all"/>
                           </div>
                        </div>
                      ))}
                   </div>
                 </div>
                 
                 <div className="space-y-4">
                    <label className="block text-[10px] font-black text-blue-400 mb-3 uppercase tracking-[0.2em]">Recubrimientos Base</label>
                    <div className="flex items-center gap-4 bg-slate-50 p-4 rounded-xl border border-slate-100"><div className="flex-1"><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest block mb-1">Cimentación (m)</label><input type="number" step="0.01" value={configuracion?.recubrimientoCimentacion} onChange={(e) => setConfiguracion({...configuracion, recubrimientoCimentacion: parseFloat(e.target.value)||0})} className="w-full bg-white border border-slate-200 rounded-lg p-2 font-bold text-slate-700 outline-none focus:border-blue-500 transition-all"/></div></div>
                    <div className="flex items-center gap-4 bg-slate-50 p-4 rounded-xl border border-slate-100"><div className="flex-1"><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest block mb-1">Superestructura (m)</label><input type="number" step="0.01" value={configuracion?.recubrimientoEstructura} onChange={(e) => setConfiguracion({...configuracion, recubrimientoEstructura: parseFloat(e.target.value)||0})} className="w-full bg-white border border-slate-200 rounded-lg p-2 font-bold text-slate-700 outline-none focus:border-blue-500 transition-all"/></div></div>
                 </div>
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
        ) : (`;

const oldRegex = /\) \: showSettings \? \((.*?)\) \: \(/s;

if (oldRegex.test(code)) {
    code = code.replace(oldRegex, newSettingsBlock);
    fs.writeFileSync('src/App.tsx', code);
    console.log('Successfully replaced Settings block');
} else {
    console.log('Failed to match legacy Settings block.');
}
