const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf8');

const newWallBlock = `) : showWallGenerator ? (
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
                       <tr key={m.id} className={\`transition-colors \${selectedWallRows.includes(m.id) ? 'bg-indigo-50/50' : 'hover:bg-slate-50'}\`}>
                         <td className="p-1 text-center border-r border-slate-200"><button onClick={() => toggleWallRow(m.id)} className={\`\${selectedWallRows.includes(m.id) ? 'text-indigo-600' : 'text-slate-300 hover:text-indigo-500'} transition-colors\`}>{selectedWallRows.includes(m.id) ? <CheckSquare size={14} /> : <Square size={14} />}</button></td>
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
               <button onClick={() => updateActiveMuros(prev => { const lastRow = prev[prev.length - 1]; const nextClave = lastRow ? getNextClaveValue(lastRow.clave) : ''; return [...prev, { id: \`M-\${Date.now()}\`, eje: lastRow?.eje || '', clave: nextClave, largo: '', ancho: '', alto: '', huecos: [], castillos: [], tipoAplanadoC1: '', tipoAplanadoC2: '', tipoRecubrimientoC1: '', tipoRecubrimientoC2: '' }]; })} className="px-6 py-2 border-2 border-dashed border-indigo-200 text-indigo-600 font-black rounded-lg hover:bg-indigo-50 text-[10px] uppercase tracking-wider">+ Agregar Muro</button>
               <button onClick={handleCopySelectedWalls} className={\`px-6 py-2 rounded-lg font-black text-[10px] uppercase flex items-center gap-2 transition-colors \${selectedWallRows.length > 0 ? 'bg-indigo-600 text-white hover:bg-indigo-700 shadow-md' : 'bg-slate-100 hover:bg-slate-200 text-slate-600'}\`}><Copy size={14}/> {copyStatus === 'copied-wall-list' ? '¡Copiado!' : selectedWallRows.length > 0 ? \`Copiar (\${selectedWallRows.length})\` : 'Copiar Todas'}</button>
               <button onClick={handlePasteSelectedWalls} disabled={!copiedWallRowsList || copiedWallRowsList.length === 0} className={\`px-6 py-2 rounded-lg font-black text-[10px] uppercase flex items-center gap-2 transition-all \${copiedWallRowsList && copiedWallRowsList.length > 0 ? 'bg-emerald-100 hover:bg-emerald-200 text-emerald-700 shadow-sm cursor-pointer' : 'bg-slate-50 text-slate-300 cursor-not-allowed'}\`}><ClipboardPaste size={14}/> Pegar Datos</button>
               <button onClick={() => {
                   const levelNameExport = niveles.find(n=>n.id===activeLevelId)?.nombre||'TODOS LOS NIVELES';
                   const renderExVal = (v) => {
                       const num = parseFloat(v);
                       if (isNaN(num) || num === 0) return '-';
                       return num.toFixed(2);
                   };

                   let html = \`<table style="border-collapse: collapse; font-family: Arial, sans-serif;">
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
                            MATRIZ DE MUROS - NIVEL: \${levelNameExport}
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
                   </thead><tbody>\`;

                   muros.forEach((m, i) => {
                      const bruto = calcWallArea(m.largo, m.alto), dedH = getWallHuecosTotal(m.huecos), dedC = getWallCastillosTotal(m.castillos, m.alto), neto = Math.max(0, bruto - dedH - dedC);
                      html += \`<tr style="text-align: center; background-color: #ffffff; height: 18pt;">
                         <td style="border: none; background-color: #ffffff;"></td>
                         <td style="border: 1px solid #c7d2fe; color: #a6a6a6; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; text-align: center; mso-number-format:'0';">\${i+1}</td>
                         <td style="border: 1px solid #c7d2fe; color: #312eb5; font-weight: bold; text-transform: uppercase; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; text-align: center;">\${m.eje || ''}</td>
                         <td style="border: 1px solid #c7d2fe; color: #000000; font-weight: bold; text-transform: uppercase; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; text-align: center;">\${m.clave || ''}</td>
                         <td style="border: 1px solid #c7d2fe; color: #000066; font-weight: bold; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; mso-number-format:'0\\.00'; text-align: center;">\${renderExVal(m.largo)}</td>
                         <td style="border: 1px solid #c7d2fe; color: #000066; font-weight: bold; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; mso-number-format:'0\\.00'; text-align: center;">\${renderExVal(m.ancho)}</td>
                         <td style="border: 1px solid #c7d2fe; color: #000066; font-weight: bold; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; mso-number-format:'0\\.00'; text-align: center;">\${renderExVal(m.alto)}</td>
                         <td style="border: 1px solid #c7d2fe; color: #404040; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; mso-number-format:'0\\.00'; text-align: center;">\${renderExVal(bruto)}</td>
                         <td style="border: 1px solid #c7d2fe; color: #c00000; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; mso-number-format:'0\\.00'; text-align: center;">\${dedH > 0 ? dedH.toFixed(2) : '-'}</td>
                         <td style="border: 1px solid #c7d2fe; color: #e36c0a; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; mso-number-format:'0\\.00'; text-align: center;">\${dedC > 0 ? dedC.toFixed(2) : '-'}</td>
                         <td style="border: 1px solid #c7d2fe; color: #00b050; font-weight: bold; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; mso-number-format:'0\\.00'; text-align: center;">\${renderExVal(neto)}</td>
                         <td style="border: 1px solid #c7d2fe; color: #800080; font-weight: bold; text-transform: uppercase; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; text-align: center;">\${m.tipoAplanadoC1 || ''}</td>
                         <td style="border: 1px solid #c7d2fe; color: #800080; font-weight: bold; text-transform: uppercase; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; text-align: center;">\${m.tipoAplanadoC2 || ''}</td>
                         <td style="border: 1px solid #c7d2fe; color: #800080; font-weight: bold; text-transform: uppercase; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; text-align: center;">\${m.tipoRecubrimientoC1 || ''}</td>
                         <td style="border: 1px solid #c7d2fe; color: #800080; font-weight: bold; text-transform: uppercase; font-size: 11pt; font-family: Arial, sans-serif; vertical-align: middle; text-align: center;">\${m.tipoRecubrimientoC2 || ''}</td>
                      </tr>\`;
                   });
                   html += \`</tbody></table><br/><br/><br/>\`;

                   html += \`<table style="width: 100%; border: none; font-family: Arial, sans-serif;"><tr>
                      <td style="width: 40px; border: none;"></td>
                      <td style="vertical-align: top; border: none; padding-right: 20px;">
                         <table style="border-collapse: collapse; text-align: center; width: 250px;">
                            <thead><tr style="height: 20pt;"><th colspan="2" style="background-color: #00b050; color: white; font-size: 9pt; font-family: Arial, sans-serif; padding: 8px; border: 1px solid #00b050; vertical-align: middle; text-transform: uppercase;">ÁREA NETA POR ESPESOR</th></tr></thead>
                            <tbody>\`;
                            if(Object.keys(wallSummary.murosPorAncho).length === 0) html += \`<tr style="height: 18pt;"><td colspan="2" style="text-align: center; color: #94a3b8; font-size: 8pt; font-family: Arial, sans-serif; border: 1px solid #e2e8f0; vertical-align: middle;">Sin datos</td></tr>\`;
                            Object.entries(wallSummary.murosPorAncho).forEach(([ancho, area]) => { 
                               html += \`<tr style="text-align: center; background-color: #ffffff; height: 18pt;">
                                  <td style="font-weight: bold; color: #00b050; font-size: 8pt; font-family: Arial, sans-serif; border: none; border-bottom: 1px solid #e2e8f0; border-left: 1px solid #e2e8f0; vertical-align: middle; text-align: center;">Muro \${parseFloat(ancho).toFixed(2)}m</td>
                                  <td style="font-weight: bold; color: #00b050; font-size: 9pt; font-family: Arial, sans-serif; border: none; border-bottom: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0; vertical-align: middle; mso-number-format:'0\\.00'; text-align: center;">\${area.toFixed(2)}</td>
                               </tr>\`; 
                            });
                   html += \`</tbody></table></td>
                      <td style="vertical-align: top; border: none; padding-right: 20px;">
                         <table style="border-collapse: collapse; text-align: center; width: 250px;">
                            <thead><tr style="height: 20pt;"><th colspan="2" style="background-color: #7030a0; color: white; font-size: 9pt; font-family: Arial, sans-serif; padding: 8px; border: 1px solid #7030a0; vertical-align: middle; text-transform: uppercase;">DESGLOSE APLANADOS</th></tr></thead>
                            <tbody>\`;
                            if(Object.keys(wallSummary.aplanados).length === 0) html += \`<tr style="height: 18pt;"><td colspan="2" style="text-align: center; color: #94a3b8; font-size: 8pt; font-family: Arial, sans-serif; border: 1px solid #e2e8f0; vertical-align: middle;">Sin datos</td></tr>\`;
                            Object.entries(wallSummary.aplanados).forEach(([tipo, area]) => { 
                               html += \`<tr style="text-align: center; background-color: #ffffff; height: 18pt;">
                                  <td style="font-weight: bold; color: #7030a0; text-transform: uppercase; font-size: 8pt; font-family: Arial, sans-serif; border: none; border-bottom: 1px solid #e2e8f0; border-left: 1px solid #e2e8f0; vertical-align: middle; text-align: center;">\${tipo}</td>
                                  <td style="font-weight: bold; color: #7030a0; font-size: 9pt; font-family: Arial, sans-serif; border: none; border-bottom: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0; vertical-align: middle; mso-number-format:'0\\.00'; text-align: center;">\${area.toFixed(2)}</td>
                               </tr>\`; 
                            });
                   html += \`</tbody></table></td>
                      <td style="vertical-align: top; border: none;">
                         <table style="border-collapse: collapse; text-align: center; width: 250px;">
                            <thead><tr style="height: 20pt;"><th colspan="2" style="background-color: #d200d2; color: white; font-size: 9pt; font-family: Arial, sans-serif; padding: 8px; border: 1px solid #d200d2; vertical-align: middle; text-transform: uppercase;">DESGLOSE RECUBRIMIENTOS</th></tr></thead>
                            <tbody>\`;
                            if(Object.keys(wallSummary.recubrimientos).length === 0) html += \`<tr style="height: 18pt;"><td colspan="2" style="text-align: center; color: #94a3b8; font-size: 8pt; font-family: Arial, sans-serif; border: 1px solid #e2e8f0; vertical-align: middle;">Sin datos</td></tr>\`;
                            Object.entries(wallSummary.recubrimientos).forEach(([tipo, area]) => { 
                               html += \`<tr style="text-align: center; background-color: #ffffff; height: 18pt;">
                                  <td style="font-weight: bold; color: #d200d2; text-transform: uppercase; font-size: 8pt; font-family: Arial, sans-serif; border: none; border-bottom: 1px solid #e2e8f0; border-left: 1px solid #e2e8f0; vertical-align: middle; text-align: center;">\${tipo}</td>
                                  <td style="font-weight: bold; color: #d200d2; font-size: 9pt; font-family: Arial, sans-serif; border: none; border-bottom: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0; vertical-align: middle; mso-number-format:'0\\.00'; text-align: center;">\${area.toFixed(2)}</td>
                               </tr>\`; 
                            });
                   html += \`</tbody></table></td></tr></table>\`;

                   exportFormattedExcel(html, \`Matriz_Muros_\${levelNameExport}\`);
               }} className="px-6 py-2 bg-emerald-600 text-white font-black rounded-lg hover:bg-emerald-500 text-[10px] uppercase tracking-wider flex items-center gap-2 ml-auto"><FileDown size={14}/> Exportar Excel</button>
             </div>
             <div className="bg-slate-50 p-6 border-t border-slate-200 shrink-0">
                <h3 className="text-sm font-black uppercase tracking-widest text-slate-800 mb-6 flex items-center gap-2"><LayoutDashboard size={18} className="text-indigo-600" /> Resumen para Generadores</h3>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                  <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-100">
                     <span className="text-[10px] font-black text-emerald-600 uppercase tracking-widest block mb-3 border-b border-slate-100 pb-2">Área Neta por Espesor</span>
                     {Object.keys(wallSummary.murosPorAncho).length === 0 && <span className="text-xs font-bold text-slate-300">Sin capturas</span>}
                     {Object.entries(wallSummary.murosPorAncho).map(([ancho, area]) => ( <div key={ancho} className="flex justify-between items-center border-b border-slate-50 last:border-0 py-2"><span className="truncate w-2/3 uppercase text-slate-700 text-2xl font-black">Muro {ancho}m</span><div className="flex items-center gap-2"><span className="text-emerald-700 font-black text-xl">{area.toFixed(2)} m2</span><button onClick={() => handleCopyWalls('TOTAL_NETO', ancho)} className="p-1.5 bg-slate-100 text-slate-500 rounded-lg hover:bg-emerald-100 hover:text-emerald-700 transition-colors shadow-sm ml-1" title={\`Copiar muros de \${ancho}m\`}><Clipboard size={14}/></button></div></div> ))}
                  </div>
                  <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-100">
                     <span className="text-[10px] font-black text-purple-600 uppercase tracking-widest block mb-3 border-b border-slate-100 pb-2">Desglose Aplanados</span>
                     {Object.keys(wallSummary.aplanados).length === 0 && <span className="text-xs font-bold text-slate-300">Sin capturas</span>}
                     {Object.entries(wallSummary.aplanados).map(([tipo, area]) => ( <div key={tipo} className="flex justify-between items-center border-b border-slate-50 last:border-0 py-2"><span className="truncate w-2/3 uppercase text-slate-700 text-sm font-black">{tipo}</span><div className="flex items-center gap-2"><span className="text-purple-700 font-black text-lg">{area.toFixed(2)} m2</span><button onClick={() => handleCopyWalls('APLANADO', tipo)} className="p-1.5 bg-slate-100 text-slate-500 rounded-lg hover:bg-purple-100 hover:text-purple-700 transition-colors shadow-sm ml-1" title={\`Copiar detalles de muros con \${tipo}\`}><Clipboard size={14}/></button></div></div> ))}
                  </div>
                  <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-100">
                     <span className="text-[10px] font-black text-fuchsia-600 uppercase tracking-widest block mb-3 border-b border-slate-100 pb-2">Desglose Recubrimientos</span>
                     {Object.keys(wallSummary.recubrimientos).length === 0 && <span className="text-xs font-bold text-slate-300">Sin capturas</span>}
                     {Object.entries(wallSummary.recubrimientos).map(([tipo, area]) => ( <div key={tipo} className="flex justify-between items-center border-b border-slate-50 last:border-0 py-2"><span className="truncate w-2/3 uppercase text-slate-700 text-sm font-black">{tipo}</span><div className="flex items-center gap-2"><span className="text-fuchsia-700 font-black text-lg">{area.toFixed(2)} m2</span><button onClick={() => handleCopyWalls('RECUBRIMIENTO', tipo)} className="p-1.5 bg-slate-100 text-slate-500 rounded-lg hover:bg-fuchsia-100 hover:text-fuchsia-700 transition-colors shadow-sm ml-1" title={\`Copiar detalles de muros con \${tipo}\`}><Clipboard size={14}/></button></div></div> ))}
                  </div>
                </div>
             </div>
           </div>
        ) : showSettings ? (`;

const replaceRegex = /\) \: showWallGenerator \? \((.*?)\) \: showSettings \? \(/s;

if (replaceRegex.test(code)) {
    code = code.replace(replaceRegex, newWallBlock);
    fs.writeFileSync('src/App.tsx', code);
    console.log('Successfully replaced WallBlock');
} else {
    console.log('Failed to match WallBlock regex');
}
