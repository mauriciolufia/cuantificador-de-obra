const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf8');

const tableCode = `        ) : activePartidaId ? (
          <div className="overflow-x-auto bg-white rounded-lg shadow border border-gray-200 flex-1">
            <table className="text-sm text-left border-collapse table-fixed w-full">
              <thead className="bg-slate-800 text-white uppercase tracking-wider select-none sticky top-0 z-30 font-black">
                <tr>
                  <th className="px-1 py-2 border border-slate-700 text-center sticky left-0 bg-slate-800 z-30" style={{ width: colWidths.select }}><button onClick={() => setSelectedIds(selectedIds.length === conceptos.length && conceptos.length > 0 ? [] : conceptos.map(c => c.id))}>{conceptos.length > 0 && selectedIds.length === conceptos.length ? <CheckSquare size={14} /> : <Square size={14} />}</button></th>
                  {[ { id: 'id', label: 'Código' }, { id: 'clave', label: 'Clave' }, { id: 'cc', label: 'CC' }, { id: 'justificacion', label: 'Justif.' }, { id: 'descripcion', label: 'Descripción' }, { id: 'unidad', label: 'Unid.' } ].map((col) => ( <th key={col.id} className="px-1.5 py-2 border border-slate-700 relative group text-[9px] font-black leading-tight" style={{ width: colWidths[col.id] }}><div className="truncate w-full">{col.label}</div><div onMouseDown={(e) => startResizing(col.id, e, 'summary')} className="absolute right-0 top-0 bottom-0 w-2 cursor-col-resize hover:bg-blue-400 z-40 transition-colors" /></th> ))}
                  {niveles.map(n => ( <th key={n.id} className="px-1 py-1.5 border border-slate-700 text-center bg-blue-900 relative group" style={{ width: colWidths.niveles }}><div className="flex flex-col items-center gap-0.5 w-full overflow-hidden"><span className="text-[8px] md:text-[9px] font-black truncate max-w-full px-1">{n.nombre}</span><button onClick={() => setActiveLevelReport(n)} title="Ver Cédula de Cuantificación" className="bg-blue-700 hover:bg-blue-600 p-0.5 rounded transition-colors"><ListPlus size={12} /></button></div></th> ))}
                  <th className="px-1 py-2 border border-slate-700 text-center bg-black relative" style={{ width: colWidths.totales }}><span className="text-[9px] font-black">TOTAL</span></th><th className="px-1 py-2 border border-slate-700" style={{ width: 30 }}></th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {conceptos.map((c, idx) => (
                  <tr key={c.id} className={idx % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                    <td className="p-0 border text-center sticky left-0 bg-inherit z-10"><button onClick={() => toggleSelectConcept(c.id)} className={\`\${selectedIds.includes(c.id) ? 'text-blue-600' : 'text-slate-300'}\`}>{selectedIds.includes(c.id) ? <CheckSquare size={14} /> : <Square size={14} />}</button></td>
                    <td className="p-0 border font-bold text-blue-700 leading-tight"><DebouncedCell value={c.id} onChange={(v) => updateConceptoId(c.id, v)} onDoubleClick={() => { setConceptSelectorFor(c.id); setConceptSearchTerm(''); }} title="Doble clic para buscador" className="w-full h-full p-1.5 bg-transparent text-left outline-none text-blue-700 font-black min-w-0 cursor-pointer hover:bg-blue-50/50 text-[16px]" /></td>
                    <td className="p-0 border text-center text-[10px]"><DebouncedCell value={c.clave} onChange={(v) => updateConceptoField(c.id, 'clave', v)} className="w-full h-full p-1.5 bg-transparent text-center outline-none min-w-0" /></td>
                    <td className="p-0 border text-center text-[10px]"><DebouncedCell value={c.cc} onChange={(v) => updateConceptoField(c.id, 'cc', v)} className="w-full h-full p-1.5 bg-transparent text-center outline-none min-w-0" /></td>
                    <td className="p-0 border italic text-gray-400 text-[9px]"><DebouncedCell isTextArea rows={2} value={c.justificacion} onChange={(v) => updateConceptoField(c.id, 'justificacion', v)} className="w-full h-full p-1 bg-transparent outline-none resize-none min-w-0 leading-tight" /></td>
                    <td className="p-0 border text-xs md:text-sm font-bold"><DebouncedCell isTextArea rows={2} value={c.descripcion} onChange={(v) => updateConceptoField(c.id, 'descripcion', v)} className="w-full h-full p-1 bg-transparent outline-none resize-none min-w-0 leading-tight uppercase" /></td>
                    <td className="p-0 border text-center font-black"><DebouncedCell value={c.unidad} onChange={(v) => updateConceptoField(c.id, 'unidad', v)} className="w-full h-full p-1.5 bg-transparent text-center outline-none text-blue-700 uppercase min-w-0 text-[15px]" /></td>
                    {niveles.map(n => { const vol = getVolumenNivel(c.id, n.id); return <td key={n.id} onClick={() => { setEditingModal({ concepto: c, nivel: n }); setSelectedGeneratorRows([]); }} className={\`p-1 md:p-1.5 border text-center cursor-pointer font-black text-[11px] md:text-xs leading-tight \${vol > 0 ? 'bg-emerald-50 text-emerald-800' : 'text-slate-200 hover:bg-blue-50'}\`}>{vol.toFixed(2)}</td>; })}
                    <td className="p-1 md:p-1.5 border text-center font-black bg-gray-100 text-[11px] md:text-xs leading-tight text-slate-800">{getTotalConcepto(c.id).toFixed(2)}</td>
                    <td className="p-0 border text-center"><button onClick={() => updateActiveConceptos(prev => prev.filter(x => x.id !== c.id))} className="text-red-300 hover:text-red-500 transition-colors p-1"><Trash2 size={14} /></button></td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div className="p-3 md:p-4 flex gap-4 sticky bottom-0 bg-white border-t">
              <button onClick={() => updateActiveConceptos(prev => [...prev, { id: \`C-\${Date.now().toString(36).toUpperCase()}\`, clave:'', cc:'', justificacion:'', descripcion:'', unidad:'m2' }])} className="px-5 py-2 bg-slate-700 text-white rounded-xl text-[10px] md:text-xs font-black shadow-md transition-transform active:scale-95 uppercase tracking-wider">+ Agregar Concepto</button>
              <button onClick={() => {
                  let html = \`<table style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">
                      <colgroup>
                         <col width="120" />
                         <col width="100" />
                         <col width="100" />
                         <col width="250" />
                         <col width="450" />
                         <col width="80" />
                         \${niveles.map(() => '<col width="100" />').join('')}
                         <col width="100" />
                      </colgroup>
                      <thead>
                      <tr style="height: 25pt;">
                         <td colspan="\${6 + niveles.length + 1}" style="background-color: #0b1a30; color: #ffffff; font-size: 10pt; font-weight: bold; text-align: center; vertical-align: middle; padding: 8px; border: none; text-transform: uppercase;">
                            CATÁLOGO GENERAL - PARTIDA: \${activePartida.nombre}
                         </td>
                      </tr>
                      <tr><td colspan="\${6 + niveles.length + 1}" style="height: 10px; border: none; background-color: #ffffff;"></td></tr>
                      <tr style="height: 20pt;">
                         <td style="background-color: #1f3864; color: #ffffff; text-align: center; vertical-align: middle; font-weight: bold; font-size: 10pt; border: 1px solid #ffffff;">CÓDIGO</td>
                         <td style="background-color: #1f3864; color: #ffffff; text-align: center; vertical-align: middle; font-weight: bold; font-size: 10pt; border: 1px solid #ffffff;">CLAVE</td>
                         <td style="background-color: #1f3864; color: #ffffff; text-align: center; vertical-align: middle; font-weight: bold; font-size: 10pt; border: 1px solid #ffffff;">CC</td>
                         <td style="background-color: #1f3864; color: #ffffff; text-align: center; vertical-align: middle; font-weight: bold; font-size: 10pt; border: 1px solid #ffffff;">JUSTIFICACIÓN</td>
                         <td style="background-color: #1f3864; color: #ffffff; text-align: center; vertical-align: middle; font-weight: bold; font-size: 10pt; border: 1px solid #ffffff;">DESCRIPCIÓN</td>
                         <td style="background-color: #1f3864; color: #ffffff; text-align: center; vertical-align: middle; font-weight: bold; font-size: 10pt; border: 1px solid #ffffff;">UNID.</td>
                         \${niveles.map(n => \`<td style="background-color: #1f3864; color: #ffffff; text-align: center; vertical-align: middle; font-weight: bold; font-size: 10pt; border: 1px solid #ffffff;">\${n.nombre.toUpperCase()}</td>\`).join('')}
                         <td style="background-color: #000000; color: #ffffff; text-align: center; vertical-align: middle; font-weight: bold; font-size: 10pt; border: 1px solid #ffffff;">TOTAL</td>
                      </tr>
                  </thead><tbody>\`;

                  conceptos.forEach((c) => {
                      const just = (c.justificacion||'').replace(/\\n/g, ' ');
                      const desc = (c.descripcion||'').replace(/\\n/g, ' ');
                      const total = getTotalConcepto(c.id);
                      
                      html += \`<tr style="height: 25pt; background-color: #f4f6f9;">
                         <td style="font-weight: bold; color: #002060; text-align: center; vertical-align: middle; padding: 6px; border: 1px solid #b4c6e7; font-size: 12pt;">\${c.id}</td>
                         <td style="font-weight: bold; color: #404040; text-align: center; vertical-align: middle; padding: 6px; border: 1px solid #b4c6e7; font-size: 10pt;">\${c.clave || ''}</td>
                         <td style="font-weight: bold; color: #404040; text-align: center; vertical-align: middle; padding: 6px; border: 1px solid #b4c6e7; font-size: 10pt;">\${c.cc || ''}</td>
                         <td style="color: #404040; text-align: center; vertical-align: middle; padding: 6px; border: 1px solid #b4c6e7; font-size: 10pt;">\${just}</td>
                         <td style="font-weight: bold; color: #404040; text-align: center; vertical-align: middle; padding: 6px; border: 1px solid #b4c6e7; font-size: 10pt;">\${desc}</td>
                         <td style="color: #0070c0; font-weight: bold; text-align: center; vertical-align: middle; padding: 6px; border: 1px solid #b4c6e7; font-size: 12pt;">\${c.unidad || ''}</td>
                         \${niveles.map(n => {
                             const vol = getVolumenNivel(c.id, n.id);
                             if (vol > 0) {
                                 return \`<td style="font-weight: bold; background-color: #e2efda; color: #00b050; text-align: right; vertical-align: middle; padding: 6px; border: 1px solid #b4c6e7; font-size: 10pt; mso-number-format:'0\\\\.00';">\${vol.toFixed(2)}</td>\`;
                             } else {
                                 return \`<td style="color: #d9d9d9; text-align: right; vertical-align: middle; padding: 6px; border: 1px solid #b4c6e7; font-size: 10pt;">-</td>\`;
                             }
                         }).join('')}
                         <td style="font-weight: bold; color: #000000; text-align: right; vertical-align: middle; padding: 6px; border: 1px solid #b4c6e7; font-size: 10pt; mso-number-format:'0\\\\.00';">\${total > 0 ? total.toFixed(2) : '-'}</td>
                      </tr>\`;
                  });
                  html += '</tbody></table>';
                  exportFormattedExcel(html, \`Matriz_Conceptos_\${activePartida.nombre}\`);
              }} className="px-5 py-2 bg-emerald-600 text-white rounded-xl text-[10px] md:text-xs font-black shadow-md transition-transform active:scale-95 uppercase tracking-wider hover:bg-emerald-500 flex items-center gap-2 ml-auto"><FileDown size={16}/> Exportar Excel Elegante</button>
            </div>
          </div>
        ) : (`;

const replaceTarget = `        ) : (
             <div className="bg-white rounded-3xl min-h-[400px] border border-slate-100 shadow-sm flex items-center justify-center flex-col animate-in fade-in duration-500">`;

if (code.includes(replaceTarget)) {
    code = code.replace(replaceTarget, tableCode + '\n             <div className="bg-white rounded-3xl min-h-[400px] border border-slate-100 shadow-sm flex items-center justify-center flex-col animate-in fade-in duration-500">');
    fs.writeFileSync('src/App.tsx', code);
    console.log('Successfully injected activePartidaId table');
} else {
    console.log('Target not found!!');
}
