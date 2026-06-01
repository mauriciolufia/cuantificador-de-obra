import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

export const exportRealExcelElegante = async (activePartida, conceptos, niveles, generadores) => {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Xdifica';
  
  const safeSheetName = (name) => {
    return (name || "Hoja").substring(0, 31).replace(/[\\/?*[\]]/g, '');
  };

  const summarySheet = workbook.addWorksheet(safeSheetName('Catálogo General'));

  const cellMap = {};

  const borderAll = {
    top: { style: 'thin', color: { argb: 'FFB4C6E7' } },
    left: { style: 'thin', color: { argb: 'FFB4C6E7' } },
    bottom: { style: 'thin', color: { argb: 'FFB4C6E7' } },
    right: { style: 'thin', color: { argb: 'FFB4C6E7' } }
  };

  const borderDashed = {
    top: { style: 'dashed', color: { argb: 'FF808080' } },
    left: { style: 'dashed', color: { argb: 'FF808080' } },
    bottom: { style: 'dashed', color: { argb: 'FF808080' } },
    right: { style: 'dashed', color: { argb: 'FF808080' } }
  };

  for (const n of niveles) {
    let nivelHasData = false;
    for (const c of conceptos) {
      const rows = generadores[c.id]?.[n.id]?.rows || [];
      if (rows.length > 0) nivelHasData = true;
    }

    if (!nivelHasData) continue;

    const sheetName = safeSheetName(`Nivel ${n.nombre}`);
    const sheet = workbook.addWorksheet(sheetName);

    sheet.columns = [
      { key: 'pad', width: 3 },
      { key: 'loc', width: 25 },
      { key: 'largo', width: 12 },
      { key: 'ancho', width: 12 },
      { key: 'kg_ml', width: 12 },
      { key: 'alto', width: 12 },
      { key: 'vol_pza', width: 15 },
      { key: 'piezas', width: 10 },
      { key: 'volTotal', width: 15 },
    ];
    
    // CÉDULA DE CUANTIFICACIÓN (GENERADOR) - NIVEL: [NOMBRE]
    sheet.getRow(1).height = 15;
    sheet.mergeCells('B2:I2');
    const titleCell = sheet.getCell('B2');
    titleCell.value = `CÉDULA DE CUANTIFICACIÓN (GENERADOR) - NIVEL: ${n.nombre}`;
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B1A30' } };
    titleCell.font = { color: { argb: 'FFFFFFFF' }, bold: true, size: 12 };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    sheet.getRow(2).height = 25;
    sheet.getRow(3).height = 15;

    let currentRow = 4;

    for (const c of conceptos) {
      const rows = generadores[c.id]?.[n.id]?.rows || [];
      if (rows.length === 0) continue;

      // Concept Title Row
      sheet.mergeCells(`C${currentRow}:H${currentRow}`);
      
      const charCount = (c.descripcion || "").length;
      const estimatedLines = Math.ceil(charCount / 70) || 1;
      sheet.getRow(currentRow).height = Math.max(25, estimatedLines * 15 + 10);

      const codCell = sheet.getCell(`B${currentRow}`);
      codCell.value = `CÓD: ${c.id}`;
      codCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } };
      codCell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
      codCell.alignment = { horizontal: 'center', vertical: 'middle' };
      codCell.border = borderAll;

      const descCell = sheet.getCell(`C${currentRow}`);
      descCell.value = c.descripcion;
      descCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } };
      descCell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
      descCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      descCell.border = borderAll;

      // Unid cell
      const unidCell = sheet.getCell(`I${currentRow}`);
      unidCell.value = (c.unidad || "").toUpperCase();
      unidCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } };
      unidCell.font = { color: { argb: 'FF00B0F0' }, bold: true };
      unidCell.alignment = { horizontal: 'center', vertical: 'middle' };
      unidCell.border = borderAll;

      currentRow++;

      // Data headers
      const headers = ['LOCALIZACIÓN', 'LARGO', 'ANCHO', 'KG/ML', 'ALTO', 'VOL/PZA', 'PZAS', 'VOLUMEN TOT'];
      headers.forEach((h, i) => {
        const cell = sheet.getCell(currentRow, i + 2);
        cell.value = h;
        cell.font = { color: { argb: 'FF002060' }, bold: true }; // Dark blue header text
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = borderAll;
      });

      currentRow++;
      let rowStart = currentRow;

      rows.forEach((r) => {
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

        const isActuallyZero = L === 0 && A === 0 && H === 0 && K === 0;

        let overrideFormula = null;
        
        if (r.isPasos) {
            overrideFormula = `ROUNDDOWN(C${currentRow}/4, 0)`; // C is Largo
        }

        const locParts = [r.eje, r.claveLoc].filter(Boolean).join(" - ");

        const cellLoc = sheet.getCell(`B${currentRow}`);
        cellLoc.value = locParts;

        const cellL = sheet.getCell(`C${currentRow}`);
        cellL.value = L > 0 ? parseFloat(r.largo) : "-";
        
        const cellA = sheet.getCell(`D${currentRow}`);
        cellA.value = A > 0 ? parseFloat(r.ancho) : "-";
        
        const cellK = sheet.getCell(`E${currentRow}`);
        cellK.value = K > 0 ? parseFloat(r.kg_ml) : "-";
        
        const cellH = sheet.getCell(`F${currentRow}`);
        cellH.value = H > 0 ? parseFloat(r.alto) : "-";

        const cellVolPza = sheet.getCell(`G${currentRow}`);
        if (overrideFormula) {
            cellVolPza.value = { formula: overrideFormula };
        } else if (isActuallyZero) {
            cellVolPza.value = "-";
        } else if (typeof r.overrideVolPza === "number") {
            cellVolPza.value = r.overrideVolPza;
        } else {
            let factors = [];
            if (L > 0) factors.push(`C${currentRow}`);
            if (A > 0) factors.push(`D${currentRow}`);
            if (H > 0) factors.push(`F${currentRow}`);
            if (K > 0) factors.push(`E${currentRow}`);
            if (factors.length > 0) {
               cellVolPza.value = { formula: factors.join('*') };
            } else {
               cellVolPza.value = "-";
            }
        }
        
        const cellPzas = sheet.getCell(`H${currentRow}`);
        cellPzas.value = pzas;

        const cellTotal = sheet.getCell(`I${currentRow}`);
        if (typeof r.overrideTotal === "number") {
            cellTotal.value = r.overrideTotal;
        } else {
            cellTotal.value = { formula: `G${currentRow}*H${currentRow}` };
        }

        // Output formatting & alignment
        ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach((col) => {
            const cell = sheet.getCell(`${col}${currentRow}`);
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = borderAll;
            if (col !== 'B' && cell.value !== "-") {
                cell.numFmt = '0.00';
            }
        });

        currentRow++;
      });

      // Subtotal row
      sheet.mergeCells(`B${currentRow}:H${currentRow}`);
      const lblCell = sheet.getCell(`B${currentRow}`);
      lblCell.value = 'TOTAL:';
      lblCell.alignment = { horizontal: 'right', vertical: 'middle' };
      lblCell.font = { color: { argb: 'FF808080' } };
      lblCell.border = borderAll;

      const valCell = sheet.getCell(`I${currentRow}`);
      valCell.value = { formula: `SUM(I${rowStart}:I${currentRow-1})` };
      valCell.numFmt = '0.00';
      valCell.font = { color: { argb: 'FF0070C0' }, bold: true };
      valCell.alignment = { horizontal: 'center', vertical: 'middle' };
      valCell.border = borderAll;

      // Save reference mapping
      if (!cellMap[c.id]) cellMap[c.id] = {};
      cellMap[c.id][n.id] = `'${sheetName}'!I${currentRow}`;

      currentRow += 2;

      // Croquis area
      const cLevel = generadores[c.id]?.[n.id] || {};
      const refImages = cLevel.images || (cLevel.image ? [cLevel.image] : []);
      
      if (refImages.length === 0) {
          const rowsNeeded = 10;
          sheet.mergeCells(`B${currentRow}:I${currentRow + rowsNeeded}`);
          const croquisCell = sheet.getCell(`B${currentRow}`);
          croquisCell.value = "[ ESPACIO PARA INSERTAR CROQUIS / IMAGEN DE REFERENCIA ]";
          croquisCell.font = { color: { argb: 'FFA6A6A6' } };
          croquisCell.border = borderDashed;
          croquisCell.alignment = { horizontal: 'center', vertical: 'middle' };
          currentRow += rowsNeeded + 2;
      } else {
          for (let i = 0; i < refImages.length; i++) {
              const refImage = refImages[i];
              let rowsNeeded = 10;
              let mergedCroquis = false;

              if (refImage && refImage.startsWith('data:image')) {
                  try {
                      const imgDim = await new Promise((resolve) => {
                         const img = new Image();
                         img.onload = () => resolve({ w: img.width, h: img.height });
                         img.onerror = () => resolve({ w: 500, h: 300 });
                         img.src = refImage;
                      });

                      let imgW = imgDim.w;
                      let imgH = imgDim.h;
                      const MAX_W = 750; // aprox width for columns B to I
                      if (imgW > MAX_W) {
                          imgH = Math.round((imgH * MAX_W) / imgW);
                          imgW = MAX_W;
                      }
                      // Calculate rows needed (approx 20px per row)
                      rowsNeeded = Math.max(10, Math.ceil(imgH / 20));

                      sheet.mergeCells(`B${currentRow}:I${currentRow + rowsNeeded}`);
                      mergedCroquis = true;
                      
                      const croquisCell = sheet.getCell(`B${currentRow}`);
                      croquisCell.border = borderAll;

                      const base64str = refImage.split(';base64,').pop();
                      const extFormat = refImage.includes('png') ? 'png' : 'jpeg';
                      const imageId = workbook.addImage({
                        base64: base64str,
                        extension: extFormat,
                      });
                      
                      // We'll use ext instead of br so aspect ratio is solid
                      sheet.addImage(imageId, {
                        tl: { col: 1.05, row: currentRow - 0.9 }, // inset slightly
                        ext: { width: imgW, height: imgH }
                      });
                  } catch (e) {
                      console.error("Error adding image to Excel:", e);
                  }
              } 
              
              if (!mergedCroquis) {
                  sheet.mergeCells(`B${currentRow}:I${currentRow + rowsNeeded}`);
                  const croquisCell = sheet.getCell(`B${currentRow}`);
                  if (refImage && refImage.startsWith('data:image')) {
                      croquisCell.value = "[ ERROR AL CARGAR IMAGEN ]";
                      croquisCell.font = { color: { argb: 'FFFF0000' } };
                  } else {
                      croquisCell.value = "[ ESPACIO PARA INSERTAR CROQUIS / IMAGEN DE REFERENCIA ]";
                      croquisCell.font = { color: { argb: 'FFA6A6A6' } };
                      croquisCell.border = borderDashed;
                  }
                  croquisCell.alignment = { horizontal: 'center', vertical: 'middle' };
              }
              
              currentRow += rowsNeeded + 2;
          }
      }
    }
  }

  // ========== CATÁLOGO GENERAL ==========
  const catCols = [
    { key: 'pad', width: 3 },
    { key: 'codigo', width: 15 },
    { key: 'clave', width: 12 },
    { key: 'cc', width: 12 },
    { key: 'justificacion', width: 25 },
    { key: 'descripcion', width: 45 },
    { key: 'unidad', width: 10 },
  ];
  
  niveles.forEach((n) => {
    catCols.push({ key: `niv_${n.id}`, width: 12 });
  });
  
  catCols.push({ key: 'total', width: 15 });

  summarySheet.columns = catCols;

  // Header Title
  summarySheet.getRow(1).height = 15;
  summarySheet.mergeCells(2, 2, 2, catCols.length);
  const head1 = summarySheet.getCell('B2');
  head1.value = `CATÁLOGO GENERAL - PARTIDA: ${activePartida.nombre}`;
  head1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B1A30' } };
  head1.font = { color: { argb: 'FFFFFFFF' }, bold: true, size: 11 };
  head1.alignment = { horizontal: 'center', vertical: 'middle' };
  
  summarySheet.getRow(2).height = 25;
  summarySheet.getRow(3).height = 15;

  // Header Row 4
  const mainHeaders = ['CÓDIGO', 'CLAVE', 'CC', 'JUSTIFICACIÓN', 'DESCRIPCIÓN', 'UNID.'];
  
  niveles.forEach(n => mainHeaders.push(n.nombre));
  mainHeaders.push('TOTAL');

  mainHeaders.forEach((h, i) => {
    const cell = summarySheet.getCell(4, i + 2);
    cell.value = h;
    // El "Total" tiene fondo negro
    if (i === mainHeaders.length - 1) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } };
    } else {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B1A30' } };
    }
    cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = {
        left: { style: 'thin', color: { argb: 'FFFFFFFF' } },
        right: { style: 'thin', color: { argb: 'FFFFFFFF' } }
    };
  });
  summarySheet.getRow(4).height = 20;

  let rIdx = 5;
  conceptos.forEach((c) => {
      const rowVals = [
          "", // Pad
          c.id,
          c.clave || "",
          c.cc || "",
          c.justificacion || "",
          c.descripcion || "",
          c.unidad || ""
      ];

      const added = summarySheet.addRow(rowVals);
      
      // Formatting first 6 cols
      const cCell = summarySheet.getCell(`B${rIdx}`);
      cCell.font = { color: { argb: 'FF002060' }, bold: true };
      cCell.alignment = { horizontal: 'center', vertical: 'middle' };

      const descCell = summarySheet.getCell(`F${rIdx}`);
      descCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

      const uCell = summarySheet.getCell(`G${rIdx}`);
      uCell.font = { color: { argb: 'FF0070C0' }, bold: true };
      uCell.alignment = { horizontal: 'center', vertical: 'middle' };

      // Niveles
      const baseColStart = 8; 
      let totalFormulaParts = [];

      niveles.forEach((n, i) => {
          const colNum = baseColStart + i; 
          const cell = summarySheet.getRow(rIdx).getCell(colNum);
          
          if (cellMap[c.id] && cellMap[c.id][n.id]) {
              cell.value = { formula: cellMap[c.id][n.id] };
              cell.numFmt = '0.00';
              const colLetter = cell._column.letter;
              totalFormulaParts.push(`${colLetter}${rIdx}`);
          } else {
              cell.value = "-";
          }
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.font = { color: { argb: 'FF006368' }, bold: true };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE3FDE4' } };
      });

      // Total
      const totalCell = summarySheet.getRow(rIdx).getCell(baseColStart + niveles.length);
      if (totalFormulaParts.length > 0) {
          totalCell.value = { formula: `SUM(${totalFormulaParts.join(',')})` };
          totalCell.numFmt = '0.00';
      } else {
          totalCell.value = "-";
      }
      totalCell.font = { bold: true };
      totalCell.alignment = { horizontal: 'center', vertical: 'middle' };

      // Apply borders to all row cells
      for(let cc = 2; cc <= catCols.length; cc++) {
          const cell = summarySheet.getCell(rIdx, cc);
          if (cc === 6) { // Descripcion
              cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          } else {
              cell.alignment = { horizontal: 'center', vertical: 'middle' };
          }
          cell.border = borderAll;
      }

      rIdx++;
  });

  const buffer = await workbook.xlsx.writeBuffer();
  saveAs(new Blob([buffer]), `Proyecto_${activePartida.nombre}.xlsx`);
};

