let templateWorkbook = null;
let salesData = null;
let groupedData = {};

// Şablonu yükleme
document.getElementById('uploadTemplate').addEventListener('change', async (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = async (e) => {
        const arrayBuffer = e.target.result;
        templateWorkbook = arrayBuffer;
        console.log("Şablon Yüklendi");
    };

    reader.readAsArrayBuffer(file);
});

// 1) SATIŞ DOSYASINI OKURKEN: raw:false ile formattı string al
document.getElementById('uploadSalesData').addEventListener('change', (event) => {
  const file = event.target.files[0];
  const reader = new FileReader();

  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // ÖNEMLİ: raw:false → hücrelerin gösterilen (string) değerini alır
    // böylece uzun sayılar bilimsel gösterime düşmeden string olarak gelir.
    let json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

    groupedData = groupBySender(json);
    populateDropdown(Object.keys(groupedData));
    console.log("Müşteri Satış Verileri Yüklendi");
  };

  reader.readAsArrayBuffer(file);
});


function excelDateToJSDate(serial) {
    const excelEpoch = new Date(1899, 11, 30); // Excel tarihlerinin başlangıç noktası
    return new Date(excelEpoch.getTime() + serial * 86400000); // Günleri milisaniyeye çevir
}

function groupBySender(data) {
    const headers = data[0];
    const senderIndex = headers.indexOf('GÖNDEREN');
    if (senderIndex === -1) {
        console.warn("GÖNDEREN sütunu bulunamadı.");
        return {};
    }

    const grouped = {};
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const sender = row[senderIndex];
        if (!grouped[sender]) {
            grouped[sender] = [headers];
        }
        grouped[sender].push(row);
    }

    return grouped;
}

function populateDropdown(customers) {
    const dropdown = document.getElementById('customerSelect');
    dropdown.innerHTML = '<option value="">Müşteri seçin</option>';

    customers.forEach((customer) => {
        const option = document.createElement('option');
        option.value = customer;
        option.textContent = customer;
        dropdown.appendChild(option);
    });

    document.getElementById('generateExcel').disabled = false;
    document.getElementById('generateAllExcel').disabled = false;
}

document.getElementById('generateExcel').addEventListener('click', async () => {
    const selectedCustomer = document.getElementById('customerSelect').value;
    if (!selectedCustomer || !templateWorkbook || !groupedData[selectedCustomer]) {
        alert("Lütfen geçerli bir müşteri ve şablon yükleyin.");
        return;
    }

    const customerData = groupedData[selectedCustomer];
    const totalDebt = calculateTotalNavlun(customerData);

    const filledTemplate = await fillTemplate(templateWorkbook, selectedCustomer, customerData, totalDebt);

    const blob = new Blob([filledTemplate], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${selectedCustomer}_Haftalık_Gönderi_Raporu.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
});

function calculateTotalNavlun(data) {
    console.log(data)
    const headers = data[0];
    const navlunIndex = headers.indexOf('NAVLUN TOTAL');

    if (navlunIndex === -1) {
        console.warn("NAVLUN sütunu bulunamadı.");
        return 0;
    }

    let total = 0;
    for (let i = 1; i < data.length; i++) {
        const navlun = parseFloat(data[i][navlunIndex]) || 0;
        total += navlun;
    }
    console.log(total)
    return total.toLocaleString('tr-TR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// 2) EXCELJS'E YAZARKEN: ID ve TAKİP sütunlarını Text (@) yap ve stringe çevir
async function fillTemplate(templateFile, sender, data, totalDebt) {
  const workbook = new ExcelJS.Workbook();
  let indexFormula = 5;
  await workbook.xlsx.load(templateFile);

  const sheet = workbook.getWorksheet(1);
  const startRow = 2;

  // Toplamı sayı olarak yazmak istiyorsan numFmt kullan (yoksa string kalsın)
  // sheet.getCell('L3').value = Number(
  //   (totalDebt || '0').toString().replace(/\./g,'').replace(',', '.')
  // );
  // sheet.getCell('L3').numFmt = '#,##0.00';

  // Şu anki haliyle string ise dokunma:
  sheet.getCell('L3').value = totalDebt;

  // Burada yazım hatası var: 0,0 → 0.0 olmalı
  sheet.getCell('L4').value = 0.0;

  // Başlıklar
  const headers = data[0] || [];
  const dateColIdx   = headers.indexOf('TARİH');
  const idColIdx     = headers.indexOf('SHIPREADY ID');
  const takipColIdx  = headers.indexOf('TAKİP');

  // Bu sütunların tamamını Text formatına al
  if (idColIdx !== -1)    sheet.getColumn(idColIdx + 1).numFmt = '@';
  if (takipColIdx !== -1) sheet.getColumn(takipColIdx + 1).numFmt = '@';

  // Satır ekleme
  data.slice(1).forEach((row, index) => {
    indexFormula++;

    // Çalışacağımız kopya (orijinali bozma)
    const r = Array.isArray(row) ? [...row] : row;

    // TARİH sayısal geldiyse Date stringe çevir
    if (dateColIdx !== -1 && typeof r[dateColIdx] === 'number') {
      r[dateColIdx] = excelDateToJSDate(r[dateColIdx]).toLocaleDateString('tr-TR');
    }

    // UZUN NUMARALARI STRINGE ZORLA (bilimsel gösterimi önler, precision’ı korur)
    if (idColIdx !== -1 && r[idColIdx] != null) {
      r[idColIdx] = String(r[idColIdx]); // başına ' eklemene gerek yok
    }
    if (takipColIdx !== -1 && r[takipColIdx] != null) {
      r[takipColIdx] = String(r[takipColIdx]);
    }

    const newRow = sheet.insertRow(startRow + index, r);

    // Stil kopyalama (varsa şablondaki ilk veri satırından)
    r.forEach((_, colIndex) => {
      const sourceCell = sheet.getRow(startRow).getCell(colIndex + 1);
      const targetCell = newRow.getCell(colIndex + 1);
      targetCell.style = { ...sourceCell.style };
    });
  });

  // Formül
  const indexFormulaText         = "L" + indexFormula;
  const indexFormulaTextOneUpper = "L" + (indexFormula - 1);
  const indexFormulaTextTwoUpper = "L" + (indexFormula - 2);
  sheet.getCell(indexFormulaText).value = { formula: `${indexFormulaTextOneUpper} + ${indexFormulaTextTwoUpper}` };

  const buffer = await workbook.xlsx.writeBuffer();
  return buffer;
}


function downloadExcel(workbook, fileName) {
    XLSX.writeFile(workbook, fileName);
}

document.getElementById('generateAllExcel').addEventListener('click', async () => {
    if (!templateWorkbook || !Object.keys(groupedData).length) {
        alert("Lütfen bir şablon ve müşteri verisi yükleyin.");
        return;
    }

    const zip = new JSZip();

    for (const customer of Object.keys(groupedData)) {
        const customerData = groupedData[customer];
        const totalDebt = calculateTotalNavlun(customerData);
        const filledTemplate = await fillTemplate(templateWorkbook, customer, customerData, totalDebt);
        zip.file(`${customer}_Haftalık_Gönderi_Raporu.xlsx`, filledTemplate, { binary: true });
    }

    zip.generateAsync({ type: 'blob' }).then((content) => {
        const url = URL.createObjectURL(content);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Toplu_Haftalık_Rapor.zip`;
        a.click();
        URL.revokeObjectURL(url);
    });
});
