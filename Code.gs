/** * WEB APP ENTRY POINT */

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

// ============================================
// SINGLE DOGET - REMOVED DUPLICATE
// ============================================
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('RECON ERP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getModuleHtml(moduleName) {
  try {
    var fileName = moduleName.indexOf('Val') === 0 ? 'AdminVal' : moduleName;
    
    // Use a template so that <?!= include(...) ?> works
    var template = HtmlService.createTemplateFromFile(fileName);
    
    return template.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).getContent();
    
  } catch (e) {
    console.error("Template Error: " + e.message);
    return "<div style='color:red; padding:20px; background:#fff1f1; border:1px solid red;'>" +
           "<h4>⚠️ Error in " + moduleName + " or its Included Files</h4>" +
           "<p><b>Message:</b> " + e.message + "</p>" +
           "<p>Check for syntax errors (missing brackets, etc.) in your JS file.</p></div>";
  }
}

// ============================================
// FIXED: getValidationData() - NOW INCLUDES ALL DATA
// ============================================
function getValidationData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Validation");
  if (!sheet) return {};
  
  const data = sheet.getDataRange().getValues();
  const categories = [];
  const subCategories = [];
  const brands = [];
  const suppliers = [];
  const sites = [];
  const currencies = [];

  for (let i = 1; i < data.length; i++) {
    // Col A (0): Faena/Sites
    if (data[i][0]) {
      sites.push({ 
        name: data[i][0], 
        email: data[i][1] || "",
        suffix: data[i][2] || ""
      });
    }
    
    // Col F (5): Categories, Col G (6): Email
    if (data[i][5]) {
      categories.push({ 
        name: data[i][5],
        email: data[i][6] || ""
      });
    }
    
    // Col I (8): SubCategories
    if (data[i][8]) {
      subCategories.push(data[i][8]);
    }
    
    // Col K (10): Brands
    if (data[i][10]) {
      brands.push(data[i][10]);
    }
    
    // Col P (15): Suppliers, Col Q (16): Email Pro
    if (data[i][15]) {
      suppliers.push({ 
        name: data[i][15],
        email: data[i][16] || ""
      });
    }
    
    // Col S (18): Currency, Col T (19): Symbol
    if (data[i][18]) {
      currencies.push({ 
        name: data[i][18],
        symbol: data[i][19] || ""
      });
    }
  }

  return {
    categories: categories,
    subCategories: subCategories,
    brands: brands,
    suppliers: suppliers,
    sites: sites,
    currencies: currencies
  };
}

/**
 * Updates or Adds an item to the Validation sheet
 */
function updateValidationItem(type, oldName, newData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Validation");
  const data = sheet.getDataRange().getValues();

  const config = {
    'Sites': 0, 'Categories': 5, 'SubCats': 8, 'Brands': 10, 'Suppliers': 15, 'Currency': 18
  };
  const startCol = config[type];
  let rowIndex = -1;

  if (oldName) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][startCol] == oldName) { rowIndex = i + 1; break; }
    }
  }

  if (rowIndex === -1) {
    for (let i = 1; i < data.length; i++) {
      if (!data[i][startCol]) { rowIndex = i + 1; break; }
    }
    if (rowIndex === -1) rowIndex = data.length + 1;
  }

  if (type === 'Sites') {
    sheet.getRange(rowIndex, 1, 1, 3).setValues([[newData.name, newData.email, newData.suffix]]);
  } else if (type === 'Categories') {
    sheet.getRange(rowIndex, 6, 1, 2).setValues([[newData.name, newData.email]]);
  } else if (type === 'SubCats') {
    sheet.getRange(rowIndex, 9).setValue(newData.name);
  } else if (type === 'Brands') {
    sheet.getRange(rowIndex, 11).setValue(newData.name);
  } else if (type === 'Suppliers') {
    sheet.getRange(rowIndex, 16, 1, 2).setValues([[newData.name, newData.email]]);
  } else if (type === 'Currency') {
    sheet.getRange(rowIndex, 19, 1, 2).setValues([[newData.name, newData.symbol]]);
  }

  Logger.log("✅ Updated " + type + ": " + newData.name);
  return { success: true };
}

/** * PDF GENERATION & FOLDER LOGIC */
function generateProfessionalPDF(poData, targetFolder) {
  let html = HtmlService.createTemplateFromFile('PO_Template').getRawContent();
  
  const logoFileId = "1DqBJCUuFzontoeRt8Lh6yhLYifo60uRC"; 
  try {
    const blob = DriveApp.getFileById(logoFileId).getBlob();
    const base64Logo = "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes()).replace(/(\r\n|\n|\r)/gm, "");
    html = html.replace('{{LOGO_BASE64}}', base64Logo);
  } catch (e) {
    html = html.replace('{{LOGO_BASE64}}', ""); 
  }

  html = html.replace('{{PO_NO}}', poData.poId || "")
             .replace('{{DATE}}', poData.date || "")
             .replace('{{PROMISED}}', poData.date || "")
             .replace('{{SUPPLIER_NAME}}', poData.supplier || "")
             .replace('{{SUPPLIER_RUT}}', poData.supplierRut || "")
             .replace('{{SUPPLIER_ADDRESS}}', poData.supplierAddress || "")
             .replace('{{CURRENCY}}', poData.moneda || "CLP")
             .replace('{{TERMS}}', poData.condicionPago || "")
             .replace('{{NOTES}}', (poData.comments || "") + " " + (poData.numCotizacion || ""));

  let rowsHtml = "";
  poData.items.forEach((item, index) => {
    rowsHtml += `<tr>
      <td style="text-align:center;">${index + 1}</td>
      <td>${item.pn || ""}</td>
      <td>${item.desc || ""}</td>
      <td style="text-align:center;">${item.qty || 0}</td>
      <td style="text-align:right;">${Number(item.price || 0).toLocaleString('es-CL')}</td>
      <td style="text-align:right;">${(item.qty * item.price).toLocaleString('es-CL')}</td>
    </tr>`;
  });
  html = html.replace('{{TABLE_ROWS}}', rowsHtml);

  html = html.replace('{{SUBTOTAL}}', poData.subtotal.toLocaleString('es-CL'))
             .replace('{{IVA}}', poData.iva.toLocaleString('es-CL'))
             .replace('{{TOTAL}}', poData.total.toLocaleString('es-CL'));

  const blob = Utilities.newBlob(html, 'text/html', 'PO_' + poData.poId + '.html');
  const pdfFile = targetFolder.createFile(blob.getAs('application/pdf'));
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return pdfFile.getUrl();
}

function saveRequirementToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Ordenes");
  const values = sheet.getDataRange().getValues();
  
  let lastOrderId = 1000;
  if (values.length > 1) {
    lastOrderId = Math.max(...values.slice(1).map(r => Number(r[0]) || 0));
  }
  const newOrderId = lastOrderId + 1;
  const fechaStr = Utilities.formatDate(new Date(), "GMT-4", "dd/MM/yyyy HH:mm");

  const rowsToAdd = data.items.map(item => [
    newOrderId,
    fechaStr,
    item.pn,
    item.geo,
    data.faena,
    item.desc,
    item.qty,
    "Pending",
    "",
    "",
    "",
    ""
  ]);

  sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, 12).setValues(rowsToAdd);
  
  return { success: true, orderId: newOrderId };
}

function getSiteData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Validation");
    const data = sheet.getRange("A2:C" + Math.max(sheet.getLastRow(), 5)).getValues();
    return data.map(row => ({
      name: row[0] ? String(row[0]).trim() : "",
      email: row[1] ? String(row[1]).trim() : "",
      suffix: row[2] ? String(row[2]).trim() : ""
    })).filter(item => item.name !== "");
  } catch (e) { 
    Logger.log("Error in getSiteData: " + e);
    return []; 
  }
}

// ============================================
// FIXED: getInventoryData() - STANDARDIZED KEYS
// ============================================
function getInventoryData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Master"); 
  const data = sheet.getDataRange().getValues();
  
  data.shift(); 

  return data.map((row, index) => ({
    id: index,
    geo: String(row[0] || ""),           // Changed from 'g' to 'geo'
    pn: String(row[1] || ""),            // Changed from 'p' to 'pn'
    category: String(row[2] || ""),
    sub_categoria: String(row[3] || ""),
    marca: String(row[4] || ""),
    en: String(row[5] || ""),            // Keep 'en' for English
    es: String(row[6] || ""),            // Keep 'es' for Spanish
    un: String(row[7] || ""),
    min: Number(row[8]) || 0,
    max: Number(row[9]) || 0,
    stock: Number(row[10]) || 0,
    stock_fuera: Number(row[11]) || 0,
    costo_clp: Number(row[12]) || 0,
    proveedor: String(row[13] || ""),
    eta: String(row[14] || "")
  })).filter(i => i.geo || i.pn);
}

function getAutoSelectedSite() {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Validation");
    if (!sheet) return "";

    const data = sheet.getRange("A2:B" + sheet.getLastRow()).getValues();
    
    const match = data.find(row => String(row[1]).toLowerCase().trim() === userEmail);
    
    return match ? match[0] : "";
  } catch (e) {
    return "";
  }
}

function getRecipientsByCategory(category) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Validation");
  const data = sheet.getRange("F2:G" + sheet.getLastRow()).getValues();
  
  const match = data.find(row => 
    row[0].toString().toLowerCase().trim() === category.toLowerCase().trim()
  );

  return match ? match[1] : "admin@company.com"; 
}

function getUserSignature() {
  try {
    const me = Session.getActiveUser().getEmail();
    const draft = GmailApp.createDraft(me, "Signature Lookup", "");
    
    Utilities.sleep(1500);
    
    let html = draft.getMessage().getBody();
    draft.deleteDraft();
    
    if (html.length < 15) {
       return `<div><strong>JITENDRA KHATRI</strong><br>
               <a href="mailto:PCAVIATOR@GMAIL.COM">PCAVIATOR@GMAIL.COM</a><br>
               +56-9-41539249</div>`;
    }
    
    return html;
  } catch (e) {
    return "<div><strong>JITENDRA KHATRI</strong><br>+56-9-41539249</div>";
  }
}

function processFinalOrder(payload) {
  saveRequirementToSheet(payload);

  MailApp.sendEmail({
    to: payload.recipient,
    subject: payload.subject,
    htmlBody: payload.emailBody
  });

  return { success: true };
}

function getRequirementsHistory() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Ordenes"); 
    
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];

    const cleanData = data.slice(1).filter(row => row[0] !== "" && row[0] !== null);

    return cleanData.map(row => ({
      order: row[0] || "N/A",
      fecha: row[1] ? row[1].toLocaleString() : "",
      pn: row[2] || "",
      geo: row[3] || "",
      faena: row[4] || "",
      desc: row[5] || "",
      cant: row[6] || 0,
      estado: row[7] || "Pending"
    })).reverse(); 

  } catch (e) {
    console.error("Critical Server Error: " + e.message);
    return [];
  }
}

function generateExportFile(data, type) {
  try {
    const headers = ["Order #", "Date", "Faena", "Geo Code", "P/N", "Description", "Qty", "Status"];
    let content = headers.join(",") + "\n";
    
    data.forEach(r => {
      const cleanDesc = r.desc ? r.desc.replace(/"/g, '""') : "";
      content += `${r.order},${r.fecha},${r.faena},${r.geo},${r.pn},"${cleanDesc}",${r.cant},${r.estado}\n`;
    });

    const fileName = `RECON_${type.toUpperCase()}_${new Date().getTime()}.csv`;
    const file = DriveApp.createFile(fileName, content, MimeType.CSV);
    
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return file.getDownloadUrl().replace("?e=download", ""); 
  } catch (e) {
    throw new Error("Export failed: " + e.message);
  }
}

function master_ExportToDrive(data, headers, keys, moduleName, format) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
    
    const exportsFolder = shared_getOrCreateFolder(parentFolder, "Exports");
    const targetFolder = shared_getOrCreateFolder(exportsFolder, moduleName);

    const dateStr = Utilities.formatDate(new Date(), "GMT-4", "yyyy-MM-dd_HH-mm");
    const baseName = `${moduleName}_${dateStr}`;
    let blob;

    if (format === 'CSV') {
      let csvContent = "\uFEFF" + headers.join(",") + "\n";
      data.forEach(row => {
        const line = keys.map(key => `"${String(row[key] || "").replace(/"/g, '""')}"`);
        csvContent += line.join(",") + "\n";
      });
      blob = Utilities.newBlob(csvContent, MimeType.CSV, baseName + ".csv");
    } 
    else if (format === 'EXCEL') {
      blob = internal_generateExcelBlob(data, headers, keys, baseName);
    } 
    else if (format === 'PDF') {
      blob = internal_generatePdfBlob(data, headers, keys, moduleName, baseName);
    }

    const file = targetFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return { fileName: file.getName() };
  } catch (e) {
    throw new Error(e.toString());
  }
}

function internal_generatePdfBlob(data, headers, keys, moduleName, baseName) {
  const logoId = '1DqBJCUuFzontoeRt8Lh6yhLYifo60uRC'; 
  let logoBase64 = "";
  try {
    const logoBlob = DriveApp.getFileById(logoId).getBlob();
    logoBase64 = "data:" + logoBlob.getContentType() + ";base64," + Utilities.base64Encode(logoBlob.getBytes());
  } catch(e) { console.warn("Logo fetch failed"); }

  let dateRangeText = "All Dates";
  if (data.length > 0) {
    const dates = data.map(item => new Date(item.fecha)).filter(d => !isNaN(d));
    if (dates.length > 0) {
      const minDate = new Date(Math.min.apply(null, dates));
      const maxDate = new Date(Math.max.apply(null, dates));
      const fmt = (d) => Utilities.formatDate(d, "GMT-4", "dd/MM/yyyy");
      dateRangeText = fmt(minDate) === fmt(maxDate) ? fmt(minDate) : fmt(minDate) + " - " + fmt(maxDate);
    }
  }

  const timestamp = Utilities.formatDate(new Date(), "GMT-4", "dd/MM/yyyy HH:mm");

  let html = `
    <html>
    <head>
      <style>
        body { font-family: sans-serif; color: #333; margin: 20px; }
        .header { display: flex; border-bottom: 2px solid #444; padding-bottom: 10px; margin-bottom: 10px; }
        .logo { height: 50px; }
        .info { flex-grow: 1; text-align: right; }
        .date-range-bar { background: #f8f9fa; padding: 5px 10px; border: 1px solid #dee2e6; font-size: 10pt; margin-bottom: 15px; font-weight: bold; }
        h2 { margin: 0; color: #1a73e8; text-transform: uppercase; }
        table { width: 100%; border-collapse: collapse; }
        th { background-color: #1a73e8; color: white; border: 1px solid #bdc1c6; padding: 8px; font-size: 10pt; text-align: left; }
        td { border: 1px solid #bdc1c6; padding: 6px; font-size: 9pt; vertical-align: top; }
        .footer { margin-top: 30px; font-size: 8pt; color: #70757a; text-align: center; }
      </style>
    </head>
    <body>
      <div class="header">
        ${logoBase64 ? `<img src="${logoBase64}" class="logo">` : '<div></div>'}
        <div class="info">
          <h2>${moduleName} Report</h2>
          <p style="margin:0;">Generated: ${timestamp}</p>
        </div>
      </div>
      
      <div class="date-range-bar">
        Period: ${dateRangeText}
      </div>

      <table>
        <thead><tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead>
        <tbody>`;

  data.forEach(row => {
    html += `<tr>${keys.map(key => `<td>${row[key] || ""}</td>`).join('')}</tr>`;
  });

  html += `</tbody></table>
      <div class="footer">ERP System Internal Document - Confidential</div>
    </body></html>`;

  return HtmlService.createHtmlOutput(html).getAs('application/pdf').setName(baseName + ".pdf");
}

function shared_getOrCreateFolder(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function updateRequirementOnSheet(orderId, header, items) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Ordenes"); 
  if (!sheet) throw new Error("Sheet 'Ordenes' not found.");

  const fullData = sheet.getDataRange().getValues();
  const headers = fullData[0];
  const numColumns = headers.length; 

  const colIdx = {
    order:  headers.indexOf("Order"),
    date:   headers.indexOf("Fecha/Hora"),
    pn:     headers.indexOf("P/N"),
    geo:    headers.indexOf("Geo Code"),
    site:   headers.indexOf("Faena"),
    desc:   headers.indexOf("Desc Es"),
    qty:    headers.indexOf("Cantidad"),
    status: headers.indexOf("Estado")
  };

  for (let key in colIdx) {
    if (colIdx[key] === -1) {
      throw new Error("Could not find column header for: " + key + ". Check for hidden spaces in Row 1.");
    }
  }

  for (let i = fullData.length - 1; i >= 1; i--) {
    const rowOrder = fullData[i][colIdx.order];
    if (rowOrder && rowOrder.toString() === orderId.toString()) {
      sheet.deleteRow(i + 1);
    }
  }

  const cleanItems = items.filter(item => item.geo && item.geo.trim() !== "");

  const newRows = cleanItems.map(item => {
    let row = new Array(numColumns).fill(""); 
    
    row[colIdx.order]  = orderId;
    row[colIdx.date]   = new Date(); 
    row[colIdx.site]   = header.faena; 
    row[colIdx.geo]    = item.geo; 
    row[colIdx.pn]     = item.pn || ""; 
    row[colIdx.desc]   = item.desc || "";
    row[colIdx.qty]    = item.cant;
    row[colIdx.status] = "Pending";
    
    return row;
  });

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, numColumns).setValues(newRows);
  }
  
  return { success: true };
}

function updateInventoryItem(item) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Master");
  const data = sheet.getDataRange().getValues();
  const geoCode = item.geo; // Changed from item.g
  
  let rowIndex = -1;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == geoCode) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) rowIndex = data.length + 1;
  
  const rowValues = [
    item.geo,          // A - Changed from item.g
    item.pn,           // B - Changed from item.p
    item.category,     // C
    item.sub_categoria, // D
    item.marca,        // E
    item.en,           // F
    item.es,           // G
    item.un,           // H
    item.min,          // I
    item.max,          // J
    "",                // K
    "",                // L
    item.costo_clp,    // M - Changed from item.costo
    item.proveedor,    // N
    item.eta           // O
  ];
  
  sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  Logger.log("✅ Inventory item updated: " + geoCode);
  return { success: true, row: rowIndex };
}

function getPendingRequirementsForRFQ(site) {
  const allHistory = getRequirementsHistory();
  const masterData = getInventoryData();
  
  const masterMap = new Map();
  masterData.forEach(item => {
    masterMap.set(item.pn, { en: item.en, es: item.es });
  });

  let pending = allHistory.filter(item => item.estado && item.estado.toString().toLowerCase().includes('pend'));
  if (site) {
    pending = pending.filter(item => item.faena === site);
  }

  const groupedReqs = {};
  pending.forEach(item => {
    if (!groupedReqs[item.order]) {
      groupedReqs[item.order] = {
        id: item.order,
        date: item.fecha,
        site: item.faena,
        items: []
      };
    }
    
    const official = masterMap.get(item.pn) || { en: item.desc, es: item.desc };
    
    groupedReqs[item.order].items.push({
      p: item.pn,
      geo: item.geo,
      en: official.en,
      es: official.es,
      qty: item.cant,
      supplier: ""
    });
  });

  Logger.log("✅ Loaded " + Object.keys(groupedReqs).length + " pending RFQ orders");
  return Object.values(groupedReqs);
}

function getItemsByOrderId(orderId) {
  const allHistory = getRequirementsHistory();
  return allHistory.filter(item => item.order === orderId);
}

function sendFinalRFQ(emailData) {
  try {
    if (!emailData.to || !emailData.subject || !emailData.body) {
      throw new Error("Missing email fields: to, subject, or body");
    }
    
    const signature = getUserSignature();
    
    const finalHtml = `
      <div style="font-family: Arial, sans-serif; color: #333; line-height: 1.6;">
        ${emailData.body}
        <br><br>
        ${signature}
      </div>
    `;
    
    MailApp.sendEmail({
      to: emailData.to,
      subject: emailData.subject,
      htmlBody: finalHtml
    });
    
    Logger.log("✅ Email sent successfully to: " + emailData.to);
    
    // Update sheet status
    updateQuotedStatus(emailData.orderId, emailData.supplier);
    
    return { success: true, message: "Email sent successfully!" };
    
  } catch(e) {
    Logger.log("❌ Error in sendFinalRFQ: " + e.message);
    return { success: false, message: e.message };
  }
}

function updateQuotedStatus(orderId, supplierName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Ordenes");
    
    if (!sheet) {
      throw new Error('Sheet "Ordenes" not found');
    }
    
    const data = sheet.getDataRange().getValues();
    
    const timestamp = Utilities.formatDate(new Date(), "GMT-4", "dd/MM HH:mm");
    const newStatus = "Quoted: " + timestamp;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const currentOrder = row[0];
      const currentStatus = row[7] ? row[7].toString().toLowerCase() : '';
      
      if (currentOrder == orderId && currentStatus.includes('pend')) {
        sheet.getRange(i + 1, 8).setValue(newStatus);
        Logger.log(`✅ Updated Order #${orderId} status to: ${newStatus}`);
      }
    }
    
  } catch(e) {
    Logger.log("⚠️ Warning: Could not update sheet status: " + e.message);
  }
}