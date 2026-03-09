const SPREADSHEET_ID = "1d5eFnUM65lzpaJUJXSdpZG0BmK70SCf3TF5K4tljCZw";
const SHEET_NAME = "Database";

/* ================= WEB APP ================= */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle("Monitoring Rekon Bitrans")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ================= HELPER ================= */

function cleanHeader(h){
  return String(h).trim().toLowerCase();
}

function cleanText(v){
  if(!v) return "";
  return String(v).replace(/\s+/g,' ').trim();
}

/* ================= UNIVERSAL DATE PARSER ================= */

function parseDate(val){

  if(!val) return null;

  if(Object.prototype.toString.call(val) === "[object Date]")
    return val;

  // dd/mm/yyyy
  if(typeof val === "string" && val.includes("/")){
    const p = val.split("/");
    if(p.length===3)
      return new Date(p[2], p[1]-1, p[0]);
  }

  // yyyy-mm-dd
  if(typeof val === "string" && val.includes("-")){
    const d = new Date(val);
    if(!isNaN(d)) return d;
  }

  // excel serial
  if(typeof val === "number"){
    return new Date(Math.round((val-25569)*86400*1000));
  }

  return null;
}

/* ================= MAIN FUNCTION ================= */

function getMonitoringData() {

  try{

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" tidak ditemukan`);

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if(lastRow <= 1)
      return {monitoring:[],lastUpdate:"-"};

    const values = sheet.getRange(1,1,lastRow,lastCol).getValues();

    const rawHeader = values.shift();
    const header = rawHeader.map(cleanHeader);

    // ===== MAPPING KOLOM =====
    const colOP = header.indexOf("kode op");
    const colNamaOP = header.indexOf("nama op");
    const colCustomer = header.indexOf("customer");
    const colProduk = header.indexOf("produk"); // <<<<<< BARU
    const colTOH = header.indexOf("nama toh/woh");
    const colDate = header.indexOf("tanggal sla");
    const colStatus = header.indexOf("status rekon");
    const colTanggalUpdate = header.indexOf("tanggal update");

    if (colOP === -1 || colNamaOP === -1 || colCustomer === -1 || colDate === -1 || colStatus === -1){
      throw new Error("Header wajib tidak ditemukan. Pastikan ada: Kode OP, Nama OP, Customer, Produk, Tanggal SLA, Status Rekon");
    }

    let data=[];
    let lastUpdate=null;

    values.forEach(r=>{

      const op = cleanText(r[colOP]);
      const namaop = cleanText(r[colNamaOP]);
      const customer = cleanText(r[colCustomer]);
      const produk = colProduk>-1 ? cleanText(r[colProduk]) : "-";  // <<<<<< PRODUK

      if(!op || !r[colDate]) return;

      const d=parseDate(r[colDate]);
      if(!d) return;

      let status=cleanText(r[colStatus]);
      if(!status) status="Non Rekon";

      const monthKey = d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0");

      data.push({
        op: op,
        namaop: namaop,
        customer: customer,
        produk: produk,        // <<<<<< DIKIRIM KE HTML
        toh: (colTOH>-1 && r[colTOH]) ? cleanText(r[colTOH]) : "-",
        month: monthKey,
        day: String(d.getDate()),
        status: status
      });

      // ===== LAST UPDATE =====
      if(colTanggalUpdate>-1 && r[colTanggalUpdate]){
        const u=parseDate(r[colTanggalUpdate]);
        if(u && (!lastUpdate || u>lastUpdate))
          lastUpdate=u;
      }

    });

    return {
      monitoring:data,
      lastUpdate: lastUpdate
        ? Utilities.formatDate(lastUpdate, Session.getScriptTimeZone(),"dd MMM yyyy HH:mm")
        : "-"
    };

  }catch(err){

    Logger.log(err);

    return {
      error:err.toString(),
      monitoring:[],
      lastUpdate:"-"
    };
  }

}

// AUTO KIRIM EMAIL TOH/WOH

function sendDailyNonRekonEmail(){

  const DASHBOARD_URL = "https://script.google.com/macros/s/AKfycbwkI628xKLAchR69opeYdGp4gdTYmwS45XaTJjN53-bHr2sHj9B8XmnKirgSnlllsDF8A/exec";

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const db = ss.getSheetByName("Database");
  const master = ss.getSheetByName("Master_OP_OH");

  if(!db) throw new Error("Sheet Database tidak ditemukan");
  if(!master) throw new Error("Sheet Master_OP_OH tidak ditemukan");

  const dbData = db.getDataRange().getValues();
  const dbHeader = dbData.shift();

  const masterData = master.getDataRange().getValues();
  const masterHeader = masterData.shift();

  const colOP       = dbHeader.indexOf("Kode OP");
  const colNamaOP   = dbHeader.indexOf("Nama OP");
  const colCustomer = dbHeader.indexOf("Customer");
  const colProduk   = dbHeader.indexOf("Produk");
  const colTanggal  = dbHeader.indexOf("Tanggal SLA");
  const colStatus   = dbHeader.indexOf("Status Rekon");
  const colTOH      = dbHeader.indexOf("Nama TOH/WOH");

  const colMasterTOH = masterHeader.indexOf("Nama TOH/WOH");
  const colEmail     = masterHeader.indexOf("Email TOH");

  if(colEmail === -1)
    throw new Error("Kolom 'Email TOH' belum ada di Master_OP_OH");

  const timezone = Session.getScriptTimeZone();
  const todayKey = Utilities.formatDate(new Date(), timezone, "yyyy-MM-dd");


  /* =========================================================
     1. EMAIL MAPPING
     ========================================================= */

  let emailMap = {};

  masterData.forEach(r=>{
    if(r[colMasterTOH] && r[colEmail]){
      const name = r[colMasterTOH].toString().trim().toLowerCase();
      emailMap[name] = r[colEmail].toString().trim();
    }
  });

  /* =========================================================
     2. CARI SLA TERAKHIR
     ========================================================= */

  let lastSLA = null;

  dbData.forEach(r=>{
    if(!r[colTanggal]) return;
    let d = new Date(r[colTanggal]);
    if(isNaN(d)) return;
    if(!lastSLA || d > lastSLA) lastSLA = d;
  });

  if(!lastSLA){
    Logger.log("Tidak ada tanggal SLA ditemukan");
    return;
  }

  /* =========================================================
     3. VALIDASI MAX 3 BULAN KE BELAKANG
     ========================================================= */

  let monthsToCheck = [];

  for(let i=0;i<3;i++){
    let d = new Date(lastSLA.getFullYear(), lastSLA.getMonth()-i, 1);
    monthsToCheck.push(
      Utilities.formatDate(d, timezone, "yyyy-MM")
    );
  }

  const bulanIndo = [
    "Januari","Februari","Maret","April","Mei","Juni",
    "Juli","Agustus","September","Oktober","November","Desember"
  ];

  /* =========================================================
     4. KUMPULKAN DATA
     ========================================================= */

  let laporan = {}; // laporan[monthKey][toh][key]

  dbData.forEach(r=>{

    if(!r[colTanggal]) return;

    let tgl = new Date(r[colTanggal]);
    if(isNaN(tgl)) return;

    const monthKey = Utilities.formatDate(tgl, timezone, "yyyy-MM");
    if(!monthsToCheck.includes(monthKey)) return;

    if(String(r[colStatus]).trim().toLowerCase() === "rekon") return;

    const toh = String(r[colTOH]).trim();
    if(!toh) return;

    const key = r[colOP] + "||" + r[colCustomer] + "||" + r[colProduk];

    if(!laporan[monthKey]) laporan[monthKey] = {};
    if(!laporan[monthKey][toh]) laporan[monthKey][toh] = {};
    if(!laporan[monthKey][toh][key]){
      laporan[monthKey][toh][key] = {
        op: r[colOP],
        nama: r[colNamaOP],
        cust: r[colCustomer],
        produk: r[colProduk],
        tanggal:[]
      };
    }

    const day = tgl.getDate();
    if(!laporan[monthKey][toh][key].tanggal.includes(day))
      laporan[monthKey][toh][key].tanggal.push(day);

  });

  /* =========================================================
     5. KIRIM EMAIL
     ========================================================= */

  let totalEmailSent = 0;

  Object.keys(laporan).forEach(monthKey=>{

    const [year,month] = monthKey.split("-");
    const periodeBulan = bulanIndo[parseInt(month)-1] + " " + year;

    Object.keys(laporan[monthKey]).forEach(toh=>{

      const email = emailMap[toh.toLowerCase()];
      if(!email) return;

      let rows="";

      Object.values(laporan[monthKey][toh]).forEach(d=>{

        const tanggalList = d.tanggal.sort((a,b)=>a-b).join(", ");

        rows += `
        <tr>
          <td>${d.op}</td>
          <td>${d.nama}</td>
          <td>${d.cust}</td>
          <td>${d.produk}</td>
          <td style="color:#c0392b;font-weight:bold">${tanggalList}</td>
        </tr>`;
      });

      if(!rows) return;

      const html=`
      <div style="font-family:Arial,sans-serif">
        <h2 style="color:#c0392b;">⚠ Monitoring Non-Rekon Bitrans</h2>
        <p>Yth. <b>${toh}</b>,</p>

        <p>
        Berikut daftar <b>Non-Rekon Periode ${periodeBulan}</b>
        yang masih belum terselesaikan:
        </p>

        <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse">
          <tr style="background:#2c3e50;color:white">
            <th>Kode OP</th>
            <th>Nama OP</th>
            <th>Customer</th>
            <th>Produk</th>
            <th>Tanggal Belum Rekon</th>
          </tr>
          ${rows}
        </table>

        <br><br>

        <a href="${DASHBOARD_URL}" 
           style="background:#2980b9;color:white;
                  padding:10px 18px;
                  text-decoration:none;
                  border-radius:5px;
                  display:inline-block;">
           🔎 Buka Dashboard Monitoring
        </a>

        <br><br>
        <small>Email otomatis Monitoring Rekon</small>
        <br><br>
        <p>Regards, Team TMS</p>
      </div>
      `;

      MailApp.sendEmail({
        to: email,
        subject: "⚠ Non-Rekon Bitrans Periode "+periodeBulan,
        htmlBody: html
      });

      totalEmailSent++;

    });

  });

  /* =========================================================
     6. SIMPAN STATUS JIKA ADA EMAIL TERKIRIM
     ========================================================= */

  Logger.log("Total email terkirim: " + totalEmailSent);

}

// DOWNLOAD REPORT EXCEL

function downloadNonRekonReport(selectedMonth){

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const db = ss.getSheetByName("Database");
  const data = db.getDataRange().getValues();
  const header = data.shift();

  const colOP       = header.indexOf("Kode OP");
  const colNamaOP   = header.indexOf("Nama OP");
  const colCustomer = header.indexOf("Customer");
  const colProduk   = header.indexOf("Produk");
  const colTanggal  = header.indexOf("Tanggal SLA");
  const colStatus   = header.indexOf("Status Rekon");
  const colTOH      = header.indexOf("Nama TOH/WOH");

  const timezone = Session.getScriptTimeZone();

  let output = [
    ["Kode OP","Nama OP","Customer","Produk","TOH","Tanggal SLA","Status"]
  ];

  data.forEach(r => {

    if(!r[colTanggal]) return;

    let tgl = new Date(r[colTanggal]);
    if(isNaN(tgl)) return;

    const monthKey = Utilities.formatDate(tgl, timezone, "MMMM yyyy");
    if(monthKey !== selectedMonth) return;

    if(String(r[colStatus]).trim().toLowerCase() === "rekon") return;

    output.push([
      r[colOP],
      r[colNamaOP],
      r[colCustomer],
      r[colProduk],
      r[colTOH],
      Utilities.formatDate(tgl, timezone, "dd-MM-yyyy"),
      r[colStatus]
    ]);
  });

  if(output.length <= 1){
    throw new Error("Tidak ada data Non-Rekon pada bulan tersebut");
  }

  const fileName = "Rekap_Non_Rekon_" + selectedMonth.replace(" ","_");

  const blob = Utilities.newBlob(
    output.map(r => r.join("\t")).join("\n"),
    "application/vnd.ms-excel",
    fileName + ".xls"
  );

  return blob;
}

// FUNGSI TOMBOL DOWNLOAD REPORT

function downloadNonRekonReport(selectedMonth){

  if(!selectedMonth){
    throw new Error("Bulan tidak dipilih.");
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const db = ss.getSheetByName("Database");
  const data = db.getDataRange().getValues();
  const header = data.shift();

  // ===== SAFE HEADER FINDER (anti beda spasi / case) =====
  function findColumn(name){
    return header.findIndex(h =>
      String(h).trim().toLowerCase() === name.toLowerCase()
    );
  }

  const colOP       = findColumn("Kode OP");
  const colNamaOP   = findColumn("Nama OP");
  const colCustomer = findColumn("Customer");
  const colProduk   = findColumn("Produk");
  const colTanggal  = findColumn("Tanggal SLA");
  const colStatus   = findColumn("Status Rekon");
  const colTOH      = findColumn("Nama TOH/WOH");

  if([colOP,colNamaOP,colCustomer,colProduk,colTanggal,colStatus,colTOH].includes(-1)){
    throw new Error("Header kolom tidak sesuai dengan database.");
  }

  const timezone = Session.getScriptTimeZone();

  // ===== NORMALISASI BULAN (anti 2025-3 vs 2025-03) =====
  function normalizeMonth(m){
    const parts = String(m).split("-");
    if(parts.length !== 2) return m;
    return parts[0] + "-" + parts[1].padStart(2,"0");
  }

  selectedMonth = normalizeMonth(selectedMonth);

  let output = [
    ["Kode OP","Nama OP","Customer","Produk","TOH","Tanggal SLA","Status"]
  ];

  data.forEach(r => {

    if(!r[colTanggal]) return;

    let tgl = new Date(r[colTanggal]);
    if(isNaN(tgl)) return;

    const monthKey =
      Utilities.formatDate(tgl, timezone, "yyyy-MM");

    if(monthKey !== selectedMonth) return;

    if(String(r[colStatus]).trim().toLowerCase() === "rekon") return;

    output.push([
      r[colOP],
      r[colNamaOP],
      r[colCustomer],
      r[colProduk],
      r[colTOH],
      Utilities.formatDate(tgl, timezone, "dd-MM-yyyy"),
      r[colStatus]
    ]);
  });

  if(output.length <= 1){
    throw new Error("Tidak ada data Non-Rekon pada bulan tersebut.");
  }

  // ===============================
  // BUAT FILE SEMENTARA
  // ===============================

  const tempSS = SpreadsheetApp.create("TEMP_EXPORT");
  const sheet = tempSS.getActiveSheet();

  sheet.getRange(1,1,output.length,output[0].length).setValues(output);

  // Formatting biar rapi
  sheet.getRange("A1:G1").setFontWeight("bold");
  sheet.autoResizeColumns(1,7);
  sheet.setFrozenRows(1);

  SpreadsheetApp.flush();
  Utilities.sleep(500); // pastikan data sudah tertulis

  const fileId = tempSS.getId();

  // ===============================
  // EXPORT XLSX ASLI
  // ===============================

  const url = "https://www.googleapis.com/drive/v3/files/"
              + fileId
              + "/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

  const token = ScriptApp.getOAuthToken();

  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: "Bearer " + token
    },
    muteHttpExceptions: true
  });

  if(response.getResponseCode() !== 200){
    DriveApp.getFileById(fileId).setTrashed(true);
    throw new Error("Gagal export XLSX. Pastikan Drive API aktif.");
  }

  const blob = response.getBlob()
    .setName("Rekap_Non_Rekon_" + selectedMonth + ".xlsx");

  // HAPUS FILE SEMENTARA
  DriveApp.getFileById(fileId).setTrashed(true);

  return {
    fileName: blob.getName(),
    mimeType: blob.getContentType(),
    data: Utilities.base64Encode(blob.getBytes())
  };
}
