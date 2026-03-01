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

/* =========================================================
   AUTO EMAIL NON REKON BITRANS (FINAL)
   ========================================================= */

function sendDailyNonRekonEmail(){

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const db = ss.getSheetByName("Database");
  const master = ss.getSheetByName("Master_OP_OH");

  if(!db) throw new Error("Sheet Database tidak ditemukan");
  if(!master) throw new Error("Sheet Master_OP_OH tidak ditemukan");

  const dbData = db.getDataRange().getValues();
  const dbHeader = dbData.shift();

  const masterData = master.getDataRange().getValues();
  const masterHeader = masterData.shift();

  /* ================= INDEX KOLOM DATABASE ================= */

  const colOP       = dbHeader.indexOf("Kode OP");
  const colNamaOP   = dbHeader.indexOf("Nama OP");
  const colCustomer = dbHeader.indexOf("Customer");
  const colProduk   = dbHeader.indexOf("Produk");
  const colTanggal  = dbHeader.indexOf("Tanggal SLA");
  const colStatus   = dbHeader.indexOf("Status Rekon");
  const colTOH      = dbHeader.indexOf("Nama TOH/WOH");

  /* ================= INDEX MASTER ================= */

  const colMasterTOH = masterHeader.indexOf("Nama TOH/WOH");
  const colEmail     = masterHeader.indexOf("Email TOH");

  if(colEmail === -1)
    throw new Error("Kolom 'Email TOH' belum ada di Master_OP_OH");

  /* =========================================================
     1. MAPPING TOH -> EMAIL (ANTI HURUF BESAR KECIL)
     ========================================================= */

  let emailMap = {};

  masterData.forEach(r=>{
    if(r[colMasterTOH] && r[colEmail]){
      const name = r[colMasterTOH].toString().trim().toLowerCase();
      emailMap[name] = r[colEmail].toString().trim();
    }
  });

  Logger.log("Total TOH dengan email: "+Object.keys(emailMap).length);

  /* =========================================================
     2. CARI TANGGAL SLA TERAKHIR (INI YANG PALING PENTING)
     ========================================================= */

  let lastSLA = null;

  dbData.forEach(r=>{
    if(!r[colTanggal]) return;

    let d = new Date(r[colTanggal]);
    if(isNaN(d)) return;

    if(!lastSLA || d > lastSLA)
      lastSLA = d;
  });

  if(!lastSLA){
    Logger.log("Tidak ada tanggal SLA ditemukan");
    return;
  }

  const targetDate = Utilities.formatDate(lastSLA, Session.getScriptTimeZone(),"yyyy-MM-dd");

  Logger.log("Tanggal SLA target (yang dicek): "+targetDate);

  /* =========================================================
     3. KUMPULKAN NON REKON PER TOH
     ========================================================= */

  let laporan = {};

dbData.forEach(r=>{

  if(!r[colTanggal]) return;

  let tgl = new Date(r[colTanggal]);
  if(isNaN(tgl)) return;

  // hanya tanggal SLA terakhir (bulan aktif)
  const monthCheck = Utilities.formatDate(tgl, Session.getScriptTimeZone(),"yyyy-MM");
  const targetMonth = Utilities.formatDate(lastSLA, Session.getScriptTimeZone(),"yyyy-MM");

  if(monthCheck !== targetMonth) return;

  // hanya NON REKON
  if(String(r[colStatus]).trim().toLowerCase() === "rekon") return;

  const toh = String(r[colTOH]).trim();
  if(!toh) return;

  const key = r[colOP] + "||" + r[colCustomer] + "||" + r[colProduk];

  if(!laporan[toh]) laporan[toh] = {};
  if(!laporan[toh][key]){
    laporan[toh][key] = {
      op: r[colOP],
      nama: r[colNamaOP],
      cust: r[colCustomer],
      produk: r[colProduk],
      tanggal:[]
    };
  }

  const day = tgl.getDate();
  if(!laporan[toh][key].tanggal.includes(day))
    laporan[toh][key].tanggal.push(day);

});

  /* =========================================================
     4. KIRIM EMAIL
     ========================================================= */

  Object.keys(laporan).forEach(toh=>{

    const email = emailMap[toh.toLowerCase()];

    if(!email){
      Logger.log("Email tidak ditemukan untuk TOH: "+toh);
      return;
    }

    let rows="";

Object.values(laporan[toh]).forEach(d=>{

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

    const html=`
    <div style="font-family:Arial,sans-serif">

      <h2 style="color:#c0392b;">⚠ Monitoring Non-Rekon Bitrans</h2>

      <p>Yth. <b>${toh}</b>,</p>

      <p>Berikut daftar <b>Non-Rekon tanggal ${targetDate}</b> update terbaru yang perlu diperhatikan:</p>

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

      <br>
      <b style="color:#e67e22">
      Mohon segera dilakukan pengecekan & monitoring terhadap Operating Point Tersebut.
      </b>

      <br><br>
      <small>Ini adalah email otomatis dari Sistem Monitoring Rekon.</small>
      <br><br>
      <p>Regards, Team TMS</p>

    </div>
    `;

    MailApp.sendEmail({
      to: email,
      subject: "⚠ Non-Rekon Harian Bitrans ("+targetDate+")",
      htmlBody: html
    });

    Logger.log("Email terkirim ke: "+email);

  });

}