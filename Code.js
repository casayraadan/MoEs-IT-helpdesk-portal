
/* =========================================================
   MoES IT HELP DESK & INNOVATION PORTAL
   Backend: Google Apps Script
   Database: Google Sheets
   ========================================================= */

/* SECTION 1: CONFIGURATION */
const CONFIG = {
  SYSTEM_NAME: "MoES IT Help Desk",
  ADMIN_EMAIL: "casayraadan15@gmail.com",
  SPREADSHEET_ID: "SETUP_AUTOMATICALLY",
  SESSION_SECONDS: 21600
};

const SHEETS = {
  USERS: "Users",
  TICKETS: "Requests",
  IDEAS: "Suggestions",
  LOGS: "ActivityLog"
};

/* SECTION 2: INITIAL SETUP */
function setupSystem() {
  const props = PropertiesService.getScriptProperties();
  let dbId = props.getProperty("MOES_DB_ID");
  let ss = null;

  // Reuse the existing database so setup cannot accidentally create
  // a new empty spreadsheet every time it is run.
  if (dbId) {
    try {
      ss = SpreadsheetApp.openById(dbId);
    } catch (err) {
      throw new Error(
        "Database ID-ga kaydsan lama furi karo. Hubi rukhsadaha ama MOES_DB_ID."
      );
    }
  } else {
    ss = SpreadsheetApp.create("MoES IT Help Desk Database");
    dbId = ss.getId();
    props.setProperty("MOES_DB_ID", dbId);
  }

  createSheet_(ss, SHEETS.USERS, [
    "Email", "FullName", "Department", "PasswordHash",
    "Salt", "Role", "CreatedAt", "Active"
  ]);

  createSheet_(ss, SHEETS.TICKETS, [
    "TicketID", "CreatedAt", "Email", "SenderName",
    "Department", "Title", "Category", "Priority",
    "Description", "Status", "AssignedTo",
    "ITResponse", "UpdatedAt"
  ]);

  createSheet_(ss, SHEETS.IDEAS, [
    "IdeaID", "CreatedAt", "Email", "SenderName",
    "Department", "Title", "IdeaType", "Description",
    "Status", "ITResponse"
  ]);

  createSheet_(ss, SHEETS.LOGS, [
    "Date", "Email", "Action", "RecordID"
  ]);

  Logger.log("Database ready: " + ss.getUrl());
  Logger.log("Admin email: " + CONFIG.ADMIN_EMAIL);
}

function createSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }

  // Only add headers to a genuinely empty sheet. Never clear existing records.
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const existing = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    const matches = headers.every((header, i) => existing[i] === header);
    if (!matches) {
      throw new Error(
        'Qaab-dhismeedka sheet-ka "' + name +
        '" wuu ka duwan yahay kii la filayey. Xogta lama tirtirin; hubi headers-ka.'
      );
    }
  }

  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#166534")
    .setFontColor("#ffffff");
  sh.autoResizeColumns(1, headers.length);
}

/* SECTION 3: WEB APP */
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle(CONFIG.SYSTEM_NAME)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function db_() {
  const id = PropertiesService.getScriptProperties()
    .getProperty("MOES_DB_ID");
  if (!id) throw new Error("Run setupSystem() first.");
  return SpreadsheetApp.openById(id);
}

function sheet_(name) {
  const sh = db_().getSheetByName(name);
  if (!sh) {
    throw new Error(
      'Sheet-ka "' + name + '" lama helin. Hubi inaad hal mar ordisay setupSystem().'
    );
  }
  return sh;
}

function rows_(name) {
  const data = sheet_(name).getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1).map(row => {
    const obj = {};
    data[0].forEach((key, i) => obj[key] = row[i]);
    return obj;
  });
}

function clean_(value, max = 2000) {
  return String(value || "").trim().slice(0, max);
}

function email_(value) {
  return clean_(value, 254).toLowerCase();
}

/**
 * Convert spreadsheet date values to strings before returning data through
 * google.script.run. Apps Script cannot serialize Date objects as return values.
 */
function safeDate_(value) {
  if (value === null || value === undefined || value === "") return "";

  if (value instanceof Date) {
    if (isNaN(value.getTime())) return "";
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd HH:mm:ss"
    );
  }

  return String(value);
}

function hashPassword_(password, salt) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + ":" + password,
    Utilities.Charset.UTF_8
  );
  return Utilities.base64Encode(bytes);
}

function log_(email, action, id) {
  sheet_(SHEETS.LOGS).appendRow([
    new Date(), email, action, id || ""
  ]);
}

/* SECTION 4: REGISTRATION */
function registerUser(form) {
  const name = clean_(form.name, 100);
  const email = email_(form.email);
  const department = clean_(form.department, 120);
  const password = String(form.password || "");

  if (!name || !department ||
      !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error("Geli magaca, email sax ah iyo waaxda.");
  }

  if (password.length < 10) {
    throw new Error("Password-ku waa inuu ahaadaa ugu yaraan 10 xaraf.");
  }

  const users = rows_(SHEETS.USERS);
  if (users.some(u => email_(u.Email) === email)) {
    throw new Error("Email-kan horay ayaa loo diiwaangeliyey.");
  }

  const salt = Utilities.getUuid();
  const role = email === email_(CONFIG.ADMIN_EMAIL)
    ? "ADMIN" : "USER";

  sheet_(SHEETS.USERS).appendRow([
    email, name, department,
    hashPassword_(password, salt), salt,
    role, new Date(), true
  ]);

  log_(email, "REGISTER", "");
  return { success: true, message: "Diiwaangelintu way guulaysatay. Hadda login samee." };
}

/* SECTION 5: LOGIN AND SESSION */
function loginUser(form) {
  const email = email_(form.email);
  const password = String(form.password || "");
  const user = rows_(SHEETS.USERS).find(
    u => email_(u.Email) === email && u.Active === true
  );

  if (!user || hashPassword_(password, user.Salt) !== user.PasswordHash) {
    throw new Error("Email ama password ayaa khaldan.");
  }

  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put(
    "session_" + token,
    JSON.stringify({
      email: user.Email,
      role: user.Role,
      name: user.FullName
    }),
    CONFIG.SESSION_SECONDS
  );

  log_(email, "LOGIN", "");
  return {
    token,
    user: {
      email: user.Email,
      name: user.FullName,
      department: user.Department,
      role: user.Role
    }
  };
}

function auth_(token) {
  if (!token) throw new Error("Fadlan dib u login samee.");
  const saved = CacheService.getScriptCache()
    .get("session_" + token);
  if (!saved) throw new Error("Session-ku wuu dhacay. Dib u login samee.");

  const session = JSON.parse(saved);
  const user = rows_(SHEETS.USERS).find(
    u => email_(u.Email) === email_(session.email) &&
         u.Active === true
  );
  if (!user) throw new Error("Akoonka lama heli karo ama waa la xiray.");

  return {
    email: user.Email,
    name: user.FullName,
    department: user.Department,
    role: user.Role
  };
}

function logoutUser(token) {
  auth_(token);
  CacheService.getScriptCache().remove("session_" + token);
  return { success: true };
}

/* SECTION 6: SUBMIT IT REQUEST */
function submitRequest(token, form) {
  const user = auth_(token);
  if (!form || typeof form !== "object") {
    throw new Error("Xogta codsiga lama helin. Fadlan mar kale isku day.");
  }
  const title = clean_(form.title, 180);
  const category = clean_(form.category, 60);
  const description = clean_(form.description, 4000);
  const priority = clean_(form.priority, 20);

  if (!title || !category || description.length < 10) {
    throw new Error("Buuxi cinwaanka, qaybta iyo faahfaahin ku filan.");
  }

  const allowed = ["Low", "Normal", "High", "Urgent"];
  if (!allowed.includes(priority)) {
    throw new Error("Mudnaanta codsigu ma saxna.");
  }

  const id = "IT-" + Utilities.getUuid().slice(0, 8).toUpperCase();
  const now = new Date();

  sheet_(SHEETS.TICKETS).appendRow([
    id, now, user.email, user.name, user.department,
    title, category, priority, description,
    "New", "", "", now
  ]);

  log_(user.email, "SUBMIT_REQUEST", id);
  return { success: true, id };
}

/* SECTION 7: SUBMIT IDEAS / SUGGESTIONS */
function submitIdea(token, form) {
  const user = auth_(token);
  if (!form || typeof form !== "object") {
    throw new Error("Xogta fikradda lama helin. Fadlan mar kale isku day.");
  }
  const title = clean_(form.title, 180);
  const type = clean_(form.type, 60);
  const description = clean_(form.description, 4000);

  if (!title || !type || description.length < 10) {
    throw new Error("Buuxi cinwaanka, nooca fikradda iyo faahfaahinta.");
  }

  const id = "IDEA-" + Utilities.getUuid().slice(0, 8).toUpperCase();

  sheet_(SHEETS.IDEAS).appendRow([
    id, new Date(), user.email, user.name, user.department,
    title, type, description, "New", ""
  ]);

  log_(user.email, "SUBMIT_IDEA", id);
  return { success: true, id };
}


/* SECTION 8: DASHBOARD DATA */
function getDashboard(token) {
  try {
    // 1. Hubi user-ka
    const user =auth_(token);

    if (!user || !user.email) {
      throw new Error(
        "Akoonka lama xaqiijin. Fadlan dib u login samee."
      );
    }

    const admin = ["ADMIN", "IT_STAFF"].includes(
      String(user.role || "").trim().toUpperCase()
    );

    // 2. Hubi xogta tickets iyo ideas
    const ticketsData = rows_(SHEETS.TICKETS);
    const ideasData = rows_(SHEETS.IDEAS);

    let tickets = Array.isArray(ticketsData)
      ? ticketsData
      : [];

    let ideas = Array.isArray(ideasData)
      ? ideasData
      : [];

    // 3. User-ka caadiga ahi wuxuu arkaa xogtiisa oo keliya
    if (!admin) {
      const userEmail = email_(user.email);

      tickets = tickets.filter(
        t => email_(t.Email) === userEmail
      );

      ideas = ideas.filter(
        i => email_(i.Email) === userEmail
      );
    }

    // 4. Tirakoobka dashboard-ka
    const statuses = {};
    const categories = {};

    tickets.forEach(t => {
      const status = String(t.Status || "Unknown").trim();
      const category = String(t.Category || "Other").trim();

      statuses[status] = (statuses[status] || 0) + 1;
      categories[category] = (categories[category] || 0) + 1;
    });

    // 5. U diyaari xogta dashboard-ka
    return {
      user: {
        name: user.name || "",
        email: user.email,
        department: user.department || "",
        role: String(user.role || "").trim().toUpperCase()
      },

      total: tickets.length,

      newCount: tickets.filter(
        t => String(t.Status || "").trim() === "New"
      ).length,

      progressCount: tickets.filter(
        t => String(t.Status || "").trim() === "In Progress"
      ).length,

      resolvedCount: tickets.filter(
        t => String(t.Status || "").trim() === "Resolved"
      ).length,

      statuses: statuses,
      categories: categories,

      tickets: tickets.slice(-200).reverse().map(t => ({
        id: t.TicketID || "",
        date: safeDate_(t.CreatedAt),
        name: t.SenderName || "",
        department: t.Department || "",
        title: t.Title || "",
        category: t.Category || "",
        priority: t.Priority || "",
        status: t.Status || "",
        response: t.ITResponse || ""
      })),

      ideas: ideas.slice(-100).reverse().map(i => ({
        id: i.IdeaID || "",
        date: safeDate_(i.CreatedAt),
        title: i.Title || "",
        type: i.IdeaType || "",
        status: i.Status || ""
      }))
    };

  } catch (error) {
    console.error(
      "getDashboard error:",
      error && error.message ? error.message : String(error)
    );
    throw new Error(
      "Dashboard-ka xogtiisa lama soo akhrin karin: " +
      (error && error.message ? error.message : String(error))
    );
  }
}


/* SECTION 9: ADMIN / IT STAFF ACTIONS */
function updateRequest(token, request) {
  const user = auth_(token);
  if (!["ADMIN", "IT_STAFF"].includes(user.role)) {
    throw new Error("Ma lihid rukhsad maamulka codsiyada.");
  }

  const id = clean_(request.id, 50);
  const status = clean_(request.status, 30);
  const response = clean_(request.response, 2000);
  const assigned = clean_(request.assignedTo, 120);
  const allowed = ["New", "In Progress", "Pending", "Resolved", "Rejected"];

  if (!allowed.includes(status)) throw new Error("Status aan sax ahayn.");

  const sh = sheet_(SHEETS.TICKETS);
  const data = sh.getDataRange().getValues();
  const index = data.findIndex((r, i) => i > 0 && r[0] === id);
  if (index < 1) throw new Error("Codsiga lama helin.");

  // Column J: Status, K: AssignedTo, L: ITResponse, M: UpdatedAt
  sh.getRange(index + 1, 10, 1, 4).setValues([[
    status, assigned, response, new Date()
  ]]);

  log_(user.email, "UPDATE_REQUEST", id);
  return { success: true };
}

/* SECTION 10: EXPORT WORD / PDF */
function exportReport(token, format) {
  const user = auth_(token);
  if (!["ADMIN", "IT_STAFF"].includes(user.role)) {
    throw new Error("Warbixintan waxaa dhoofin kara IT Staff/Admin.");
  }

  const tickets = rows_(SHEETS.TICKETS);
  const safe = s => String(s == null ? "" : s).replace(
    /[&<>"']/g, c => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;",
      '"': "&quot;", "'": "&#39;"
    }[c])
  );

  const body = tickets.map(t =>
    "<tr>" +
    [t.TicketID, t.CreatedAt, t.Department, t.SenderName,
     t.Title, t.Category, t.Priority, t.Status, t.ITResponse]
      .map(v => "<td>" + safe(v) + "</td>").join("") +
    "</tr>"
  ).join("");

  const html =
    "<html><head><meta charset='UTF-8'><style>" +
    "body{font-family:Arial,sans-serif;color:#222}" +
    "h1{color:#166534}table{border-collapse:collapse;width:100%;font-size:10px}" +
    "th,td{border:1px solid #ccc;padding:6px;text-align:left}" +
    "th{background:#166534;color:white}" +
    "</style></head><body>" +
    "<h1>MoES IT Help Desk — Requests Report</h1>" +
    "<p>Generated: " + safe(new Date().toLocaleString()) + "</p>" +
    "<table><thead><tr>" +
    ["ID","Date","Department","Sender","Title","Category","Priority","Status","IT Response"]
      .map(v => "<th>" + v + "</th>").join("") +
    "</tr></thead><tbody>" + body + "</tbody></table></body></html>";

  const stamp = Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm"
  );

  let blob, filename;
  if (String(format).toLowerCase() === "pdf") {
    blob = Utilities.newBlob(html, "text/html", "report.html")
      .getAs("application/pdf");
    filename = "MoES_IT_Report_" + stamp + ".pdf";
  } else if (String(format).toLowerCase() === "word") {
    // HTML document opened by Microsoft Word.
    blob = Utilities.newBlob(html, "application/msword",
      "MoES_IT_Report_" + stamp + ".doc");
    filename = "MoES_ICT_Report_" + stamp + ".doc";
  } else {
    throw new Error("Dooro PDF ama Word.");
  }

  blob.setName(filename);
  const file = DriveApp.createFile(blob);
  file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);
  log_(user.email, "EXPORT_" + format, filename);

  return {
    success: true,
    name: filename,
    url: file.getDownloadUrl()
  };
}
