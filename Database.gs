/**
 * Database.gs
 * Multi-Store Hierarchical Data Access Layer & Dynamic Schema Management
 * Google Sheets Enterprise Attendance Platform
 */

const Database = {
  DEFAULT_DB_NAME: "Mother_Hub_Attendance_Database",

  TAB_STORES: "Stores",
  TAB_HUBS_LEGACY: "Hubs",
  TAB_USERS: "Users",
  TAB_LOGS: "Attendance_Logs",

  STORES_HEADERS: [
    "Store_ID", "Store_Name", "Latitude", "Longitude", "Radius_Meters",
    "Primary_Admin_Email", "Status", "Created_By", "Created_At"
  ],

  USERS_HEADERS: [
    "User_ID", "Email", "Full_Name", "Employee_Code", "Role",
    "Assigned_Store_ID", "Reporting_Manager_Email", "Status",
    "Created_At", "Last_Login_At"
  ],

  LOGS_HEADERS: [
    "Log_ID", "Timestamp", "Email", "Full_Name", "Employee_Code",
    "Store_ID", "Store_Name", "Reporting_Manager_Email",
    "Check_Type", "User_Lat", "User_Lng", "Calculated_Distance_Meters",
    "GPS_Accuracy_Meters", "Device_Type", "Selfie_Drive_URL", "Verification_Status"
  ],

  DEFAULT_SUPER_ADMIN_EMAIL: "superadmin@example.com",

  /**
   * Retrieves or auto-provisions the active Google Spreadsheet
   */
  getSpreadsheet: function() {
    const props = PropertiesService.getScriptProperties();
    let sheetId = props.getProperty('SPREADSHEET_ID');

    // 1. Try bound spreadsheet
    try {
      const active = SpreadsheetApp.getActiveSpreadsheet();
      if (active) {
        if (!sheetId) props.setProperty('SPREADSHEET_ID', active.getId());
        return active;
      }
    } catch (e) {}

    // 2. Try configured SPREADSHEET_ID
    if (sheetId) {
      try {
        const ss = SpreadsheetApp.openById(sheetId);
        try {
          DriveApp.getFileById(sheetId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
        } catch (shareErr1) {}
        return ss;
      } catch (e) {
        Logger.log("Configured spreadsheet inaccessible, auto-locating: " + e.message);
      }
    }

    // 3. Search for existing or create new
    try {
      const files = DriveApp.getFilesByName(this.DEFAULT_DB_NAME);
      while (files.hasNext()) {
        const file = files.next();
        if (!file.isTrashed()) {
          try {
            const ss = SpreadsheetApp.openById(file.getId());
            props.setProperty('SPREADSHEET_ID', file.getId());
            try {
              file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
            } catch (shareErr2) {}
            return ss;
          } catch (openErr) {
            Logger.log("File found but open failed: " + openErr.message);
          }
        }
      }
    } catch (driveErr) {
      Logger.log("Drive search error: " + driveErr.message);
    }

    try {
      const newSs = SpreadsheetApp.create(this.DEFAULT_DB_NAME);
      props.setProperty('SPREADSHEET_ID', newSs.getId());
      try {
        const newFile = DriveApp.getFileById(newSs.getId());
        newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
      } catch (shareErr3) {}
      this.ensureInitialized(newSs.getId());
      return newSs;
    } catch (createErr) {
      Logger.log("Create spreadsheet error: " + createErr.message);
      throw new Error("Unable to open database. Please ensure deployment is set to 'Execute as: Me'. Details: " + createErr.message);
    }
  },

  /**
   * Self-healing tab retriever: ensures requested tab exists and is formatted
   */
  getTab: function(tabName) {
    const ss = this.getSpreadsheet();
    let sheet = ss.getSheetByName(tabName);

    // Backward-compatibility: if looking for 'Stores' and only legacy 'Hubs' exists, rename it
    if (!sheet && tabName === this.TAB_STORES) {
      const legacySheet = ss.getSheetByName(this.TAB_HUBS_LEGACY);
      if (legacySheet) {
        legacySheet.setName(this.TAB_STORES);
        sheet = legacySheet;
      }
    }

    if (!sheet) {
      this.ensureInitialized(ss.getId());
      sheet = ss.getSheetByName(tabName);
    }
    return sheet;
  },

  /**
   * Dynamic, Self-Healing Database Seeding & Schema Sync
   * 100% idempotent and non-destructive.
   */
  ensureInitialized: function(optionalSpreadsheetId) {
    let ss = null;
    if (optionalSpreadsheetId) {
      try {
        ss = SpreadsheetApp.openById(optionalSpreadsheetId);
      } catch (e) {
        ss = this.getSpreadsheet();
      }
    } else {
      ss = this.getSpreadsheet();
    }

    if (!ss) throw new Error("Could not access or initialize database spreadsheet.");

    try {
      const ssFile = DriveApp.getFileById(ss.getId());
      ssFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    } catch (shareErr) {}

    // 1. Stores Tab (Migrate from Hubs if present)
    let storesSheet = ss.getSheetByName(this.TAB_STORES);
    if (!storesSheet) {
      const legacyHubs = ss.getSheetByName(this.TAB_HUBS_LEGACY);
      if (legacyHubs) {
        legacyHubs.setName(this.TAB_STORES);
        storesSheet = legacyHubs;
      } else {
        storesSheet = ss.insertSheet(this.TAB_STORES);
      }
    }
    this._syncHeadersAndFormat(storesSheet, this.STORES_HEADERS);

    // Seed 1 default starter store ONLY when Stores table is completely empty
    if (storesSheet.getLastRow() <= 1) {
      const now = new Date().toISOString();
      const defaultStore = [
        "STR-001",
        "Headquarters",
        12.9352,
        77.6245,
        100,
        this.DEFAULT_SUPER_ADMIN_EMAIL,
        "ACTIVE",
        "SYSTEM",
        now
      ];
      storesSheet.appendRow(defaultStore);
    } else {
      // Auto-backfill any empty or boolean TRUE status cells from legacy sheets
      const storeData = storesSheet.getDataRange().getValues();
      for (let sIdx = 1; sIdx < storeData.length; sIdx++) {
        const rawStatus = storeData[sIdx][6];
        const statusStr = String(rawStatus !== undefined && rawStatus !== null ? rawStatus : "").trim().toUpperCase();
        if (!statusStr || statusStr === "TRUE" || rawStatus === true) {
          storesSheet.getRange(sIdx + 1, 7).setValue("ACTIVE");
        }
      }
    }

    // 2. Users Tab
    let usersSheet = ss.getSheetByName(this.TAB_USERS);
    if (!usersSheet) {
      usersSheet = ss.insertSheet(this.TAB_USERS);
    }
    this._syncHeadersAndFormat(usersSheet, this.USERS_HEADERS);
    this._seedSuperAdmin(usersSheet);

    // 3. Attendance Logs Tab
    let logsSheet = ss.getSheetByName(this.TAB_LOGS);
    if (!logsSheet) {
      logsSheet = ss.insertSheet(this.TAB_LOGS);
    }
    this._syncHeadersAndFormat(logsSheet, this.LOGS_HEADERS);

    // Clean up empty default "Sheet1" if present
    const defaultSheet = ss.getSheetByName("Sheet1");
    if (defaultSheet && ss.getSheets().length > 1) {
      try { ss.deleteSheet(defaultSheet); } catch (e) {}
    }

    // Auto-resize columns
    [storesSheet, usersSheet, logsSheet].forEach(sheet => {
      try {
        const lastCol = sheet.getLastColumn();
        if (lastCol > 0) sheet.autoResizeColumns(1, lastCol);
      } catch (e) {}
    });

    return {
      success: true,
      spreadsheetUrl: ss.getUrl(),
      sheets: [this.TAB_STORES, this.TAB_USERS, this.TAB_LOGS]
    };
  },

  /**
   * Alias for backward compatibility
   */
  seedSheetHeaders: function(optionalSpreadsheetId) {
    return this.ensureInitialized(optionalSpreadsheetId);
  },

  init: function() {
    return this.ensureInitialized();
  },

  /**
   * Non-destructive header synchronization:
   * Appends any newly introduced schema columns to Row 1 without altering existing data.
   */
  _syncHeadersAndFormat: function(sheet, schemaHeaders) {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow === 0 || lastCol === 0) {
      sheet.appendRow(schemaHeaders);
    } else {
      const existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
      const missingHeaders = [];
      schemaHeaders.forEach(sh => {
        if (!existingHeaders.includes(sh)) {
          missingHeaders.push(sh);
        }
      });

      if (missingHeaders.length > 0) {
        sheet.getRange(1, lastCol + 1, 1, missingHeaders.length).setValues([missingHeaders]);
      }
    }

    const updatedLastCol = sheet.getLastColumn();
    if (updatedLastCol > 0) {
      const headerRange = sheet.getRange(1, 1, 1, updatedLastCol);
      headerRange.setFontWeight("bold")
        .setFontColor("#FFFFFF")
        .setBackground("#18181B")
        .setFontFamily("Plus Jakarta Sans")
        .setFontSize(10)
        .setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }
  },

  /**
   * Seeds exactly ONE Super Admin with an example email ID when sheet is seeded
   */
  _seedSuperAdmin: function(usersSheet) {
    const data = usersSheet.getDataRange().getValues();
    let hasSuperAdmin = false;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][4]).trim().toUpperCase() === "SUPER_ADMIN") {
        hasSuperAdmin = true;
        break;
      }
    }

    // Only create 1 default Super Admin with example email if no Super Admin exists
    if (!hasSuperAdmin) {
      const now = new Date().toISOString();
      usersSheet.appendRow([
        "USR-001",
        this.DEFAULT_SUPER_ADMIN_EMAIL,
        "Super Admin",
        "EMP-001",
        "SUPER_ADMIN",
        "STR-001",
        "",
        "ACTIVE",
        now,
        now
      ]);
    }
  },

  /**
   * Searches for a user record by email
   */
  findUserByEmail: function(email) {
    if (!email) return null;
    const sheet = this.getTab(this.TAB_USERS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return null;

    const emailNorm = email.trim().toLowerCase();
    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][1]).trim().toLowerCase();
      if (rowEmail === emailNorm) {
        return {
          rowIndex: i + 1,
          userId: data[i][0],
          email: data[i][1],
          fullName: data[i][2],
          employeeCode: data[i][3],
          role: data[i][4] || "USER",
          assignedStoreId: data[i][5] || "",
          reportingManagerEmail: data[i][6] || "",
          status: data[i][7] || "ACTIVE",
          createdAt: data[i][8],
          lastLoginAt: data[i][9]
        };
      }
    }
    return null;
  },

  /**
   * Retrieves all users (for Super Admin Directory)
   */
  getAllUsers: function() {
    const sheet = this.getTab(this.TAB_USERS);
    const data = sheet.getDataRange().getValues();
    const users = [];

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      users.push({
        userId: data[i][0],
        email: data[i][1],
        fullName: data[i][2],
        employeeCode: data[i][3],
        role: data[i][4] || "USER",
        assignedStoreId: data[i][5] || "",
        reportingManagerEmail: data[i][6] || "",
        status: data[i][7] || "ACTIVE",
        createdAt: data[i][8],
        lastLoginAt: data[i][9]
      });
    }
    return users;
  },

  /**
   * Inserts or pre-provisions a user
   */
  insertUser: function(userData) {
    const sheet = this.getTab(this.TAB_USERS);
    const userId = "USR-" + Utilities.getUuid().slice(0, 8).toUpperCase();
    const now = new Date().toISOString();

    const row = [
      userId,
      userData.email.trim().toLowerCase(),
      userData.fullName ? userData.fullName.trim() : "",
      userData.employeeCode ? userData.employeeCode.trim().toUpperCase() : "",
      userData.role || "USER",
      userData.assignedStoreId || "",
      userData.reportingManagerEmail ? userData.reportingManagerEmail.trim().toLowerCase() : "",
      userData.status || "ACTIVE",
      now,
      now
    ];

    sheet.appendRow(row);
    return this.findUserByEmail(userData.email);
  },

  /**
   * Updates an existing user's profile, role, store, or status
   */
  updateUser: function(email, updateFields) {
    const user = this.findUserByEmail(email);
    if (!user) return null;

    const sheet = this.getTab(this.TAB_USERS);
    if (updateFields.fullName !== undefined) sheet.getRange(user.rowIndex, 3).setValue(updateFields.fullName);
    if (updateFields.employeeCode !== undefined) sheet.getRange(user.rowIndex, 4).setValue(updateFields.employeeCode);
    if (updateFields.role !== undefined) sheet.getRange(user.rowIndex, 5).setValue(updateFields.role);
    if (updateFields.assignedStoreId !== undefined) sheet.getRange(user.rowIndex, 6).setValue(updateFields.assignedStoreId);
    if (updateFields.reportingManagerEmail !== undefined) sheet.getRange(user.rowIndex, 7).setValue(updateFields.reportingManagerEmail);
    if (updateFields.status !== undefined) sheet.getRange(user.rowIndex, 8).setValue(updateFields.status);
    if (updateFields.lastLoginAt !== undefined) sheet.getRange(user.rowIndex, 10).setValue(updateFields.lastLoginAt);

    return this.findUserByEmail(email);
  },

  /**
   * Updates user's last login timestamp
   */
  updateLastLogin: function(email) {
    const user = this.findUserByEmail(email);
    if (!user) return;
    const sheet = this.getTab(this.TAB_USERS);
    sheet.getRange(user.rowIndex, 10).setValue(new Date().toISOString());
  },

  /**
   * Retrieves all stores (with optional inactive filter)
   */
  getAllStores: function(includeInactive) {
    const sheet = this.getTab(this.TAB_STORES);
    const data = sheet.getDataRange().getValues();
    const stores = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      const statusRaw = row[6];
      const statusStr = String(statusRaw !== undefined && statusRaw !== null ? statusRaw : "").trim().toUpperCase();
      const isInactive = (statusStr === "INACTIVE" || statusStr === "FALSE" || statusRaw === false);
      
      if (includeInactive || !isInactive) {
        stores.push({
          storeId: String(row[0]).trim(),
          storeName: String(row[1]).trim(),
          latitude: Number(row[2]),
          longitude: Number(row[3]),
          radiusMeters: Number(row[4]) || 100,
          primaryAdminEmail: String(row[5] || "").trim(),
          status: isInactive ? "INACTIVE" : "ACTIVE",
          createdBy: String(row[7] || ""),
          createdAt: row[8] || ""
        });
      }
    }
    return stores;
  },

  /**
   * Retrieves all active stores
   */
  getActiveStores: function() {
    return this.getAllStores(false);
  },

  /**
   * Retrieves a single Store by ID OR Store Name (flexible matching)
   */
  getStoreById: function(storeIdOrName) {
    if (!storeIdOrName) return null;
    const sheet = this.getTab(this.TAB_STORES);
    const data = sheet.getDataRange().getValues();
    const query = String(storeIdOrName).trim().toUpperCase();

    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][0] || "").trim().toUpperCase();
      const rowName = String(data[i][1] || "").trim().toUpperCase();
      if (rowId === query || rowName === query) {
        const rawStatus = data[i][6];
        const statusStr = String(rawStatus !== undefined && rawStatus !== null ? rawStatus : "").trim().toUpperCase();
        const isInactive = (statusStr === "INACTIVE" || statusStr === "FALSE" || rawStatus === false);
        return {
          rowIndex: i + 1,
          storeId: String(data[i][0]).trim(),
          storeName: String(data[i][1]).trim(),
          latitude: Number(data[i][2]),
          longitude: Number(data[i][3]),
          radiusMeters: Number(data[i][4]) || 100,
          primaryAdminEmail: String(data[i][5] || "").trim(),
          status: isInactive ? "INACTIVE" : "ACTIVE",
          createdBy: String(data[i][7] || ""),
          createdAt: data[i][8] || ""
        };
      }
    }
    return null;
  },

  /**
   * Generates next sequential Store ID: STR-001, STR-002, etc.
   */
  getNextStoreId: function() {
    const sheet = this.getTab(this.TAB_STORES);
    const data = sheet.getDataRange().getValues();
    let maxNum = 0;
    for (let i = 1; i < data.length; i++) {
      const match = String(data[i][0]).match(/^STR-(\d+)$/i);
      if (match) {
        const num = parseInt(match[1], 10);
        if (num > maxNum) maxNum = num;
      }
    }
    const nextNum = maxNum + 1;
    return "STR-" + String(nextNum).padStart(3, "0");
  },

  /**
   * Upserts a store (Create if new, Update if Store_ID matches)
   */
  upsertStore: function(storeData, creatorEmail) {
    const sheet = this.getTab(this.TAB_STORES);
    const data = sheet.getDataRange().getValues();
    let targetRow = -1;

    if (storeData.storeId) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).toUpperCase() === String(storeData.storeId).toUpperCase()) {
          targetRow = i + 1;
          break;
        }
      }
    }

    const storeId = (targetRow > 0 && storeData.storeId) ? storeData.storeId : (storeData.storeId || this.getNextStoreId());
    const now = new Date().toISOString();

    const rowValues = [
      storeId,
      storeData.storeName,
      Number(storeData.latitude),
      Number(storeData.longitude),
      Number(storeData.radiusMeters) || 100,
      storeData.primaryAdminEmail ? storeData.primaryAdminEmail.trim().toLowerCase() : "",
      storeData.status || "ACTIVE",
      creatorEmail || "SUPER_ADMIN",
      now
    ];

    if (targetRow > 0) {
      sheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }
    return storeId;
  },

  /**
   * Deactivates a store
   */
  deactivateStore: function(storeId) {
    const store = this.getStoreById(storeId);
    if (!store) return false;
    const sheet = this.getTab(this.TAB_STORES);
    sheet.getRange(store.rowIndex, 7).setValue("INACTIVE");
    return true;
  },

  /**
   * Appends an attendance record to Attendance_Logs
   */
  appendAttendanceLog: function(log) {
    const sheet = this.getTab(this.TAB_LOGS);
    const logId = "LOG-" + Utilities.getUuid().slice(0, 10).toUpperCase();
    const timestamp = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd'T'HH:mm:ssXXX");

    const row = [
      logId,
      timestamp,
      log.email,
      log.fullName,
      log.employeeCode,
      log.storeId,
      log.storeName,
      log.reportingManagerEmail || "",
      log.checkType,
      log.userLat,
      log.userLng,
      log.distanceMeters,
      log.accuracyMeters,
      log.deviceType,
      log.selfieDriveUrl,
      log.verificationStatus
    ];

    sheet.appendRow(row);
    return logId;
  },

  /**
   * Fetches the last N logs for a single user
   */
  getRecentLogsForUser: function(email, limit) {
    const sheet = this.getTab(this.TAB_LOGS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const emailNorm = String(email).trim().toLowerCase();
    const results = [];

    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][2]).trim().toLowerCase() === emailNorm) {
        results.push(this._formatLogRow(data[i]));
        if (results.length >= (limit || 5)) break;
      }
    }
    return results;
  },

  /**
   * Hierarchical Scoped Query: Retrieves attendance logs based on caller role
   */
  getScopedAttendanceLogs: function(callerRole, callerStoreId, callerEmail, filterStoreId, limit) {
    const sheet = this.getTab(this.TAB_LOGS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const maxLimit = limit || 100;
    const results = [];
    const callerEmailNorm = String(callerEmail).trim().toLowerCase();

    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const logStoreId = String(row[5]);
      const logManagerEmail = String(row[7]).trim().toLowerCase();

      let isAllowed = false;

      if (callerRole === "SUPER_ADMIN") {
        if (!filterStoreId || filterStoreId === "ALL" || logStoreId === filterStoreId) {
          isAllowed = true;
        }
      } else if (callerRole === "STORE_ADMIN") {
        if (logManagerEmail === callerEmailNorm || logStoreId === callerStoreId) {
          isAllowed = true;
        }
      }

      if (isAllowed) {
        results.push(this._formatLogRow(row));
        if (results.length >= maxLimit) break;
      }
    }
    return results;
  },

  /**
   * Hierarchical Scoped Query: Retrieves users roster based on caller role
   */
  getScopedUsers: function(callerRole, callerStoreId, callerEmail) {
    const sheet = this.getTab(this.TAB_USERS);
    const data = sheet.getDataRange().getValues();
    const users = [];
    const callerEmailNorm = String(callerEmail).trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;

      const userRole = String(data[i][4]);
      const userStoreId = String(data[i][5]);
      const userManager = String(data[i][6]).trim().toLowerCase();

      let isAllowed = false;

      if (callerRole === "SUPER_ADMIN") {
        isAllowed = true;
      } else if (callerRole === "STORE_ADMIN") {
        if (userManager === callerEmailNorm || userStoreId === callerStoreId) {
          isAllowed = true;
        }
      }

      if (isAllowed) {
        users.push({
          userId: data[i][0],
          email: data[i][1],
          fullName: data[i][2],
          employeeCode: data[i][3],
          role: userRole,
          assignedStoreId: userStoreId,
          reportingManagerEmail: data[i][6],
          status: data[i][7],
          createdAt: data[i][8],
          lastLoginAt: data[i][9]
        });
      }
    }
    return users;
  },

  /**
   * Super Admin Global KPIs Calculation
   */
  getGlobalKpis: function() {
    const activeStores = this.getActiveStores();
    const allUsers = this.getAllUsers();
    
    let totalStoreAdmins = 0;
    let totalStaff = 0;

    allUsers.forEach(u => {
      if (u.status === "ACTIVE") {
        if (u.role === "STORE_ADMIN") totalStoreAdmins++;
        else if (u.role === "USER") totalStaff++;
      }
    });

    // Count today's check-ins
    const logsSheet = this.getTab(this.TAB_LOGS);
    const data = logsSheet.getDataRange().getValues();
    const todayStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd");
    let totalCheckInsToday = 0;

    for (let i = 1; i < data.length; i++) {
      const ts = String(data[i][1]);
      const checkType = String(data[i][8]);
      if (ts.startsWith(todayStr) && checkType === "CHECK_IN") {
        totalCheckInsToday++;
      }
    }

    return {
      totalActiveStores: activeStores.length,
      totalStoreAdmins: totalStoreAdmins,
      totalStaff: totalStaff,
      totalCheckInsToday: totalCheckInsToday
    };
  },

  /**
   * Store Admin KPIs Calculation (Scoped to single store)
   */
  getStoreKpis: function(storeId) {
    const logsSheet = this.getTab(this.TAB_LOGS);
    const data = logsSheet.getDataRange().getValues();
    const todayStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd");
    
    let checkedInToday = 0;
    let checkedOutToday = 0;

    for (let i = 1; i < data.length; i++) {
      const ts = String(data[i][1]);
      const rowStoreId = String(data[i][5]);
      const checkType = String(data[i][8]);

      if (rowStoreId === storeId && ts.startsWith(todayStr)) {
        if (checkType === "CHECK_IN") checkedInToday++;
        else if (checkType === "CHECK_OUT") checkedOutToday++;
      }
    }

    const allUsers = this.getAllUsers();
    const totalStoreStaff = allUsers.filter(u => u.assignedStoreId === storeId && u.status === "ACTIVE" && u.role === "USER").length;

    return {
      checkedInToday: checkedInToday,
      checkedOutToday: checkedOutToday,
      totalStoreStaff: totalStoreStaff
    };
  },

  _formatLogRow: function(r) {
    return {
      logId: r[0],
      timestamp: r[1],
      email: r[2],
      fullName: r[3],
      employeeCode: r[4],
      storeId: r[5],
      storeName: r[6],
      reportingManagerEmail: r[7],
      checkType: r[8],
      userLat: r[9],
      userLng: r[10],
      distanceMeters: r[11],
      accuracyMeters: r[12],
      deviceType: r[13],
      selfieDriveUrl: r[14],
      verificationStatus: r[15]
    };
  }
};
