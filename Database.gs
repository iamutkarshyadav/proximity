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
    "Primary_Admin_Email", "Status", "Created_By", "Created_At", "Manager_Emails",
    "Regional_Manager_Emails", "Region_Name"
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
    "GPS_Accuracy_Meters", "Device_Type", "Selfie_Drive_URL", "Verification_Status",
    "Session_ID", "Work_Duration_Minutes"
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
   * Retrieves all stores (with optional inactive filter and populated manager roster)
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
        const storeId = String(row[0]).trim();
        const primaryEmail = String(row[5] || "").trim().toLowerCase();
        const managerEmailsRaw = String(row[9] || "").trim();
        const regionalManagerEmailsRaw = String(row[10] || "").trim();
        const regionalManagerEmails = regionalManagerEmailsRaw
          ? regionalManagerEmailsRaw.split(',').map(e => e.trim().toLowerCase()).filter(Boolean)
          : [];
        const regionName = String(row[11] || "").trim();
        const managers = this.getStoreManagers(storeId, primaryEmail);

        stores.push({
          storeId: storeId,
          storeName: String(row[1]).trim(),
          latitude: Number(row[2]),
          longitude: Number(row[3]),
          radiusMeters: Number(row[4]) || 100,
          primaryAdminEmail: primaryEmail,
          managerEmails: managerEmailsRaw,
          regionalManagerEmails: regionalManagerEmails,
          regionalManagerEmailsRaw: regionalManagerEmailsRaw,
          regionName: regionName,
          managers: managers,
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
        const storeId = String(data[i][0]).trim();
        const primaryEmail = String(data[i][5] || "").trim().toLowerCase();
        const managerEmailsRaw = String(data[i][9] || "").trim();
        const regionalManagerEmailsRaw = String(data[i][10] || "").trim();
        const regionalManagerEmails = regionalManagerEmailsRaw
          ? regionalManagerEmailsRaw.split(',').map(e => e.trim().toLowerCase()).filter(Boolean)
          : [];
        const regionName = String(data[i][11] || "").trim();
        const managers = this.getStoreManagers(storeId, primaryEmail);

        return {
          rowIndex: i + 1,
          storeId: storeId,
          storeName: String(data[i][1]).trim(),
          latitude: Number(data[i][2]),
          longitude: Number(data[i][3]),
          radiusMeters: Number(data[i][4]) || 100,
          primaryAdminEmail: primaryEmail,
          managerEmails: managerEmailsRaw,
          regionalManagerEmails: regionalManagerEmails,
          regionalManagerEmailsRaw: regionalManagerEmailsRaw,
          regionName: regionName,
          managers: managers,
          status: isInactive ? "INACTIVE" : "ACTIVE",
          createdBy: String(data[i][7] || ""),
          createdAt: data[i][8] || ""
        };
      }
    }
    return null;
  },

  /**
   * Retrieves all managers assigned to a store (supports multiple managers)
   */
  getStoreManagers: function(storeId, optionalPrimaryEmail) {
    if (!storeId) return [];
    const allUsers = this.getAllUsers();
    const cleanId = String(storeId).trim().toUpperCase();
    const primaryEmail = optionalPrimaryEmail ? String(optionalPrimaryEmail).trim().toLowerCase() : "";

    const managers = [];
    const seenEmails = {};

    allUsers.forEach(u => {
      const uEmail = String(u.email || "").trim().toLowerCase();
      const uStore = String(u.assignedStoreId || "").trim().toUpperCase();
      const uRole = String(u.role || "").trim().toUpperCase();
      const isStoreAdmin = (uRole === 'STORE_ADMIN' || uRole === 'SUPER_ADMIN');

      if (isStoreAdmin && uStore === cleanId && u.status === 'ACTIVE') {
        seenEmails[uEmail] = true;
        managers.push({
          userId: u.userId,
          email: uEmail,
          fullName: u.fullName || uEmail,
          employeeCode: u.employeeCode || "MGR",
          role: u.role,
          isPrimary: (uEmail === primaryEmail),
          status: u.status
        });
      }
    });

    // If primary admin email is set on store but not in Users as STORE_ADMIN, ensure included
    if (primaryEmail && !seenEmails[primaryEmail]) {
      const pUser = this.findUserByEmail(primaryEmail);
      managers.unshift({
        userId: pUser ? pUser.userId : "USR-PRIMARY",
        email: primaryEmail,
        fullName: pUser ? pUser.fullName : "Primary Manager",
        employeeCode: pUser ? pUser.employeeCode : "MGR-001",
        role: "STORE_ADMIN",
        isPrimary: true,
        status: pUser ? pUser.status : "ACTIVE"
      });
    }

    return managers;
  },

  /**
   * Adds or promotes a manager to an existing store (Multi-Manager)
   */
  addStoreManager: function(storeId, email, fullName, employeeCode, isPrimary, callerEmail) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
    } catch (e) {}

    try {
      const cleanEmail = String(email).trim().toLowerCase();
      const store = this.getStoreById(storeId);
      if (!store) throw new Error("Store not found: " + storeId);

      let user = this.findUserByEmail(cleanEmail);
      if (user) {
        this.updateUser(cleanEmail, {
          role: "STORE_ADMIN",
          assignedStoreId: store.storeId,
          fullName: fullName ? fullName.trim() : user.fullName,
          status: "ACTIVE"
        });
      } else {
        const empCode = employeeCode ? employeeCode.trim().toUpperCase() : ("MGR-" + Math.floor(1000 + Math.random() * 9000));
        this.insertUser({
          email: cleanEmail,
          fullName: fullName ? fullName.trim() : "Store Manager",
          employeeCode: empCode,
          role: "STORE_ADMIN",
          assignedStoreId: store.storeId,
          reportingManagerEmail: callerEmail || "",
          status: "ACTIVE"
        });
      }

      // If isPrimary or store has no primary admin, set primary
      if (isPrimary || !store.primaryAdminEmail) {
        this.setPrimaryStoreManager(store.storeId, cleanEmail);
      }

      this.syncStoreManagerEmails(store.storeId);
      return this.getStoreById(store.storeId);
    } finally {
      lock.releaseLock();
    }
  },

  /**
   * Removes a manager from a store
   */
  removeStoreManager: function(storeId, email, callerEmail) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
    } catch (e) {}

    try {
      const cleanEmail = String(email).trim().toLowerCase();
      const store = this.getStoreById(storeId);
      if (!store) throw new Error("Store not found.");

      const user = this.findUserByEmail(cleanEmail);
      if (user && String(user.assignedStoreId).trim().toUpperCase() === String(store.storeId).trim().toUpperCase()) {
        this.updateUser(cleanEmail, {
          assignedStoreId: "",
          status: "INACTIVE"
        });
      }

      // If removed user was primary, reassign primary to another manager if available
      const remainingManagers = this.getStoreManagers(store.storeId);
      const remainingActive = remainingManagers.filter(m => m.email !== cleanEmail);
      const nextPrimary = remainingActive.length > 0 ? remainingActive[0].email : "";

      if (store.primaryAdminEmail === cleanEmail) {
        this.setPrimaryStoreManager(store.storeId, nextPrimary);
      }

      this.syncStoreManagerEmails(store.storeId);
      return this.getStoreById(store.storeId);
    } finally {
      lock.releaseLock();
    }
  },

  /**
   * Sets the designated primary manager for a store
   */
  setPrimaryStoreManager: function(storeId, email) {
    const store = this.getStoreById(storeId);
    if (!store) return null;
    const cleanEmail = String(email || "").trim().toLowerCase();
    const sheet = this.getTab(this.TAB_STORES);
    sheet.getRange(store.rowIndex, 6).setValue(cleanEmail);
    this.syncStoreManagerEmails(storeId);
    return this.getStoreById(storeId);
  },

  /**
   * Synchronizes the Manager_Emails column in the Stores sheet
   */
  syncStoreManagerEmails: function(storeId) {
    const store = this.getStoreById(storeId);
    if (!store) return;
    const sheet = this.getTab(this.TAB_STORES);
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
    let colIdx = headers.indexOf("Manager_Emails") + 1;

    // If Manager_Emails column missing, append it
    if (colIdx <= 0) {
      sheet.getRange(1, lastCol + 1).setValue("Manager_Emails");
      colIdx = lastCol + 1;
    }

    const managers = this.getStoreManagers(store.storeId);
    const emailList = managers.map(m => m.email).join(", ");
    sheet.getRange(store.rowIndex, colIdx).setValue(emailList);
  },

  /**
   * Retrieves all stores assigned to a Regional Manager (Many-to-Many)
   */
  getStoresForRegionalManager: function(rmEmail) {
    if (!rmEmail) return [];
    const cleanEmail = String(rmEmail).trim().toLowerCase();
    const allStores = this.getActiveStores();
    return allStores.filter(s => {
      if (s.regionalManagerEmails && s.regionalManagerEmails.includes(cleanEmail)) return true;
      if (s.primaryAdminEmail && s.primaryAdminEmail === cleanEmail) return true;
      return false;
    });
  },

  /**
   * Retrieves all Regional Managers assigned to a specific store
   */
  getStoreRegionalManagers: function(storeId) {
    const store = this.getStoreById(storeId);
    if (!store) return [];
    const emails = store.regionalManagerEmails || [];
    const allUsers = this.getAllUsers();
    const userMap = {};
    allUsers.forEach(u => {
      userMap[String(u.email).trim().toLowerCase()] = u;
    });

    return emails.map(email => {
      const u = userMap[email];
      return {
        email: email,
        fullName: u ? u.fullName : email,
        employeeCode: u ? u.employeeCode : "RM",
        role: "REGIONAL_MANAGER",
        status: u ? u.status : "ACTIVE"
      };
    });
  },

  /**
   * Adds a Regional Manager to a Store (Many-to-Many)
   */
  addStoreRegionalManager: function(storeId, rmEmail, rmName, rmCode, callerEmail) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
    } catch (e) {}

    try {
      const cleanEmail = String(rmEmail).trim().toLowerCase();
      const store = this.getStoreById(storeId);
      if (!store) throw new Error("Store not found: " + storeId);

      // Ensure user exists as REGIONAL_MANAGER
      let user = this.findUserByEmail(cleanEmail);
      if (user) {
        if (user.role !== "SUPER_ADMIN") {
          this.updateUser(cleanEmail, {
            role: "REGIONAL_MANAGER",
            fullName: rmName ? rmName.trim() : user.fullName,
            status: "ACTIVE"
          });
        }
      } else {
        const empCode = rmCode ? rmCode.trim().toUpperCase() : ("RM-" + Math.floor(1000 + Math.random() * 9000));
        this.insertUser({
          email: cleanEmail,
          fullName: rmName ? rmName.trim() : "Regional Manager",
          employeeCode: empCode,
          role: "REGIONAL_MANAGER",
          assignedStoreId: store.storeId,
          reportingManagerEmail: callerEmail || "",
          status: "ACTIVE"
        });
      }

      // Update store's Regional_Manager_Emails
      const currentList = store.regionalManagerEmails ? [...store.regionalManagerEmails] : [];
      if (!currentList.includes(cleanEmail)) {
        currentList.push(cleanEmail);
      }

      const sheet = this.getTab(this.TAB_STORES);
      const lastCol = sheet.getLastColumn();
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
      let colIdx = headers.indexOf("Regional_Manager_Emails") + 1;
      if (colIdx <= 0) {
        sheet.getRange(1, lastCol + 1).setValue("Regional_Manager_Emails");
        colIdx = lastCol + 1;
      }

      sheet.getRange(store.rowIndex, colIdx).setValue(currentList.join(", "));
      return this.getStoreById(store.storeId);
    } finally {
      lock.releaseLock();
    }
  },

  /**
   * Removes a Regional Manager from a Store
   */
  removeStoreRegionalManager: function(storeId, rmEmail, callerEmail) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
    } catch (e) {}

    try {
      const cleanEmail = String(rmEmail).trim().toLowerCase();
      const store = this.getStoreById(storeId);
      if (!store) throw new Error("Store not found: " + storeId);

      const currentList = store.regionalManagerEmails ? [...store.regionalManagerEmails] : [];
      const updatedList = currentList.filter(e => e !== cleanEmail);

      const sheet = this.getTab(this.TAB_STORES);
      const lastCol = sheet.getLastColumn();
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
      let colIdx = headers.indexOf("Regional_Manager_Emails") + 1;
      if (colIdx <= 0) {
        sheet.getRange(1, lastCol + 1).setValue("Regional_Manager_Emails");
        colIdx = lastCol + 1;
      }

      sheet.getRange(store.rowIndex, colIdx).setValue(updatedList.join(", "));
      return this.getStoreById(store.storeId);
    } finally {
      lock.releaseLock();
    }
  },

  /**
   * Sets the full list of Regional Managers for a Store
   */
  setStoreRegionalManagers: function(storeId, rmEmailsList) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
    } catch (e) {}

    try {
      const store = this.getStoreById(storeId);
      if (!store) throw new Error("Store not found: " + storeId);

      const cleanList = Array.isArray(rmEmailsList) 
        ? rmEmailsList.map(e => String(e).trim().toLowerCase()).filter(Boolean)
        : String(rmEmailsList || "").split(',').map(e => String(e).trim().toLowerCase()).filter(Boolean);

      const sheet = this.getTab(this.TAB_STORES);
      const lastCol = sheet.getLastColumn();
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
      let colIdx = headers.indexOf("Regional_Manager_Emails") + 1;
      if (colIdx <= 0) {
        sheet.getRange(1, lastCol + 1).setValue("Regional_Manager_Emails");
        colIdx = lastCol + 1;
      }

      sheet.getRange(store.rowIndex, colIdx).setValue(cleanList.join(", "));
      return this.getStoreById(store.storeId);
    } finally {
      lock.releaseLock();
    }
  },

  /**
   * Batch upsert staff roster for a Mother Hub (Exclusively invoked by RM or Super Admin)
   */
  batchUpsertHubRoster: function(storeId, rosterRows, callerEmail) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(15000);
    } catch (e) {}

    try {
      const store = this.getStoreById(storeId);
      if (!store) throw new Error("Store not found: " + storeId);

      if (!Array.isArray(rosterRows) || rosterRows.length === 0) {
        throw new Error("No roster entries provided for upload.");
      }

      let upsertCount = 0;
      const hubManagerEmails = [];

      rosterRows.forEach((row, idx) => {
        if (!row.email) return;
        const cleanEmail = String(row.email).trim().toLowerCase();
        if (!cleanEmail || cleanEmail.indexOf('@') === -1) return;

        const fullName = String(row.fullName || row.name || cleanEmail.split('@')[0]).trim();
        const empCode = String(row.employeeCode || row.empCode || ("EMP-" + Math.floor(1000 + Math.random() * 9000))).trim().toUpperCase();
        const isManager = String(row.designation || row.role || "").toUpperCase().indexOf("MGR") > -1 || String(row.designation || "").toUpperCase().indexOf("MANAGER") > -1;

        if (isManager) {
          hubManagerEmails.push(cleanEmail);
        }

        const existing = this.findUserByEmail(cleanEmail);
        if (existing) {
          if (existing.role === 'SUPER_ADMIN') {
            // Top-level Super Admin cannot be downgraded or altered by Regional Manager roster uploads
            return;
          }
          this.updateUser(cleanEmail, {
            fullName: fullName,
            employeeCode: empCode,
            assignedStoreId: store.storeId,
            role: "USER", // All hub staff and hub managers are role USER!
            status: "ACTIVE"
          });
        } else {
          this.insertUser({
            email: cleanEmail,
            fullName: fullName,
            employeeCode: empCode,
            role: "USER", // All hub staff and hub managers are role USER!
            assignedStoreId: store.storeId,
            reportingManagerEmail: callerEmail || "",
            status: "ACTIVE"
          });
        }
        upsertCount++;
      });

      // Update store manager emails if any manager designations were provided
      if (hubManagerEmails.length > 0) {
        const sheet = this.getTab(this.TAB_STORES);
        const existingMgrs = store.managerEmails ? store.managerEmails.split(',').map(e => e.trim().toLowerCase()) : [];
        const mergedMgrs = Array.from(new Set([...existingMgrs, ...hubManagerEmails]));
        sheet.getRange(store.rowIndex, 10).setValue(mergedMgrs.join(", "));
      }

      return {
        success: true,
        count: upsertCount,
        storeId: store.storeId,
        storeName: store.storeName
      };
    } finally {
      lock.releaseLock();
    }
  },

  /**
   * Adds a single staff member or manager to a Mother Hub
   */
  addSingleStaffToHub: function(storeId, email, fullName, employeeCode, designation, callerEmail) {
    return this.batchUpsertHubRoster(storeId, [{
      email: email,
      fullName: fullName,
      employeeCode: employeeCode,
      designation: designation
    }], callerEmail);
  },

  /**
   * Removes or deactivates a staff member from a Mother Hub
   */
  removeStaffFromHub: function(storeId, email, callerEmail) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
    } catch (e) {}

    try {
      const cleanEmail = String(email).trim().toLowerCase();
      const user = this.findUserByEmail(cleanEmail);
      if (user) {
        if (user.role === 'SUPER_ADMIN') {
          throw new Error("Super Admin is the top-level authority and cannot be removed from any Mother Hub.");
        }
        this.updateUser(cleanEmail, {
          assignedStoreId: "",
          status: "INACTIVE"
        });
      }
      return { success: true, email: cleanEmail };
    } finally {
      lock.releaseLock();
    }
  },

  /**
   * Aggregates KPIs across multiple Mother Hubs for a Regional Manager
   */
  getRegionalAggregatedKpis: function(storeIds) {
    const validStoreIds = (Array.isArray(storeIds) ? storeIds : [storeIds]).map(id => String(id).trim().toUpperCase());
    const allUsers = this.getAllUsers();
    const todayStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd");

    let totalStaff = 0;
    const storeStaffEmails = new Set();
    allUsers.forEach(u => {
      const uStore = String(u.assignedStoreId || "").trim().toUpperCase();
      if (validStoreIds.includes(uStore) && u.status === 'ACTIVE') {
        totalStaff++;
        storeStaffEmails.add(String(u.email).trim().toLowerCase());
      }
    });

    const sheet = this.getTab(this.TAB_LOGS);
    const data = sheet.getDataRange().getValues();
    const checkedInEmails = new Set();
    const checkedOutEmails = new Set();
    const latestUserStatus = {};

    for (let i = 1; i < data.length; i++) {
      const logStoreId = String(data[i][5] || "").trim().toUpperCase();
      if (!validStoreIds.includes(logStoreId)) continue;

      const email = String(data[i][2] || "").trim().toLowerCase();
      const ts = String(data[i][1] || "");
      const checkType = String(data[i][8] || "").trim().toUpperCase();

      if (ts.startsWith(todayStr)) {
        if (checkType === 'CHECK_IN') checkedInEmails.add(email);
        if (checkType === 'CHECK_OUT') checkedOutEmails.add(email);
        latestUserStatus[email] = checkType;
      }
    }

    let onDutyNow = 0;
    Object.keys(latestUserStatus).forEach(e => {
      if (latestUserStatus[e] === 'CHECK_IN') onDutyNow++;
    });

    return {
      managedHubsCount: validStoreIds.length,
      totalStaff: totalStaff,
      checkedInToday: checkedInEmails.size,
      checkedOutToday: checkedOutEmails.size,
      onDutyNow: onDutyNow,
      absentToday: Math.max(0, totalStaff - checkedInEmails.size)
    };
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
   * Upserts a store (Create if new, Update if Store_ID matches) with lock protection
   */
  upsertStore: function(storeData, creatorEmail) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
    } catch (e) {}

    try {
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
        now,
        storeData.managerEmails || ""
      ];

      if (targetRow > 0) {
        sheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
      } else {
        sheet.appendRow(rowValues);
      }
      return storeId;
    } finally {
      lock.releaseLock();
    }
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
   * Appends an attendance record with atomic session pairing & shift duration calculation
   */
  appendAttendanceLog: function(log) {
    const sheet = this.getTab(this.TAB_LOGS);
    const logId = "LOG-" + Utilities.getUuid().slice(0, 10).toUpperCase();
    const timestamp = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd'T'HH:mm:ssXXX");

    let sessionId = log.sessionId || "";
    let workDurationMinutes = (log.workDurationMinutes !== undefined && log.workDurationMinutes !== null && log.workDurationMinutes !== "") ? Number(log.workDurationMinutes) : "";

    // Automatic Session Pairing for CHECK_OUT
    if (log.checkType === 'CHECK_OUT') {
      const data = sheet.getDataRange().getValues();
      const emailNorm = String(log.email).trim().toLowerCase();
      let openCheckInRow = -1;
      let openCheckInTs = null;
      let openCheckInSessionId = "";

      // Scan backwards to find the user's latest unclosed CHECK_IN
      for (let i = data.length - 1; i >= 1; i--) {
        const rowEmail = String(data[i][2]).trim().toLowerCase();
        const rowType = String(data[i][8]);
        if (rowEmail === emailNorm) {
          if (rowType === 'CHECK_IN') {
            openCheckInRow = i + 1;
            openCheckInTs = data[i][1];
            openCheckInSessionId = data[i][16] || "";
            break;
          } else if (rowType === 'CHECK_OUT') {
            // Already closed earlier shift
            break;
          }
        }
      }

      if (openCheckInRow > 0 && openCheckInTs) {
        sessionId = openCheckInSessionId || ("SES-" + Utilities.getUuid().slice(0, 8).toUpperCase());
        const inDate = new Date(openCheckInTs);
        const outDate = new Date();
        workDurationMinutes = Math.max(0, Math.round((outDate.getTime() - inDate.getTime()) / 60000));

        // If open check-in was missing sessionId, backfill it in column 17
        if (!openCheckInSessionId && sheet.getLastColumn() >= 17) {
          sheet.getRange(openCheckInRow, 17).setValue(sessionId);
        }
      } else {
        sessionId = "SES-" + Utilities.getUuid().slice(0, 8).toUpperCase();
        workDurationMinutes = 0;
      }
    } else if (log.checkType === 'CHECK_IN') {
      if (!sessionId) {
        sessionId = "SES-" + Utilities.getUuid().slice(0, 8).toUpperCase();
      }
    }

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
      log.verificationStatus,
      sessionId,
      workDurationMinutes
    ];

    sheet.appendRow(row);
    return logId;
  },

  /**
   * Retrieves ground staff current real-time shift status & today's summary tracker
   */
  getUserShiftStatus: function(email) {
    if (!email) {
      return { status: 'NOT_CHECKED_IN', activeShift: null, todayShifts: [], totalTodayMinutes: 0, totalTodayFormatted: "0m" };
    }

    const sheet = this.getTab(this.TAB_LOGS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { status: 'NOT_CHECKED_IN', activeShift: null, todayShifts: [], totalTodayMinutes: 0, totalTodayFormatted: "0m" };
    }

    const emailNorm = String(email).trim().toLowerCase();
    const todayStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd");

    const userLogs = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim().toLowerCase() === emailNorm) {
        userLogs.push(this._formatLogRow(data[i]));
      }
    }

    if (userLogs.length === 0) {
      return { status: 'NOT_CHECKED_IN', activeShift: null, todayShifts: [], totalTodayMinutes: 0, totalTodayFormatted: "0m" };
    }

    const latest = userLogs[userLogs.length - 1];
    let currentStatus = 'NOT_CHECKED_IN';
    let activeShift = null;

    if (latest.checkType === 'CHECK_IN') {
      currentStatus = 'ON_DUTY';
      const checkInDate = new Date(latest.timestamp);
      const now = new Date();
      const elapsedSeconds = Math.max(0, Math.floor((now.getTime() - checkInDate.getTime()) / 1000));
      const elapsedMinutes = Math.floor(elapsedSeconds / 60);

      activeShift = {
        logId: latest.logId,
        sessionId: latest.sessionId,
        storeId: latest.storeId,
        storeName: latest.storeName,
        checkInTimestamp: latest.timestamp,
        checkInTimeFormatted: this._formatTimeOnly(latest.timestamp),
        checkInPhotoUrl: latest.selfieDriveUrl,
        distanceMeters: latest.distanceMeters,
        accuracyMeters: latest.accuracyMeters,
        elapsedSeconds: elapsedSeconds,
        elapsedFormatted: this._formatDuration(elapsedMinutes)
      };
    } else if (latest.checkType === 'CHECK_OUT') {
      const latestTs = String(latest.timestamp);
      if (latestTs.startsWith(todayStr)) {
        currentStatus = 'SHIFT_COMPLETED';
      } else {
        currentStatus = 'NOT_CHECKED_IN';
      }
    }

    // Process all today's logs into paired shifts
    const todayLogs = userLogs.filter(log => String(log.timestamp).startsWith(todayStr));
    const todayShifts = [];
    let totalTodayMinutes = 0;
    let pendingIn = null;

    todayLogs.forEach(log => {
      if (log.checkType === 'CHECK_IN') {
        if (pendingIn) {
          todayShifts.push({
            sessionId: pendingIn.sessionId,
            storeId: pendingIn.storeId,
            storeName: pendingIn.storeName,
            checkInTime: this._formatTimeOnly(pendingIn.timestamp),
            checkInTimestamp: pendingIn.timestamp,
            checkInPhotoUrl: pendingIn.selfieDriveUrl,
            checkInDistance: pendingIn.distanceMeters,
            checkOutTime: "Missing Check-Out",
            checkOutTimestamp: "",
            checkOutPhotoUrl: "",
            checkOutDistance: null,
            durationMinutes: 0,
            durationFormatted: "Incomplete (Review Required)",
            status: "INCOMPLETE",
            verificationStatus: pendingIn.verificationStatus
          });
        }
        pendingIn = log;
      } else if (log.checkType === 'CHECK_OUT') {
        if (pendingIn) {
          const inDate = new Date(pendingIn.timestamp);
          const outDate = new Date(log.timestamp);
          const dur = Math.max(0, Math.round((outDate.getTime() - inDate.getTime()) / 60000));
          totalTodayMinutes += dur;

          todayShifts.push({
            sessionId: log.sessionId || pendingIn.sessionId,
            storeId: log.storeId || pendingIn.storeId,
            storeName: log.storeName || pendingIn.storeName,
            checkInTime: this._formatTimeOnly(pendingIn.timestamp),
            checkInTimestamp: pendingIn.timestamp,
            checkInPhotoUrl: pendingIn.selfieDriveUrl,
            checkInDistance: pendingIn.distanceMeters,
            checkOutTime: this._formatTimeOnly(log.timestamp),
            checkOutTimestamp: log.timestamp,
            checkOutPhotoUrl: log.selfieDriveUrl,
            checkOutDistance: log.distanceMeters,
            durationMinutes: dur,
            durationFormatted: this._formatDuration(dur),
            status: "COMPLETED",
            verificationStatus: log.verificationStatus
          });
          pendingIn = null;
        } else {
          todayShifts.push({
            sessionId: log.sessionId,
            storeId: log.storeId,
            storeName: log.storeName,
            checkInTime: "No Matching In",
            checkInTimestamp: "",
            checkInPhotoUrl: "",
            checkInDistance: null,
            checkOutTime: this._formatTimeOnly(log.timestamp),
            checkOutTimestamp: log.timestamp,
            checkOutPhotoUrl: log.selfieDriveUrl,
            checkOutDistance: log.distanceMeters,
            durationMinutes: 0,
            durationFormatted: "--",
            status: "UNMATCHED_CHECKOUT",
            verificationStatus: log.verificationStatus
          });
        }
      }
    });

    if (pendingIn) {
      const inDate = new Date(pendingIn.timestamp);
      const now = new Date();
      const currElapsedMins = Math.max(0, Math.round((now.getTime() - inDate.getTime()) / 60000));

      todayShifts.push({
        sessionId: pendingIn.sessionId,
        storeId: pendingIn.storeId,
        storeName: pendingIn.storeName,
        checkInTime: this._formatTimeOnly(pendingIn.timestamp),
        checkInTimestamp: pendingIn.timestamp,
        checkInPhotoUrl: pendingIn.selfieDriveUrl,
        checkInDistance: pendingIn.distanceMeters,
        checkOutTime: "On Duty",
        checkOutTimestamp: "",
        checkOutPhotoUrl: "",
        checkOutDistance: null,
        durationMinutes: currElapsedMins,
        durationFormatted: this._formatDuration(currElapsedMins) + " (In Progress)",
        status: "ON_DUTY",
        verificationStatus: pendingIn.verificationStatus
      });
    }

    return {
      status: currentStatus,
      activeShift: activeShift,
      todayShifts: todayShifts,
      totalTodayMinutes: totalTodayMinutes,
      totalTodayFormatted: this._formatDuration(totalTodayMinutes)
    };
  },

  /**
   * Retrieves personal paired attendance history grouped by date
   */
  getPairedUserHistory: function(email, limit) {
    if (!email) return [];
    const sheet = this.getTab(this.TAB_LOGS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const emailNorm = String(email).trim().toLowerCase();
    const userLogs = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim().toLowerCase() === emailNorm) {
        userLogs.push(this._formatLogRow(data[i]));
      }
    }

    if (userLogs.length === 0) return [];

    const dateMap = {};
    const todayStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd");

    userLogs.forEach(log => {
      const datePart = String(log.timestamp).substring(0, 10);
      if (!dateMap[datePart]) dateMap[datePart] = [];
      dateMap[datePart].push(log);
    });

    const sortedDates = Object.keys(dateMap).sort().reverse();
    const maxDays = limit || 15;
    const history = [];

    sortedDates.slice(0, maxDays).forEach(dateStr => {
      const dayLogs = dateMap[dateStr];
      const shifts = [];
      let totalDayMinutes = 0;
      let pendingIn = null;

      dayLogs.forEach(log => {
        if (log.checkType === 'CHECK_IN') {
          if (pendingIn) {
            shifts.push({
              sessionId: pendingIn.sessionId,
              storeId: pendingIn.storeId,
              storeName: pendingIn.storeName,
              checkInTime: this._formatTimeOnly(pendingIn.timestamp),
              checkInTimestamp: pendingIn.timestamp,
              checkInPhotoUrl: pendingIn.selfieDriveUrl,
              checkInDistance: pendingIn.distanceMeters,
              checkOutTime: "Missing Check-Out",
              checkOutTimestamp: "",
              checkOutPhotoUrl: "",
              checkOutDistance: null,
              durationMinutes: 0,
              durationFormatted: "Incomplete (Review Required)",
              status: "INCOMPLETE",
              verificationStatus: pendingIn.verificationStatus
            });
          }
          pendingIn = log;
        } else if (log.checkType === 'CHECK_OUT') {
          if (pendingIn) {
            const inDate = new Date(pendingIn.timestamp);
            const outDate = new Date(log.timestamp);
            const dur = Math.max(0, Math.round((outDate.getTime() - inDate.getTime()) / 60000));
            totalDayMinutes += dur;

            shifts.push({
              sessionId: log.sessionId || pendingIn.sessionId,
              storeId: log.storeId || pendingIn.storeId,
              storeName: log.storeName || pendingIn.storeName,
              checkInTime: this._formatTimeOnly(pendingIn.timestamp),
              checkInTimestamp: pendingIn.timestamp,
              checkInPhotoUrl: pendingIn.selfieDriveUrl,
              checkInDistance: pendingIn.distanceMeters,
              checkOutTime: this._formatTimeOnly(log.timestamp),
              checkOutTimestamp: log.timestamp,
              checkOutPhotoUrl: log.selfieDriveUrl,
              checkOutDistance: log.distanceMeters,
              durationMinutes: dur,
              durationFormatted: this._formatDuration(dur),
              status: "COMPLETED",
              verificationStatus: log.verificationStatus
            });
            pendingIn = null;
          } else {
            shifts.push({
              sessionId: log.sessionId,
              storeId: log.storeId,
              storeName: log.storeName,
              checkInTime: "No Matching In",
              checkInTimestamp: "",
              checkInPhotoUrl: "",
              checkInDistance: null,
              checkOutTime: this._formatTimeOnly(log.timestamp),
              checkOutTimestamp: log.timestamp,
              checkOutPhotoUrl: log.selfieDriveUrl,
              checkOutDistance: log.distanceMeters,
              durationMinutes: 0,
              durationFormatted: "--",
              status: "UNMATCHED_CHECKOUT",
              verificationStatus: log.verificationStatus
            });
          }
        }
      });

      if (pendingIn) {
        const isToday = (dateStr === todayStr);
        const inDate = new Date(pendingIn.timestamp);
        const now = new Date();
        const currElapsedMins = Math.max(0, Math.round((now.getTime() - inDate.getTime()) / 60000));

        shifts.push({
          sessionId: pendingIn.sessionId,
          storeId: pendingIn.storeId,
          storeName: pendingIn.storeName,
          checkInTime: this._formatTimeOnly(pendingIn.timestamp),
          checkInTimestamp: pendingIn.timestamp,
          checkInPhotoUrl: pendingIn.selfieDriveUrl,
          checkInDistance: pendingIn.distanceMeters,
          checkOutTime: isToday ? "On Duty" : "Missing Check-Out",
          checkOutTimestamp: "",
          checkOutPhotoUrl: "",
          checkOutDistance: null,
          durationMinutes: isToday ? currElapsedMins : 0,
          durationFormatted: isToday ? (this._formatDuration(currElapsedMins) + " (On Duty)") : "Incomplete (Review Required)",
          status: isToday ? "ON_DUTY" : "INCOMPLETE",
          verificationStatus: pendingIn.verificationStatus
        });
      }

      const dObj = new Date(dateStr + "T00:00:00");
      const dateFormatted = !isNaN(dObj.getTime())
        ? (dateStr === todayStr ? "Today · " : "") + Utilities.formatDate(dObj, "Asia/Kolkata", "EEEE, d MMM yyyy")
        : dateStr;

      history.push({
        dateStr: dateStr,
        dateFormatted: dateFormatted,
        isToday: (dateStr === todayStr),
        totalDayMinutes: totalDayMinutes,
        totalDayFormatted: this._formatDuration(totalDayMinutes),
        shifts: shifts
      });
    });

    return history;
  },

  /**
   * Advanced Date-Grouped Hierarchical Attendance Ledger for Managers & Super Admins
   */
  getDateGroupedAttendanceLogs: function(callerRole, callerStoreId, callerEmail, filterStoreId, startDate, endDate, sortKey, sortOrder, limit) {
    const sheet = this.getTab(this.TAB_LOGS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, dateGroups: [], totalRecords: 0 };
    }

    const callerEmailNorm = String(callerEmail || "").trim().toLowerCase();
    const todayStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd");

    // Filter logs by caller permissions, store filter, and date range
    const filteredLogs = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const logStoreId = String(row[5] || "").trim();
      const logManagerEmail = String(row[7] || "").trim().toLowerCase();
      const ts = String(row[1] || "");
      const datePart = ts.substring(0, 10);

      // 1. Role Scoping
      let isAllowed = false;
      if (callerRole === "SUPER_ADMIN") {
        if (!filterStoreId || filterStoreId === "ALL" || logStoreId === filterStoreId) {
          isAllowed = true;
        }
      } else if (callerRole === "REGIONAL_MANAGER") {
        const allowedStores = Array.isArray(callerStoreId) 
          ? callerStoreId.map(s => String(s).trim().toUpperCase()) 
          : String(callerStoreId || "").split(',').map(s => s.trim().toUpperCase()).filter(Boolean);
        
        if (filterStoreId && filterStoreId !== "ALL") {
          const filterNorm = String(filterStoreId).trim().toUpperCase();
          if (allowedStores.includes(filterNorm) && logStoreId.toUpperCase() === filterNorm) {
            isAllowed = true;
          }
        } else {
          if (allowedStores.includes(logStoreId.toUpperCase())) {
            isAllowed = true;
          }
        }
      } else if (callerRole === "STORE_ADMIN" || callerRole === "USER") {
        if (logStoreId === callerStoreId || logManagerEmail === callerEmailNorm) {
          isAllowed = true;
        }
      }

      if (!isAllowed) continue;

      // 2. Date Range Filter
      if (startDate && datePart < startDate) continue;
      if (endDate && datePart > endDate) continue;

      filteredLogs.push(this._formatLogRow(row));
    }

    if (filteredLogs.length === 0) {
      return { success: true, dateGroups: [], totalRecords: 0 };
    }

    // Group logs by Date
    const dateMap = {};
    filteredLogs.forEach(log => {
      const datePart = String(log.timestamp).substring(0, 10);
      if (!dateMap[datePart]) dateMap[datePart] = [];
      dateMap[datePart].push(log);
    });

    const dateGroups = [];
    const dates = Object.keys(dateMap);

    // Date Sorting
    const isDateAsc = (sortOrder === 'ASC' && (!sortKey || sortKey === 'DATE'));
    dates.sort((a, b) => isDateAsc ? a.localeCompare(b) : b.localeCompare(a));

    dates.forEach(dateStr => {
      const dayLogs = dateMap[dateStr];
      const userMap = {};

      dayLogs.forEach(log => {
        const uEmail = String(log.email).trim().toLowerCase();
        if (!userMap[uEmail]) userMap[uEmail] = [];
        userMap[uEmail].push(log);
      });

      const dayRecords = [];
      let totalDayMinutes = 0;
      let completedCount = 0;
      let onDutyCount = 0;
      let incompleteCount = 0;
      let flaggedCount = 0;
      const presentEmails = new Set();

      Object.keys(userMap).forEach(uEmail => {
        const uLogs = userMap[uEmail];
        presentEmails.add(uEmail);
        let pendingIn = null;

        uLogs.forEach(log => {
          if (log.verificationStatus && log.verificationStatus.indexOf('FLAGGED') > -1) {
            flaggedCount++;
          }

          if (log.checkType === 'CHECK_IN') {
            if (pendingIn) {
              incompleteCount++;
              dayRecords.push({
                sessionId: pendingIn.sessionId,
                email: pendingIn.email,
                fullName: pendingIn.fullName,
                employeeCode: pendingIn.employeeCode,
                storeId: pendingIn.storeId,
                storeName: pendingIn.storeName,
                checkInTime: this._formatTimeOnly(pendingIn.timestamp),
                checkInTimestamp: pendingIn.timestamp,
                checkInPhotoUrl: pendingIn.selfieDriveUrl,
                checkInDistance: pendingIn.distanceMeters,
                checkOutTime: "Missing Check-Out",
                checkOutTimestamp: "",
                checkOutPhotoUrl: "",
                checkOutDistance: null,
                durationMinutes: 0,
                durationFormatted: "Incomplete (Review Required)",
                status: "INCOMPLETE",
                verificationStatus: pendingIn.verificationStatus
              });
            }
            pendingIn = log;
          } else if (log.checkType === 'CHECK_OUT') {
            if (pendingIn) {
              const inDate = new Date(pendingIn.timestamp);
              const outDate = new Date(log.timestamp);
              const dur = Math.max(0, Math.round((outDate.getTime() - inDate.getTime()) / 60000));
              totalDayMinutes += dur;
              completedCount++;

              dayRecords.push({
                sessionId: log.sessionId || pendingIn.sessionId,
                email: log.email,
                fullName: log.fullName || pendingIn.fullName,
                employeeCode: log.employeeCode || pendingIn.employeeCode,
                storeId: log.storeId || pendingIn.storeId,
                storeName: log.storeName || pendingIn.storeName,
                checkInTime: this._formatTimeOnly(pendingIn.timestamp),
                checkInTimestamp: pendingIn.timestamp,
                checkInPhotoUrl: pendingIn.selfieDriveUrl,
                checkInDistance: pendingIn.distanceMeters,
                checkOutTime: this._formatTimeOnly(log.timestamp),
                checkOutTimestamp: log.timestamp,
                checkOutPhotoUrl: log.selfieDriveUrl,
                checkOutDistance: log.distanceMeters,
                durationMinutes: dur,
                durationFormatted: this._formatDuration(dur),
                status: "COMPLETED",
                verificationStatus: log.verificationStatus
              });
              pendingIn = null;
            } else {
              dayRecords.push({
                sessionId: log.sessionId,
                email: log.email,
                fullName: log.fullName,
                employeeCode: log.employeeCode,
                storeId: log.storeId,
                storeName: log.storeName,
                checkInTime: "No Matching In",
                checkInTimestamp: "",
                checkInPhotoUrl: "",
                checkInDistance: null,
                checkOutTime: this._formatTimeOnly(log.timestamp),
                checkOutTimestamp: log.timestamp,
                checkOutPhotoUrl: log.selfieDriveUrl,
                checkOutDistance: log.distanceMeters,
                durationMinutes: 0,
                durationFormatted: "--",
                status: "UNMATCHED_CHECKOUT",
                verificationStatus: log.verificationStatus
              });
            }
          }
        });

        if (pendingIn) {
          const isToday = (dateStr === todayStr);
          if (isToday) onDutyCount++;
          else incompleteCount++;

          const inDate = new Date(pendingIn.timestamp);
          const now = new Date();
          const currElapsedMins = Math.max(0, Math.round((now.getTime() - inDate.getTime()) / 60000));

          dayRecords.push({
            sessionId: pendingIn.sessionId,
            email: pendingIn.email,
            fullName: pendingIn.fullName,
            employeeCode: pendingIn.employeeCode,
            storeId: pendingIn.storeId,
            storeName: pendingIn.storeName,
            checkInTime: this._formatTimeOnly(pendingIn.timestamp),
            checkInTimestamp: pendingIn.timestamp,
            checkInPhotoUrl: pendingIn.selfieDriveUrl,
            checkInDistance: pendingIn.distanceMeters,
            checkOutTime: isToday ? "On Duty" : "Missing Check-Out",
            checkOutTimestamp: "",
            checkOutPhotoUrl: "",
            checkOutDistance: null,
            durationMinutes: isToday ? currElapsedMins : 0,
            durationFormatted: isToday ? (this._formatDuration(currElapsedMins) + " (In Progress)") : "Incomplete (Review Required)",
            status: isToday ? "ON_DUTY" : "INCOMPLETE",
            verificationStatus: pendingIn.verificationStatus
          });
        }
      });

      // Sort day records
      if (sortKey === 'NAME') {
        dayRecords.sort((a, b) => {
          const nameA = (a.fullName || a.email).toLowerCase();
          const nameB = (b.fullName || b.email).toLowerCase();
          return sortOrder === 'DESC' ? nameB.localeCompare(nameA) : nameA.localeCompare(nameB);
        });
      } else if (sortKey === 'DURATION') {
        dayRecords.sort((a, b) => {
          return sortOrder === 'ASC' ? a.durationMinutes - b.durationMinutes : b.durationMinutes - a.durationMinutes;
        });
      } else {
        // Default TIME (Check-In time)
        dayRecords.sort((a, b) => {
          const tA = a.checkInTimestamp || "";
          const tB = b.checkInTimestamp || "";
          return sortOrder === 'ASC' ? tA.localeCompare(tB) : tB.localeCompare(tA);
        });
      }

      const dObj = new Date(dateStr + "T00:00:00");
      const dateFormatted = !isNaN(dObj.getTime())
        ? (dateStr === todayStr ? "Today · " : "") + Utilities.formatDate(dObj, "Asia/Kolkata", "EEEE, d MMM yyyy")
        : dateStr;

      dateGroups.push({
        dateStr: dateStr,
        dateFormatted: dateFormatted,
        isToday: (dateStr === todayStr),
        kpis: {
          totalPresent: presentEmails.size,
          completedShifts: completedCount,
          activeOnDuty: onDutyCount,
          incompleteShifts: incompleteCount,
          flaggedPunches: flaggedCount,
          totalHours: Math.round((totalDayMinutes / 60) * 10) / 10,
          totalDurationFormatted: this._formatDuration(totalDayMinutes)
        },
        records: dayRecords
      });
    });

    return {
      success: true,
      startDate: startDate || "",
      endDate: endDate || "",
      sortKey: sortKey || "DATE",
      sortOrder: sortOrder || "DESC",
      dateGroups: dateGroups,
      totalRecords: filteredLogs.length
    };
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
      } else if (callerRole === "REGIONAL_MANAGER") {
        const allowedStores = Array.isArray(callerStoreId) 
          ? callerStoreId.map(s => String(s).trim().toUpperCase()) 
          : String(callerStoreId || "").split(',').map(s => s.trim().toUpperCase()).filter(Boolean);
        if (allowedStores.includes(userStoreId.toUpperCase()) || userManager === callerEmailNorm) {
          isAllowed = true;
        }
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

  _formatTimeOnly: function(isoOrDateStr) {
    if (!isoOrDateStr) return "--:--";
    const d = new Date(isoOrDateStr);
    if (isNaN(d.getTime())) return String(isoOrDateStr);
    return Utilities.formatDate(d, "Asia/Kolkata", "hh:mm a");
  },

  _formatDuration: function(minutes) {
    if (minutes === null || minutes === undefined || isNaN(minutes) || minutes < 0) return "--";
    const hrs = Math.floor(minutes / 60);
    const mins = minutes % 60;
    if (hrs === 0) return mins + "m";
    return hrs + "h " + (mins < 10 ? "0" : "") + mins + "m";
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
      verificationStatus: r[15],
      sessionId: r[16] || "",
      workDurationMinutes: (r[17] !== undefined && r[17] !== "" && r[17] !== null) ? Number(r[17]) : null
    };
  }
};
