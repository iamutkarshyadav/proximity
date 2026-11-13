/**
 * Code.gs
 * Application Controller & 3-Tier Hierarchical Routing
 * Multi-Store Enterprise Attendance Platform (Desktop-Optimized, Zero Emojis)
 */

const APP_CONFIG = {
  APP_NAME: "Proximity",
  DEFAULT_RADIUS_METERS: 100,
  GPS_TOLERANCE_BUFFER_METERS: 25,
  MAX_LOCK_WAIT_MS: 15000
};

/**
 * Web App Entry Point
 */
function doGet(e) {
  try {
    const template = HtmlService.createTemplateFromFile("Index");
    return template.evaluate()
      .setTitle(APP_CONFIG.APP_NAME)
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput(
      "<div style='font-family:sans-serif;padding:2rem;margin:2rem;border:2px solid #DC2626;background:#FEF2F2;border-radius:8px;'>" +
      "<h2 style='color:#991B1B;'>Application Error</h2>" +
      "<p style='color:#7F1D1D;'><b>Message:</b> " + err.message + "</p>" +
      "<pre style='background:#FFF;padding:1rem;border:1px solid #E5E7EB;overflow:auto;'>" + (err.stack || err) + "</pre>" +
      "</div>"
    ).setTitle("Application Error");
  }
}

/**
 * Scriptlet helper to modularly import sub-templates
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Dedicated function to seed & format sheet columns
 */
function seedSheetHeaders() {
  return Database.ensureInitialized();
}

/**
 * Retrieves authenticated session email.
/**
 * Retrieves authenticated session email.
 * Checks Session.getActiveUser() first; if blank (mobile with "Anyone" deployment),
 * falls back to clientProvidedEmail. Does NOT fall back to getEffectiveUser().
 */
function getAuthenticatedEmail(clientProvidedEmail) {
  let email = "";
  try {
    email = Session.getActiveUser().getEmail();
  } catch (e) {
    email = "";
  }

  if (!email && clientProvidedEmail) {
    email = String(clientProvidedEmail).trim().toLowerCase();
  }

  return (email || "").trim().toLowerCase();
}

/**
 * Helper: Parses raw lat,lng or Google Maps share links
 */
function parseCoordinatesOrLink(rawInput) {
  if (!rawInput) return null;
  const text = String(rawInput).trim();

  // 1. Direct Lat, Lng coordinates (e.g. "12.9698, 77.7500" or "12.9698 77.7500")
  const coordRegex = /^(-?\d{1,2}\.\d+)[,\s]+(-?\d{1,3}\.\d+)$/;
  const coordMatch = text.match(coordRegex);
  if (coordMatch) {
    return {
      latitude: parseFloat(coordMatch[1]),
      longitude: parseFloat(coordMatch[2])
    };
  }

  // 2. Google Maps URL with /@lat,lng
  const urlAtRegex = /@(-?\d{1,2}\.\d+),(-?\d{1,3}\.\d+)/;
  const urlAtMatch = text.match(urlAtRegex);
  if (urlAtMatch) {
    return {
      latitude: parseFloat(urlAtMatch[1]),
      longitude: parseFloat(urlAtMatch[2])
    };
  }

  // 3. Google Maps query parameter (?q=lat,lng)
  const queryRegex = /[?&]q=(-?\d{1,2}\.\d+),(-?\d{1,3}\.\d+)/;
  const queryMatch = text.match(queryRegex);
  if (queryMatch) {
    return {
      latitude: parseFloat(queryMatch[1]),
      longitude: parseFloat(queryMatch[2])
    };
  }

  // 4. Fallback search for any pair of floating point numbers in the string
  const anyCoords = text.match(/(-?\d{1,2}\.\d{3,})[^\d-]+(-?\d{1,3}\.\d{3,})/);
  if (anyCoords) {
    return {
      latitude: parseFloat(anyCoords[1]),
      longitude: parseFloat(anyCoords[2])
    };
  }

  return null;
}

/**
 * Dedicated In-App Photo Rendering Server Endpoint:
 * Extracts Google Drive File ID, retrieves blob, and returns base64 Data URI
 * Completely eliminates broken image icons and external Drive redirects.
 */
function getSelfieBase64(fileIdOrUrl) {
  if (!fileIdOrUrl) throw new Error("No photo available.");
  const match = String(fileIdOrUrl).match(/[-\w]{25,}/);
  if (!match) throw new Error("Invalid photo ID.");
  const file = DriveApp.getFileById(match[0]);
  const blob = file.getBlob();
  const base64 = Utilities.base64Encode(blob.getBytes());
  return "data:" + blob.getContentType() + ";base64," + base64;
}

/**
 * Hybrid Identity Resolution Bootstrapper:
 * 1. Checks serverEmail = Session.getActiveUser().getEmail()
 * 2. If empty (mobile browser with "Who has access: Anyone"), falls back to clientEmailHint
 * 3. Normalizes email and runs strict role lookup against Users sheet (or SUPER_ADMIN_SEEDS)
 * 4. Only users with SUPER_ADMIN, STORE_ADMIN, or USER roles can access the app
 */
function getInitialSessionData(clientEmailHint) {
  try {
    let serverEmail = "";
    try {
      serverEmail = Session.getActiveUser().getEmail();
    } catch (e) {
      serverEmail = "";
    }

    let email = "";
    if (serverEmail && serverEmail.trim()) {
      email = serverEmail.trim().toLowerCase();
    } else if (clientEmailHint && String(clientEmailHint).trim()) {
      email = String(clientEmailHint).trim().toLowerCase();
    }

    const serverTime = new Date().toISOString();

    if (!email) {
      return {
        success: false,
        requireEmail: true,
        message: "Please enter your workplace corporate email."
      };
    }

    // 1. Dynamic, self-healing DB initialization
    Database.ensureInitialized();
    try {
      DriveService.getRootFolder();
    } catch (driveErr) {
      Logger.log("Drive folder notice: " + driveErr.message);
    }

    let user = Database.findUserByEmail(email);

    // 2. Strict Role Enforcement: ONLY pre-provisioned users in Users sheet can enter
    if (!user) {
      return {
        success: false,
        requireEmail: true,
        unauthorized: true,
        notRegistered: true,
        email: email,
        message: "Access Denied: '" + email + "' is not registered in the company database."
      };
    }

    if (user && (user.status === "INACTIVE" || user.status === "FALSE" || user.status === false)) {
      return {
        success: false,
        requireEmail: true,
        unauthorized: true,
        notRegistered: true,
        email: email,
        message: "Account Inactive: '" + email + "' has been marked inactive. Please contact Super Admin."
      };
    }

    Database.updateLastLogin(email);

    const todayStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd");

    // 3. Strict 3-Role Data Scoping: SUPER_ADMIN | STORE_ADMIN | USER
    if (user.role === "SUPER_ADMIN") {
      const kpis = Database.getGlobalKpis();
      const stores = Database.getAllStores(true); // All stores with multi-managers
      const dateAttendance = Database.getDateGroupedAttendanceLogs("SUPER_ADMIN", "", "", "ALL", todayStr, todayStr, "TIME", "DESC", 100);
      const allUsers = Database.getAllUsers();

      return {
        success: true,
        role: "SUPER_ADMIN",
        user: user,
        kpis: kpis,
        stores: stores,
        dateAttendance: dateAttendance,
        allUsers: allUsers,
        serverTime: serverTime
      };
    } else if (user.role === "STORE_ADMIN") {
      let store = Database.getStoreById(user.assignedStoreId);
      if (!store) {
        const activeStores = Database.getActiveStores();
        store = activeStores[0] || null;
        if (store) {
          user = Database.updateUser(email, { assignedStoreId: store.storeId });
        }
      }

      const storeId = store ? store.storeId : "";
      const kpis = Database.getStoreKpis(storeId);
      const dateAttendance = Database.getDateGroupedAttendanceLogs("STORE_ADMIN", storeId, user.email, storeId, todayStr, todayStr, "TIME", "DESC", 100);
      const roster = Database.getScopedUsers("STORE_ADMIN", storeId, user.email);

      return {
        success: true,
        role: "STORE_ADMIN",
        user: user,
        store: store,
        kpis: kpis,
        dateAttendance: dateAttendance,
        roster: roster,
        serverTime: serverTime
      };
    } else {
      // Ground Staff (USER) - Live Shift Tracker & Paired Work History
      const allStores = Database.getActiveStores();
      const shiftStatus = Database.getUserShiftStatus(email);
      const pairedHistory = Database.getPairedUserHistory(email, 15);

      return {
        success: true,
        role: "USER",
        user: user,
        stores: allStores,
        shiftStatus: shiftStatus,
        pairedHistory: pairedHistory,
        serverTime: serverTime
      };
    }
  } catch (err) {
    Logger.log("getInitialSessionData error: " + err.message);
    return {
      success: false,
      requireEmail: true,
      message: "Server Error: " + err.message
    };
  }
}

/**
 * Endpoint to explicitly verify an email against Users sheet
 */
function verifyUserEmail(email) {
  if (!email) return { success: false, message: "Email is required." };
  const cleanEmail = String(email).trim().toLowerCase();
  Database.ensureInitialized();
  const user = Database.findUserByEmail(cleanEmail);

  if (!user) {
    return { success: false, exists: false, message: "Email not registered in directory." };
  }
  return {
    success: true,
    exists: true,
    email: cleanEmail,
    role: user.role,
    fullName: user.fullName
  };
}

/**
 * Super Admin: Refreshes Dashboard Data
 */
function getSuperAdminDashboardData(callerEmail, startDate, endDate, sortKey, sortOrder, filterStoreId) {
  const sessionEmail = getAuthenticatedEmail(callerEmail);
  const caller = Database.findUserByEmail(sessionEmail);
  if (!caller || caller.role !== "SUPER_ADMIN") {
    throw new Error("Super Admin access required.");
  }

  const todayStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd");
  const sDate = startDate || todayStr;
  const eDate = endDate || todayStr;
  const storeFilter = filterStoreId || "ALL";

  return {
    success: true,
    kpis: Database.getGlobalKpis(),
    stores: Database.getAllStores(true),
    dateAttendance: Database.getDateGroupedAttendanceLogs("SUPER_ADMIN", "", "", storeFilter, sDate, eDate, sortKey || "TIME", sortOrder || "DESC", 200),
    allUsers: Database.getAllUsers(),
    serverTime: new Date().toISOString()
  };
}

/**
 * Store Admin: Refreshes Dashboard Data
 */
function getStoreAdminDashboardData(callerEmail, startDate, endDate, sortKey, sortOrder) {
  const sessionEmail = getAuthenticatedEmail(callerEmail);
  const caller = Database.findUserByEmail(sessionEmail);
  if (!caller || (caller.role !== "STORE_ADMIN" && caller.role !== "SUPER_ADMIN")) {
    throw new Error("Store Admin access required.");
  }

  const store = Database.getStoreById(caller.assignedStoreId);
  const storeId = store ? store.storeId : caller.assignedStoreId;
  const kpis = Database.getStoreKpis(storeId);
  const todayStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd");
  const sDate = startDate || todayStr;
  const eDate = endDate || todayStr;

  const dateAttendance = Database.getDateGroupedAttendanceLogs(caller.role, storeId, caller.email, storeId, sDate, eDate, sortKey || "TIME", sortOrder || "DESC", 200);
  const roster = Database.getScopedUsers(caller.role, storeId, caller.email);

  return {
    success: true,
    store: store,
    kpis: kpis,
    dateAttendance: dateAttendance,
    roster: roster,
    serverTime: new Date().toISOString()
  };
}

/**
 * User: Ground Staff Real-Time Sync Endpoint
 */
function getUserDashboardData(callerEmail) {
  const sessionEmail = getAuthenticatedEmail(callerEmail);
  const user = Database.findUserByEmail(sessionEmail);
  if (!user) {
    throw new Error("User access required.");
  }
  const allStores = Database.getActiveStores();
  const shiftStatus = Database.getUserShiftStatus(sessionEmail);
  const pairedHistory = Database.getPairedUserHistory(sessionEmail, 15);

  return {
    success: true,
    user: user,
    stores: allStores,
    shiftStatus: shiftStatus,
    pairedHistory: pairedHistory,
    serverTime: new Date().toISOString()
  };
}

/**
 * Dedicated API Endpoint: Filter Date-Wise Attendance with Multi-Criteria Sorting
 */
function getDateGroupedAttendance(payload) {
  const sessionEmail = getAuthenticatedEmail(payload.callerEmail);
  const caller = Database.findUserByEmail(sessionEmail);
  if (!caller) throw new Error("Unauthorized.");

  const role = caller.role;
  if (role !== "STORE_ADMIN" && role !== "SUPER_ADMIN") {
    throw new Error("Management access required.");
  }

  const storeId = (role === "STORE_ADMIN") ? caller.assignedStoreId : (payload.filterStoreId || "ALL");
  const startDate = payload.startDate ? String(payload.startDate).trim() : "";
  const endDate = payload.endDate ? String(payload.endDate).trim() : "";
  const sortKey = payload.sortKey || "TIME";
  const sortOrder = payload.sortOrder || "DESC";

  return Database.getDateGroupedAttendanceLogs(
    role,
    caller.assignedStoreId,
    caller.email,
    storeId,
    startDate,
    endDate,
    sortKey,
    sortOrder,
    250
  );
}

/**
 * Super Admin: Multi-Manager - Assign / Add Manager to Store
 */
function assignStoreManager(payload) {
  const sessionEmail = getAuthenticatedEmail(payload.callerEmail);
  const caller = Database.findUserByEmail(sessionEmail);
  if (!caller || caller.role !== "SUPER_ADMIN") {
    throw new Error("Super Admin access required.");
  }

  if (!payload.storeId) throw new Error("Store ID is required.");
  if (!payload.managerEmail) throw new Error("Manager Email is required.");

  Database.addStoreManager(
    payload.storeId,
    payload.managerEmail,
    payload.fullName || "",
    payload.employeeCode || "",
    !!payload.isPrimary,
    sessionEmail
  );

  return getSuperAdminDashboardData(sessionEmail);
}

/**
 * Super Admin: Multi-Manager - Remove Manager from Store
 */
function removeStoreManager(payload) {
  const sessionEmail = getAuthenticatedEmail(payload.callerEmail);
  const caller = Database.findUserByEmail(sessionEmail);
  if (!caller || caller.role !== "SUPER_ADMIN") {
    throw new Error("Super Admin access required.");
  }

  if (!payload.storeId) throw new Error("Store ID is required.");
  if (!payload.managerEmail) throw new Error("Manager Email is required.");

  Database.removeStoreManager(payload.storeId, payload.managerEmail, sessionEmail);
  return getSuperAdminDashboardData(sessionEmail);
}

/**
 * Super Admin: Multi-Manager - Set Designated Primary Manager
 */
function setPrimaryStoreManager(payload) {
  const sessionEmail = getAuthenticatedEmail(payload.callerEmail);
  const caller = Database.findUserByEmail(sessionEmail);
  if (!caller || caller.role !== "SUPER_ADMIN") {
    throw new Error("Super Admin access required.");
  }

  if (!payload.storeId) throw new Error("Store ID is required.");
  if (!payload.managerEmail) throw new Error("Manager Email is required.");

  Database.setPrimaryStoreManager(payload.storeId, payload.managerEmail);
  return getSuperAdminDashboardData(sessionEmail);
}

/**
 * Super Admin: Combined "Create Store + Assign Admin"
 */
function createStoreAndAdmin(payload) {
  const sessionEmail = getAuthenticatedEmail(payload.callerEmail);
  const caller = Database.findUserByEmail(sessionEmail);
  if (!caller || caller.role !== "SUPER_ADMIN") {
    throw new Error("Super Admin access required.");
  }

  if (!payload.storeName) throw new Error("Store Name is required.");

  let lat = Number(payload.latitude);
  let lng = Number(payload.longitude);

  if (isNaN(lat) || isNaN(lng) || (!lat && !lng)) {
    const parsed = parseCoordinatesOrLink(payload.locationInput);
    if (!parsed) {
      throw new Error("Invalid location. Click 'Use My Location' or paste valid coordinates / Google Maps link.");
    }
    lat = parsed.latitude;
    lng = parsed.longitude;
  }

  const radiusMeters = parseInt(payload.radiusMeters, 10) || 100;
  const adminEmail = payload.adminEmail ? payload.adminEmail.trim().toLowerCase() : "";
  const adminName = payload.adminName ? payload.adminName.trim() : "";

  // 1. Create store with sequential STR-XXX ID
  const storeId = Database.upsertStore({
    storeName: payload.storeName.trim(),
    latitude: lat,
    longitude: lng,
    radiusMeters: radiusMeters,
    primaryAdminEmail: adminEmail,
    status: "ACTIVE"
  }, sessionEmail);

  // 2. If Store Admin email provided, promote or pre-provision
  if (adminEmail) {
    let existingUser = Database.findUserByEmail(adminEmail);
    if (existingUser) {
      Database.updateUser(adminEmail, {
        role: "STORE_ADMIN",
        assignedStoreId: storeId,
        fullName: adminName || existingUser.fullName,
        status: "ACTIVE"
      });
    } else {
      const empCode = "MGR-" + Math.floor(1000 + Math.random() * 9000);
      Database.insertUser({
        email: adminEmail,
        fullName: adminName || "Store Manager",
        employeeCode: empCode,
        role: "STORE_ADMIN",
        assignedStoreId: storeId,
        reportingManagerEmail: sessionEmail,
        status: "ACTIVE"
      });
    }
  }

  return getSuperAdminDashboardData(sessionEmail);
}

/**
 * Super Admin: Deactivate a store
 */
function deactivateStore(payload) {
  const storeId = typeof payload === 'string' ? payload : (payload && payload.storeId);
  const callerEmail = payload && payload.callerEmail;

  const sessionEmail = getAuthenticatedEmail(callerEmail);
  const caller = Database.findUserByEmail(sessionEmail);
  if (!caller || caller.role !== "SUPER_ADMIN") {
    throw new Error("Super Admin access required.");
  }

  Database.deactivateStore(storeId);
  return getSuperAdminDashboardData(sessionEmail);
}

/**
 * Super Admin: Change Store Admin on Existing Store
 */
function changeStoreAdmin(payload) {
  const sessionEmail = getAuthenticatedEmail(payload.callerEmail);
  const caller = Database.findUserByEmail(sessionEmail);
  if (!caller || caller.role !== "SUPER_ADMIN") {
    throw new Error("Super Admin access required.");
  }

  const storeId = payload.storeId;
  const targetEmail = payload.newAdminEmail ? payload.newAdminEmail.trim().toLowerCase() : "";
  const fullName = payload.fullName ? payload.fullName.trim() : "";

  if (!storeId) throw new Error("Store ID is required.");
  if (!targetEmail) throw new Error("Manager Email is required.");

  const store = Database.getStoreById(storeId);
  if (!store) throw new Error("Store not found.");

  // Update store primary admin
  Database.upsertStore({
    storeId: store.storeId,
    storeName: store.storeName,
    latitude: store.latitude,
    longitude: store.longitude,
    radiusMeters: store.radiusMeters,
    primaryAdminEmail: targetEmail,
    status: store.status
  }, sessionEmail);

  // Promote or insert user as STORE_ADMIN
  let user = Database.findUserByEmail(targetEmail);
  if (user) {
    Database.updateUser(targetEmail, {
      role: "STORE_ADMIN",
      assignedStoreId: storeId,
      fullName: fullName || user.fullName,
      status: "ACTIVE"
    });
  } else {
    const empCode = "MGR-" + Math.floor(1000 + Math.random() * 9000);
    Database.insertUser({
      email: targetEmail,
      fullName: fullName || "Store Manager",
      employeeCode: empCode,
      role: "STORE_ADMIN",
      assignedStoreId: storeId,
      reportingManagerEmail: sessionEmail,
      status: "ACTIVE"
    });
  }

  return getSuperAdminDashboardData(sessionEmail);
}

/**
 * Super Admin: Update User Profile, Role, Store or Status
 */
function updateUserRoleOrStore(payload) {
  const sessionEmail = getAuthenticatedEmail(payload.callerEmail);
  const caller = Database.findUserByEmail(sessionEmail);
  if (!caller || caller.role !== "SUPER_ADMIN") {
    throw new Error("Super Admin access required.");
  }

  if (!payload.email) throw new Error("User email is required.");
  const targetEmail = payload.email.trim().toLowerCase();

  const updateFields = {};
  if (payload.role !== undefined) updateFields.role = payload.role;
  if (payload.assignedStoreId !== undefined) updateFields.assignedStoreId = payload.assignedStoreId;
  if (payload.status !== undefined) updateFields.status = payload.status;
  if (payload.fullName !== undefined) updateFields.fullName = payload.fullName;

  Database.updateUser(targetEmail, updateFields);
  return getSuperAdminDashboardData(sessionEmail);
}

/**
 * Store Admin: Add Staff Member to their assigned Store
 */
function addStoreStaff(payload) {
  const sessionEmail = getAuthenticatedEmail(payload.callerEmail);
  const caller = Database.findUserByEmail(sessionEmail);

  if (!caller || (caller.role !== "STORE_ADMIN" && caller.role !== "SUPER_ADMIN")) {
    throw new Error("Unauthorized.");
  }

  if (!payload.email) throw new Error("Staff email is required.");
  const targetEmail = payload.email.trim().toLowerCase();

  const existing = Database.findUserByEmail(targetEmail);
  if (existing) {
    Database.updateUser(targetEmail, {
      assignedStoreId: caller.assignedStoreId,
      reportingManagerEmail: caller.email,
      fullName: payload.fullName ? payload.fullName.trim() : existing.fullName,
      status: "ACTIVE"
    });
  } else {
    const empCode = payload.employeeCode ? payload.employeeCode.trim().toUpperCase() : ("EMP-" + Math.floor(1000 + Math.random() * 9000));
    Database.insertUser({
      email: targetEmail,
      fullName: payload.fullName ? payload.fullName.trim() : "",
      employeeCode: empCode,
      role: "USER",
      assignedStoreId: caller.assignedStoreId,
      reportingManagerEmail: caller.email,
      status: "ACTIVE"
    });
  }

  return getStoreAdminDashboardData(sessionEmail);
}

/**
 * Store Admin: Remove/Suspend Staff from their Store
 */
function removeStaffFromStore(payload) {
  const targetEmail = typeof payload === 'string' ? payload : (payload && payload.email);
  const callerEmail = payload && payload.callerEmail;

  const sessionEmail = getAuthenticatedEmail(callerEmail);
  const caller = Database.findUserByEmail(sessionEmail);

  if (!caller || (caller.role !== "STORE_ADMIN" && caller.role !== "SUPER_ADMIN")) {
    throw new Error("Unauthorized.");
  }

  if (!targetEmail) throw new Error("Target email required.");

  Database.updateUser(String(targetEmail).trim().toLowerCase(), {
    status: "INACTIVE"
  });

  return getStoreAdminDashboardData(sessionEmail);
}

/**
 * Core Verification Endpoint (Ground Staff ONLY):
 * Validates distance, saves watermarked photo in store's subfolder, records atomic log
 */
function submitAttendanceLog(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(APP_CONFIG.MAX_LOCK_WAIT_MS);
  } catch (e) {
    throw new Error("Write collision prevented. Please tap submit again.");
  }

  try {
    const sessionEmail = getAuthenticatedEmail(payload.email);
    if (!sessionEmail) throw new Error("Session expired. Please refresh the page.");

    const user = Database.findUserByEmail(sessionEmail);
    if (!user) throw new Error("User profile not found.");

    // Autocaptured store from client payload or user's assigned store
    const storeIdToUse = payload.storeId || user.assignedStoreId;
    const store = Database.getStoreById(storeIdToUse);

    if (!store) throw new Error("Mother Hub record not found.");
    const isStoreActive = !["INACTIVE", "FALSE", false].includes(String(store.status || "").trim().toUpperCase());
    if (!isStoreActive) throw new Error("Selected Mother Hub is currently inactive.");

    const userLat = Number(payload.userLat);
    const userLng = Number(payload.userLng);
    const accuracy = Number(payload.accuracyMeters) || 0;
    const deviceType = String(payload.deviceType || "MOBILE").toUpperCase();

    // Verify distance with server-side Haversine
    const geoResult = GeoService.verifyGeofence(
      userLat,
      userLng,
      store.latitude,
      store.longitude,
      store.radiusMeters,
      APP_CONFIG.GPS_TOLERANCE_BUFFER_METERS
    );

    if (!geoResult.isInside) {
      throw new Error("Location outside store boundary (" + Math.round(geoResult.distanceMeters) + "m away from " + store.storeName + ", limit is " + store.radiusMeters + "m).");
    }

    let verificationStatus = "VERIFIED";
    if (accuracy > 50 || (deviceType === "LAPTOP" && accuracy > 35)) {
      verificationStatus = "FLAGGED_LOW_ACCURACY";
    }

    // Save watermarked photo to Google Drive in store-specific subfolder
    const selfieRecord = DriveService.saveSelfieBlob(
      payload.base64Selfie,
      store.storeId,
      sessionEmail,
      payload.checkType,
      store.storeName
    );

    // Write to Attendance_Logs
    const logId = Database.appendAttendanceLog({
      email: sessionEmail,
      fullName: user.fullName,
      employeeCode: user.employeeCode,
      storeId: store.storeId,
      storeName: store.storeName,
      reportingManagerEmail: user.reportingManagerEmail || "",
      checkType: payload.checkType || "CHECK_IN",
      userLat: userLat,
      userLng: userLng,
      distanceMeters: geoResult.distanceMeters,
      accuracyMeters: accuracy,
      deviceType: deviceType,
      selfieDriveUrl: selfieRecord.viewUrl,
      verificationStatus: verificationStatus
    });

    const shiftStatus = Database.getUserShiftStatus(sessionEmail);
    const pairedHistory = Database.getPairedUserHistory(sessionEmail, 15);

    return {
      success: true,
      logId: logId,
      distanceMeters: geoResult.distanceMeters,
      verificationStatus: verificationStatus,
      checkType: payload.checkType,
      storeName: store.storeName,
      selfieUrl: selfieRecord.viewUrl,
      timestamp: new Date().toISOString(),
      shiftStatus: shiftStatus,
      pairedHistory: pairedHistory
    };
  } finally {
    lock.releaseLock();
  }
}
