/**
 * DriveService.gs
 * Zero-Touch Cloud Storage Service for Attendance Verification Photos
 * Multi-Store Enterprise Attendance Platform
 */

const DriveService = {
  FOLDER_NAME: "Attendance_Selfies_Archive",

  /**
   * Retrieves the configured Drive archive root folder from Script Properties
   * or auto-creates/locates 'Attendance_Selfies_Archive' in the root Drive.
   * 
   * @return {GoogleAppsScript.Drive.Folder}
   */
  getRootFolder: function() {
    const props = PropertiesService.getScriptProperties();
    let folderId = props.getProperty('SELFIE_FOLDER_ID') || props.getProperty('DRIVE_FOLDER_ID');

    if (folderId) {
      try {
        const folder = DriveApp.getFolderById(folderId);
        if (!folder.isTrashed()) return folder;
      } catch (e) {
        Logger.log("Configured folder ID inaccessible or in trash, auto-recreating: " + e.message);
      }
    }

    // Search for existing folder by name
    const folders = DriveApp.getFoldersByName(this.FOLDER_NAME);
    while (folders.hasNext()) {
      const folder = folders.next();
      if (!folder.isTrashed()) {
        props.setProperty('SELFIE_FOLDER_ID', folder.getId());
        return folder;
      }
    }

    // Auto-create root archive folder
    const newFolder = DriveApp.createFolder(this.FOLDER_NAME);
    newFolder.setDescription("Stores watermarked JPEG verification photos for the Workplace Attendance Platform.");
    try {
      newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (permErr) {
      Logger.log("Domain policy restricted public sharing. Retaining domain-level access.");
    }
    props.setProperty('SELFIE_FOLDER_ID', newFolder.getId());
    return newFolder;
  },

  /**
   * Retrieves or auto-creates a store-specific subfolder within the root archive
   * 
   * @param {string} storeName - e.g. "Koramangala Hub"
   * @return {GoogleAppsScript.Drive.Folder}
   */
  getStoreFolder: function(storeName) {
    const root = this.getRootFolder();
    const cleanName = (storeName || "General_Store").trim();

    const subfolders = root.getFoldersByName(cleanName);
    while (subfolders.hasNext()) {
      const sub = subfolders.next();
      if (!sub.isTrashed()) return sub;
    }

    // Create subfolder for this store
    const newSub = root.createFolder(cleanName);
    try {
      newSub.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {}
    return newSub;
  },

  /**
   * Saves a base64 encoded JPEG data URL as a Drive file in the store's subfolder.
   * 
   * @param {string} base64DataUri - Raw data URI
   * @param {string} storeId - Store ID
   * @param {string} storeName - Store Name
   * @param {string} email - Staff member email
   * @param {string} checkType - "CHECK_IN" or "CHECK_OUT"
   * @return {{fileId: string, viewUrl: string, directImageUrl: string}}
   */
  saveSelfieBlob: function(base64DataUri, storeId, email, checkType, storeName) {
    if (!base64DataUri || typeof base64DataUri !== 'string') {
      throw new Error("Missing or invalid photo data.");
    }

    let cleanBase64 = base64DataUri;
    const match = base64DataUri.match(/^data:image\/[a-z]+;base64,(.+)$/i);
    if (match) {
      cleanBase64 = match[1];
    }

    const decodedBytes = Utilities.base64Decode(cleanBase64);
    const sanitizedEmail = email.replace(/[^a-zA-Z0-9]/g, '_');
    const timestampStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyyMMdd_HHmmss");
    const fileName = `PHOTO_${checkType}_${storeId}_${sanitizedEmail}_${timestampStr}.jpg`;

    const blob = Utilities.newBlob(decodedBytes, "image/jpeg", fileName);
    const targetFolder = this.getStoreFolder(storeName || storeId);
    const file = targetFolder.createFile(blob);

    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {}

    const fileId = file.getId();
    const viewUrl = `https://drive.google.com/file/d/${fileId}/view?usp=drivesdk`;
    const directImageUrl = `https://lh3.googleusercontent.com/d/${fileId}`;

    return {
      fileId: fileId,
      viewUrl: viewUrl,
      directImageUrl: directImageUrl
    };
  }
};
