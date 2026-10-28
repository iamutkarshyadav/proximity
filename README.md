# PROXIMITY // Enterprise Geofenced Attendance Platform
### Colorful Neo-Brutalist Serverless Web Application on Google Apps Script, Google Sheets & Google Drive

---

## ⚡ Architecture & Tech Stack Overview
- **Runtime:** Google Apps Script (`HtmlService` V8 Runtime)
- **Database:** Google Sheets (`SpreadsheetApp`) with dynamic self-healing schema initialization & `LockService` for shift-change write concurrency protection
- **Media Storage:** Google Drive (`DriveApp`) with automatic store subfolders (`Attendance_Selfies_Archive / [Store_Name]`)
- **100% In-App Photo Rendering:** Direct server-side Base64 streaming via `getSelfieBase64`—eliminates CORS and external Drive redirects
- **Identity & 3-Tier RBAC:** Strict separation of duties:
  1. `SUPER_ADMIN`: Global live monitoring, KPIs, store CRUD & user access management.
  2. `STORE_ADMIN`: Scoped store live feed, store KPIs & staff roster assignment.
  3. `USER`: Ground staff attendance marking with GPS proximity verification and watermarked photo submission.
- **Frontend Aesthetic:** Colorful Neo-Brutalism (Warm Cream canvas `#FFFDF7` with dot matrix, Sunny Amber header `#FEF08A`, Acid Mint & Coral status badges, Deep Indigo camera viewport, tactile hard shadows, zero mobile horizontal overflow).

---

## 📁 Repository Structure

```
attend/
├── appsscript.json             # Manifest: OAuth Scopes, V8 runtime, web app configuration
├── Code.gs                     # Server Controller: getSelfieBase64, 3-role routing, CRUD endpoints, attendance submission
├── Database.gs                 # Data Access Layer: self-healing schema sync, dynamic Super Admin seeding, KPI calculation
├── GeoService.gs               # Server-Side Haversine Verification with tolerance buffer
├── DriveService.gs             # Media Service: auto-creates root archive & store subfolders
│
├── Index.html                  # Core HTML5 Shell, dynamic top bar, in-app photo modal & admin dialogs
├── Styles.html                 # Vibrant Colorful Neo-Brutalist Design Tokens & Responsive Utilities
├── State.html                  # Centralized Reactive State Store & Subscriber Pattern
├── GeoTracker.html             # Geolocation Watcher, Drift Smoothing & Proximity Engine
├── CameraWatermark.html        # Front Camera Capture & Unclipped Canvas Watermark Stamper
├── UserDashboard.html          # Ground Staff Interface (Check In / Out, Blueprint Camera, Personal Activity)
├── StoreAdminDashboard.html    # Store Admin Interface (Banner, KPIs, Store Attendance, Staff Roster)
├── SuperAdminDashboard.html    # Master Admin Interface (KPIs, All Attendance, Stores & Admins CRUD, All Users)
├── App.html                    # Application Bootstrap, 3-Role Router, In-App Photo Viewer & Modal Handlers
└── README.md                   # Documentation
```

---

## 🗄️ Database Schema (`Mother_Hub_Attendance_Database`)

The system automatically initializes and formats three core sheets with frozen, dark `#18181B` Neo-Brutalist headers:

### 1. `Stores` Tab (Locations & Geofences)
`Store_ID` | `Store_Name` | `Latitude` | `Longitude` | `Radius_Meters` | `Primary_Admin_Email` | `Status` | `Created_By` | `Created_At`

### 2. `Users` Tab (RBAC Directory)
`User_ID` | `Email` | `Full_Name` | `Employee_Code` | `Role` | `Assigned_Store_ID` | `Reporting_Manager_Email` | `Status` | `Created_At` | `Last_Login_At`

### 3. `Attendance_Logs` Tab (Audit Log)
`Log_ID` | `Timestamp` | `Email` | `Full_Name` | `Employee_Code` | `Store_ID` | `Store_Name` | `Reporting_Manager_Email` | `Check_Type` | `User_Lat` | `User_Lng` | `Calculated_Distance_Meters` | `GPS_Accuracy_Meters` | `Device_Type` | `Selfie_Drive_URL` | `Verification_Status`
