# Colony Fund - Deployment & Setup Guide for Societies

This guide explains how any **Apartment Association, Gated Community, or Colony Welfare Society** can deploy **Colony Fund** in under 5 minutes with **zero server costs**, using Google Sheets and Google Apps Script as the complete backend!

---

## 🏛️ How the Architecture Works

```text
┌───────────────────────────────┐           ┌──────────────────────────────────────────────┐
│       Colony Fund App         │  HTTPS    │          Google Apps Script API              │
│     (Android Mobile App)      │◄─────────►│               (Code.gs)                      │
│                               │           └──────────────────────┬───────────────────────┘
│  • Resident Member Portal     │                                  │ Reads & Writes
│  • Admin / Treasurer Console  │                                  ▼
│  • Biometric & Offline Mode   │           ┌──────────────────────────────────────────────┐
│  • 1-Tap UPI Payment Intents  │           │            Society Google Sheet              │
│  • PDF Receipts & Passbooks   │           │ (Members, Loans, Contributions, Maintenance) │
└───────────────────────────────┘           └──────────────────────────────────────────────┘
```

1. **Free Cloud Database:** Your Google Sheet serves as the private database.
2. **Serverless REST API:** Google Apps Script (`Code.gs`) handles authentication, balance recalculation, loan amortization, payment verification, and automated nightly backups.
3. **No Recurring Cost:** Runs 100% within Google’s free usage tier.

---

## 🚀 Step-by-Step Deployment (5 Minutes)

### Step 1: Create a New Google Sheet
1. Open your browser and go to [sheets.new](https://sheets.new) (make sure you are logged in with your Society's official Google account).
2. Name the sheet (e.g., `Tirumala Colony Sahayanidhi 2026` or `Your Colony Name Fund`).

---

### Step 2: Open the Apps Script Editor
1. In the Google Sheets top menu, click:  
   **Extensions** ➔ **Apps Script**.
2. A new tab will open with an empty `Code.gs` editor. Delete any default placeholder code inside it (`function myFunction() { ... }`).

---

### Step 3: Paste the Backend Code
1. Open the [**`google-apps-script/Code.gs`**](./google-apps-script/Code.gs) file from this repository.
2. Copy all code and paste it directly into your Apps Script editor.
3. Click the **Save** icon (disk icon) or press `Ctrl + S`.

---

### Step 4: Run Initial Setup
1. In the Apps Script toolbar dropdown (which says `Select function`), choose `setupSpreadsheet`.
2. Click **Run**.
3. Google will show an authorization prompt:
   - Click **Review permissions**.
   - Select your Google account.
   - Click **Advanced** (at the bottom) ➔ **Go to Untitled project (unsafe)**.
   - Click **Allow**.
4. The script will automatically generate all required sheets with professional column formatting, formulas, and starter data:
   - `Members`
   - `MonthlyContributions`
   - `Loans`
   - `LoanPayments`
   - `MaintenanceCollections`
   - `ColonyExpenses`
   - `SocietyBalance`
   - `Settings`
   - `Announcements` & `Polls`

---

### Step 5: Deploy as a Web App
1. At the top-right of the Apps Script window, click the blue **Deploy** button ➔ **New deployment**.
2. Click the gear icon (⚙️) next to *Select type* and choose **Web app**.
3. Configure the settings:
   - **Description:** `Colony Fund Backend API`
   - **Execute as:** `Me (your-email@gmail.com)`
   - **Who has access:** `Anyone` *(Crucial: allows mobile app to communicate securely with your sheet)*.
4. Click **Deploy**.
5. Copy the generated **Web app URL** (it starts with `https://script.google.com/macros/s/.../exec`).

---

### Step 6: Connect the Mobile App
1. Open the **Colony Fund** app on your phone.
2. Go to **Settings** (⚙️) ➔ **Colony Connect & Google Sheet Backend**.
3. Paste your Web App URL into the field.
4. Tap **Connect & Sync**.
5. That’s it! Your entire colony ledger is live and synchronized!

---

## 👥 Managing Members & Admins

In your Google Sheet, open the **`Members`** tab:

| MemberID | Name | HouseNumber | Phone | Role | Persona | Status | Pin |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| `MEM-001` | Mallikarjun | 12-5-36/1 | 9876543210 | **admin** | **dual** | active | 1234 |
| `MEM-002` | Narsimha Rao K | 12-5-36/5 | 9848012345 | **member** | **dual** | active | 1234 |
| `MEM-003` | Resident Name | 12-5-36/12 | 9988776655 | **member** | **maintenance_only**| active | 1234 |

- **Role:** Set to `admin` for Society Treasurers/Presidents to access admin controls, or `member` for regular residents.
- **Persona:**
  - `dual`: Resident participates in both the **Welfare Fund** (₹500/mo) and **Colony Maintenance** (₹100/mo).
  - `welfare_only`: Member participates only in the mutual savings & loan fund.
  - `maintenance_only`: Resident pays only the monthly colony maintenance fee.
- **Pin:** 4-digit PIN for quick resident sign-in. Residents can change their PIN anytime in the app settings.

---

## ⚡ Automated Nightly Backups

`Code.gs` includes a built-in snapshot backup engine:
1. Open your Google Sheet.
2. In the custom top menu (e.g. `🏛️ Your Colony`), click:  
   **⏰ Enable Automated Daily 2 AM Backup**.
3. Every night at 2:00 AM, Google Apps Script will automatically take an immutable timestamped copy of your entire spreadsheet and save it into a dedicated `Backups` folder in your Google Drive!
