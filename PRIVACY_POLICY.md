# Privacy Policy for Colony Fund

**Effective Date:** October 27, 2026  
**Last Updated:** October 27, 2026  
**Application Name:** Colony Fund (Android App)

---

## 1. Introduction & Overview

**Colony Fund** is a specialized financial management and colony administration mobile application developed for residential welfare societies, apartment associations, and gated communities. 

Our core philosophy is **complete user privacy and data ownership**. Unlike commercial apps, Colony Fund does **not** collect your data on centralized third-party servers, does **not** sell personal information, and contains **zero advertisements or commercial trackers**.

---

## 2. Data Storage & Architecture

### Private Society Cloud (Google Sheets & Apps Script)
All records—including resident rosters, monthly welfare contributions, emergency loan ledgers, maintenance receipts, and expense logs—are stored exclusively in the private **Google Sheet** owned and maintained by your society's elected Treasurer or Admin committee.
- Data transmits directly between your Android device and your society’s private Google Apps Script Web API via secure HTTPS encryption (`TLS 1.3`).
- No intermediary proxy, analytics server, or external database ever intercepts or stores your community records.

### Local Device Storage
The app caches certain preferences locally on your phone for fast offline viewing:
- Selected language preference (English / Telugu).
- Preferred display theme (Dark Mode / Light Mode).
- Authentication session token / active profile identifier.

---

## 3. Permissions Requested & Why They Are Needed

Colony Fund operates on a strict principle of minimal device permissions:

| Permission | Purpose | Data Transmitted? |
| :--- | :--- | :--- |
| **Biometric (Fingerprint / Face)** | Allows quick, secure sign-in using Android’s standard `BiometricPrompt` API. | **Never.** Biometrics are processed entirely inside your device's hardware security module (TEE/Titan M) and never leave your phone. |
| **Camera** | Used exclusively on-demand to scan the society's **Google Sheet Connect QR code** during initial setup. | **No.** The camera stream is analyzed locally in real-time; no photos or video recordings are saved or uploaded. |
| **Notifications** | Delivers timely community announcements, meeting reminders, poll alerts, and monthly dues reminders scheduled by the society committee. | **No.** Notifications are triggered either locally or via your society's private Apps Script trigger. |
| **Storage / Files** | Allows you to save and export official PDF payment receipts, passbook statements, and account summaries to your device's `Downloads` folder. | **No.** Exported PDF files remain strictly on your local device. |

---

## 4. Financial & Payment Privacy

- **UPI Transactions:** When making contribution or maintenance payments, Colony Fund generates standard NPCI-compliant UPI payment intents (`upi://pay`) and delegates the transaction to your chosen banking app (Google Pay, PhonePe, Paytm, BHIM, CRED, or mobile banking).
- **No Card or PIN Data:** Colony Fund **never** asks for, accesses, or stores your bank account numbers, debit/credit card numbers, UPI PINs, or net banking passwords.
- **UTR / Reference Numbers:** When you confirm a payment, you may voluntarily record the transaction reference number (UTR) to assist your society treasurer with account reconciliation.

---

## 5. Third-Party Sharing & Advertising

- **Zero Ads:** Colony Fund is 100% ad-free.
- **Zero Third-Party Trackers:** We do not embed Google Analytics, Firebase Analytics, Facebook SDK, or any third-party marketing SDKs.
- **WhatsApp Sharing:** The app offers convenient shortcuts to share formatted accounts summaries or receipts via WhatsApp. Message transmission is handled entirely within WhatsApp according to WhatsApp's own privacy terms.

---

## 6. Data Deletion & Member Exit

- When a resident leaves or sells their flat, the society administrator can record a **Member Settlement** in the admin console.
- If you uninstall the app, all locally cached data and preferences on your device are deleted automatically.
- To request deletion of historical records from your colony's Google Sheet, please contact your society's Treasurer or Executive Committee directly.

---

## 7. Open Source & Transparency

The Google Apps Script backend and Android distribution files are openly auditable in this public repository. Community members and developers are encouraged to inspect the source code and verify our privacy commitments.

---

## 8. Contact & Queries

For any questions regarding this Privacy Policy or society data management, please contact your colony's administration committee or submit an issue in this repository:
- **Repository:** [https://github.com/knarsimharao/colony-fund-app](https://github.com/knarsimharao/colony-fund-app)
