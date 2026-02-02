# 📊 Contract & Budget Tracking System (Excel VBA)

A comprehensive Excel VBA–based system for managing contracts, competitions, financial commitments, budget tracking, and claims workflows across multiple departments.

This solution is designed to **automatically synchronize contract data**, **track execution status transitions**, and **reflect financial impacts in real time** across Budget and Claims sheets — all from structured Excel worksheets.

---

## 🚀 Key Features

### 🔗 Centralized Database

* All contracts are stored in a single sheet: **`DB_Contracts`**
* Each contract is assigned a **unique ID (التسلسل)** used as a reference across the system
* Prevents duplicate contract entries automatically

### 🔄 Automatic Status Synchronization

Changing **حالة التنفيذ (Execution Status)** in any section sheet automatically:

* Logs transition dates
* Updates the central database
* Reflects changes in Budget and Claims logic

Supported statuses:

* تحت إجراءات التعاقد
* تم التوقيع
* تحت التوريد
* بانتظار سندات الاستلام
* مطالبة مرفوعة للمالية

---

## 🧠 Smart Workflow Logic

### ✍️ When a Contract Is Entered

* Financial commitment amount is counted under **Total Commitments**
* Contract remains in its source section
* No data duplication or deletion occurs

### ✍️ When Status Changes to **تم التوقيع**

* Commitment amount is removed
* Actual contract amount is added under **Signed Contracts**
* Contract appears automatically in **Claims (2025/2026)** based on year
* Signing date and source are logged

### ✍️ When Status Changes to **مطالبة مرفوعة للمالية**

* Amount is reflected under **Raised to Finance**
* Finance date is logged
* Spending ratio is updated
* Number of signed contracts remains unchanged

### 🔄 Rollback Supported

If a contract status is reverted back to **تحت إجراءات التعاقد**:

* All dates and sources are cleared
* Budget figures are recalculated correctly
* Contract is removed from Claims
* Database status is reverted safely

---

## 📁 Project Structure

```text
📦 Excel VBA Project
 ┣ 📄 Section Sheets (Competitions, Direct Purchase, E-Market, O&M, Claims)
 ┣ 📄 DB_Contracts          # Central contracts database
 ┣ 📄 Budget Sheet          # Financial aggregation & KPIs
 ┣ 📄 ThisWorkbook          # Global event handler
 ┗ 📄 VBA Modules
     ┣ 📜 EnterData          # Insert contract into DB
     ┣ 📜 HandleStatusChange# Status transition logic
     ┣ 📜 SyncStatusToDB    # DB synchronization
     ┗ 📜 Helpers           # Header lookup & utilities
```

---

## 🧩 Technical Highlights

* Uses **Workbook_SheetChange** for global event handling
* Relies on **column headers (not column letters)** for robustness
* Event-safe design (prevents `EnableEvents = False` deadlocks)
* Fallback matching by Contract Number if ID is missing
* Modular, readable, and maintainable VBA code

---

## 🛠 How to Use

1. Fill a new contract row in any section sheet
2. Run **`EnterData`** to register it in `DB_Contracts`
3. Change **حالة التنفيذ** directly in the sheet
4. All updates happen automatically:

   * Dates
   * Database
   * Budget
   * Claims

---

## 🧪 Recommended Test Case

1. Enter a contract with status **تحت إجراءات التعاقد**
2. Run `EnterData`
3. Change status to **تم التوقيع**
4. Change status to **مطالبة مرفوعة للمالية**
5. Roll back to **تحت إجراءات التعاقد**

Expected:
✔ Correct budget totals
✔ Accurate database status
✔ Clean rollback with no residual values

---

## 🔐 Notes & Best Practices

* Ensure **macros are enabled** when opening the file
* Avoid renaming column headers unless updated in VBA constants
* If sheets are protected, allow macro editing or unprotect via VBA

---

## 📌 License

This project is intended for internal or organizational use.
You may adapt and extend it as needed.

---

## 👤 Author

Developed for enterprise-level contract and budget tracking using Excel VBA.
