# SunilP-ExcelMergerToMaster

Automated Power Automate solution to consolidate operational Excel data into a single **daily master file** for downstream ERP ingestion.

![Microsoft Community](https://img.shields.io/badge/Microsoft%20Community-Super%20User-orange?style=for-the-badge&logo=microsoft)
![License](https://img.shields.io/badge/License-MIT-yellow?style=for-the-badge)
![Power Automate](https://img.shields.io/badge/Platform-Power%20Automate-blue?style=for-the-badge&logo=power-automate)
![SharePoint](https://img.shields.io/badge/Storage-SharePoint-036C71?style=for-the-badge&logo=microsoft-sharepoint)

---

## 🔍 Real‑World Scenario (Why this exists)

In an **Oil & Gas operations environment**, field inspectors perform **mandatory safety and equipment inspections multiple times per day** across active sites. These inspections include gas leak checks, pressure readings, valve status, and safety compliance confirmations.

Because inspectors often work in **remote locations with limited connectivity**, they use **approved Excel templates** to capture inspection data offline. Once connectivity is available, completed files are uploaded to a **central SharePoint folder**.

Typical daily volume:
- Up to **60 Excel files**
- Up to **150 rows per file**
- Up to **~9,000 rows per day**

Data must be consolidated **the same day** to support compliance review and **ERP batch ingestion**.

Before automation, a coordinator manually merged dozens of files into a master workbook. This took **1–2 hours per day**, caused rework when files were locked or late, and introduced risk of human error.

**ExcelMergerToMaster** eliminates this manual consolidation step while preserving auditability and data integrity.

---

## 💡 Solution Overview

This solution uses **Power Automate** to consolidate data from **multiple Excel workbooks** stored in SharePoint into a **single daily master Excel file**.

- Runs on a **scheduled basis** aligned to business cut‑off times
- Reads **all Excel tables** from each source workbook
- Appends rows into **one fixed master table**
- Tracks source and processing status
- Handles locked or failed files safely
- Archives processed files nightly

> **Note:** Current implementation uses a **row‑by‑row append** with the Excel Online (Business) connector. A bulk‑insert version using Office Scripts is planned.

---

## ⬇️ Get Started

**Download the latest Power Platform solution package (ZIP) to automate your Excel consolidation:**

[![Download ZIP](https://img.shields.io/badge/Download-Solution-blue?style=for-the-badge&logo=github)](https://github.com/spashikanti/SunilP-ExcelMergerToMaster/releases/latest)
[![Release](https://img.shields.io/github/v/release/spashikanti/SunilP-ExcelMergerToMaster?style=for-the-badge&logo=github&color=brightgreen)](https://github.com/spashikanti/SunilP-ExcelMergerToMaster/releases/latest)
[![Total Downloads](https://img.shields.io/github/downloads/spashikanti/SunilP-ExcelMergerToMaster/total?style=for-the-badge&color=yellow)](https://github.com/spashikanti/SunilP-ExcelMergerToMaster/releases)

---

## 🤔 Why Power Automate

- Files already land in **SharePoint**
- No desktop automation required
- Aligns with **ERP batch processing**
- Delivers value quickly without retraining field users
- Focuses on removing **manual back‑office work**, not redesigning field tooling

Excel remains the intake format. ERP remains the system of record.

---

## 🕒 Processing Model

### Daily Consolidation Flow (Primary)
- **Trigger:** Scheduled (example: 4:00 PM)
- Processes all eligible Excel files in the Source folder
- Produces a **Daily Master file**  
  Example: `Master_Inspection_2026-04-14.xlsx`
- Sends summary email with processing stats and link

### Nightly Archive Flow (Secondary)
- **Trigger:** Scheduled (example: 10:00 PM)
- Archives **only successfully processed files**
- Leaves locked or failed files in the Source folder
- Archives the Daily Master file

---

## 📊 Scale and Constraints

- Up to **60 files per day**
- Up to **150 rows per file**
- Up to **~9,000 rows per run**
- Low concurrency to avoid Excel file locks
- Pagination enabled to prevent partial reads

---

## ✨ Features

- Scans SharePoint **Source** folder for `.xlsx` files
- Skips temporary lock files (`~$*.xlsx`)
- Reads **all tables** from each workbook
- Appends rows into a **single master table**
- Auto‑populates:
  - `SourceFile`
  - `ProcessedDateTime`
  - `ProcessingStatus`
- Explicit handling for locked and failed files
- Daily summary email
- Nightly archival with audit trail

---

## 📌 Processing Status Handling (Important)

Each file is evaluated and assigned a **processing status**:

| Status | Meaning | Behavior |
|------|--------|----------|
| `Processed` | File successfully read and merged | Archived at night |
| `Locked` | File was open during processing | Left in Source folder |
| `Failed` | File read error or schema issue | Left in Source folder |

**No file is archived unless it is successfully processed.**

Locked or failed files are:
- Explicitly reported in daily email
- Retried automatically in the next run

This guarantees **no data loss**.

---

## 📧 Daily Notification

After each daily run, an email is sent to authorized users containing:
- Number of files processed
- Number of rows merged
- List of locked or failed files
- Link to the Daily Master file

Locked files include owner details so corrective action can be taken.

---

## 📦 What’s in the Solution

- **Power Automate Cloud Flow**: `Merge Excel Files Data Into Master`
- **Environment Variables**
- **Connection References**
  - SharePoint
  - Excel Online (Business)

---

## 🧩 Prerequisites

- Power Platform environment (Dev / Prod)
- Permissions to import unmanaged solutions
- SharePoint library with:
  - **Source** folder for incoming files
  - **Archive** folder structure
  - Location for Daily Master files

> Each source workbook must contain at least one **Excel Table**  
> (Insert → Table). Ranges are ignored.

---

## 🔧 Environment Variables (Current Values to set)

To maintain **ALM (Application Lifecycle Management)** best practices, this solution uses Environment Variables. This allows you to move the solution from Dev to Test/Prod without editing the flow logic.

| Display Name | Name | Example Value | Notes |
| :--- | :--- | :--- | :--- |
| `EV_SrcSiteUrl` | `sp_EV_SrcSiteUrl` | `https://tenant.sharepoint.com/sites/FieldOps` | SharePoint site hosting the source inspection files. |
| `EV_SourceFolderID` | `sp_EV_SourceFolderID` | `b!abc123...` | **Internal Identifier** from *Peek code* (more reliable than plain paths). |
| `EV_MasterFilePath` | `sp_EV_MasterFilePath` | `/Shared Documents/Daily_Masters/Master.xlsx` | Server-relative path to the Daily Master workbook. |
| `EV_MasterTableName`| `sp_EV_MasterTableName`| `Table1` | The target table in the Master workbook. |
| `EV_ArchiveFolder`  | `sp_EV_ArchiveFolder`  | `/Shared Documents/Archive` | Path where processed files are moved nightly. |

---

## 🚀 Installation

1. **Download** the solution zip from the [Latest Release](https://github.com/spashikanti/SunilP-ExcelMergerToMaster/releases/latest).
2. Go to **Power Apps → Solutions → Import**.
3. **Upload** the ZIP file and complete the import wizard.
4. Set the **Current Values** for the Environment Variables listed above.
5. Open the flow and **Turn On**.

---

## 🧪 Testing

1. Upload 2–3 Excel files to the Source folder
2. Ensure each file contains a table with expected columns
3. Run the scheduled flow manually
4. Verify rows in the Daily Master table
5. Verify locked files remain in Source

---

## 🧠 Lessons Learned

- Align automation with **business cadence**, not file arrival
- Excel is often an **intake layer**, not the system of record
- Avoid endlessly growing master files
- Use **daily master files** for clarity and auditability
- Handle locked and late files explicitly
- Never archive unprocessed data
- Design emails for **operational clarity**, not noise
- Expect Excel Online connector limits
- Pagination and concurrency control are mandatory
- Late‑arriving data must be supported without data loss

---

## ❌ What Broke Initially (and How it was Fixed)

* **Excel Locking:** Files uploaded while still open caused 400/404 read failures due to system-level locks.
    * *Solution:* Implemented a filter to skip temporary lock files (`~$*.xlsx`) and added a retry policy on the Excel Online connector to handle transient busy states.
* **Defensive Pagination:** While current field templates are capped at ~150 rows, the default connector limit of 256 rows poses a risk of silent truncation if templates expand or files are merged.
    * *Solution:* Explicitly enabled **Pagination** with a 5000-row threshold to defensively ensure every record is captured regardless of future file growth.
* **Write Conflicts:** High concurrency in the "Apply to Each" loop led to intermittent "File in Use" errors on the Master workbook.
    * *Solution:* Set the **Concurrency Control** to `1` on the ingestion loop to ensure a single-threaded, stable, and sequential write process.
* **Path Unreliability:** Using plain text SharePoint paths caused intermittent 404 errors when moving the solution between environments (Dev/Test/Prod).
    * *Solution:* Switched to **Internal Folder Identifiers** (obtained via *Peek Code*) for 100% reliable file targeting across different SharePoint sites.
* **Duplicate Data on Retry:** Initial runs lacked a state check, so retrying a failed flow sometimes caused duplicate rows in the Master file.
    * *Solution:* Introduced a `ProcessingStatus` metadata check to ensure the flow only targets "Unprocessed" files, making the solution **idempotent**.

These hurdles were resolved through explicit architectural hardening rather than simple "out-of-the-box" configurations.

---

## 🔮 What I Would Improve Next

- Bulk insert using **Office Scripts**
- Optional ingestion into **Dataverse**
- Automated ERP handoff
- Validation and richer diagnostics
- Dynamic column mapping

---

## 🗺️ Roadmap

- [ ] Office Scripts bulk insert
- [ ] Enhanced diagnostics
- [ ] ERP integration option
- [ ] Video walkthrough
- [ ] Dynamic schema support

---

## 🤝 Contributing
Contributions are welcome!
- File **Issues** for bugs or feature requests.
- Submit **Pull Requests** for performance optimizations or new format support.

---

## 🤝 Community & Contribution
As a **Power Platform Super User**, I build these components to solve real-world architectural bottlenecks. Whether it's optimizing payload sizes or enhancing UI responsiveness, I'm always looking for better patterns. If you have ideas or optimizations, feel free to open a Pull Request!

# 👤 Author

<table style="border: none;">
  <tr>
    <td style="border: none;">
      <strong>Sunil Kumar Pashikanti</strong><br>
      <em>Principal Architect | Microsoft Power Platform Super User</em><br><br>

[![Community](https://img.shields.io/badge/Community-View%20Profile-indigo?style=for-the-badge&logo=microsoft)](https://community.powerplatform.com/profile/?userid=8077d18b-7b47-ee11-be6d-6045bdebe084)
&nbsp;
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-0A66C2?style=for-the-badge&logo=linkedin)](https://www.linkedin.com/in/sunil-kumar-pashikanti/)
&nbsp;
[![Blog](https://img.shields.io/badge/Blog-Blogger-FF5722?style=for-the-badge&logo=blogger)](http://sunilpashikanti.blogspot.com)
&nbsp;
[![Website](https://img.shields.io/badge/Portfolio-Visit-yellow?style=for-the-badge&logo=google-chrome)](https://sunilpashikanti.com)

</td>
  </tr>
</table>

**Support the Project:** If this solution helped you, please consider giving it a ⭐ to help others find it!



