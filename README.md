# SunilP-ExcelMergerToMaster
Automated workflow to merge data from multiple Excel workbooks into a single Master Table.

![Microsoft Community](https://img.shields.io/badge/Microsoft%20Community-Super%20User-orange?style=for-the-badge&logo=microsoft)
![License](https://img.shields.io/badge/License-MIT-yellow?style=for-the-badge)
![Power Automate](https://img.shields.io/badge/Platform-Power%20Automate-blue?style=for-the-badge&logo=power-automate)
![SharePoint](https://img.shields.io/badge/Storage-SharePoint-036C71?style=for-the-badge&logo=microsoft-sharepoint)

---

### 💡 The Solution
This solution merges data from **multiple Excel workbooks** stored in a SharePoint folder into a **single Master Excel Table**.

- Reads **all tables** found in each source workbook (loops over tables per file).
- Appends rows into **one fixed table** in the master workbook (e.g., `Table1`).
- Delivered as a **Power Platform Solution** and configurable via **Environment Variables (EVs)**.

> **Note:** This package implements a **row‑by‑row** append using the Excel Online (Business) connector. A bulk‑insert variant using Office Scripts can be added later.

---

## ⬇️ Get Started

**Download the latest Power Platform solution package (ZIP) to automate your Excel consolidation:**

[![Download ZIP](https://img.shields.io/badge/Download-Solution-blue?style=for-the-badge&logo=github)](https://github.com/spashikanti/SunilP-ExcelMergerToMaster/releases/latest)
[![Release](https://img.shields.io/github/v/release/spashikanti/SunilP-ExcelMergerToMaster?style=for-the-badge&logo=github&color=brightgreen)](https://github.com/spashikanti/SunilP-ExcelMergerToMaster/releases/latest)
[![Total Downloads](https://img.shields.io/github/downloads/spashikanti/SunilP-ExcelMergerToMaster/total?style=for-the-badge&color=yellow)](https://github.com/spashikanti/SunilP-ExcelMergerToMaster/releases)

---

## ✨ Features

- Scans a SharePoint **Source** folder and finds all `.xlsx` files  
- Skips lock files (e.g., `~$...xlsx`)  
- For each workbook: loops **all defined tables** and reads their rows  
- Appends all rows into the **master workbook’s single table** (`Table1`)  
- Pagination enabled to handle larger tables  
- Concurrency tuned to avoid Excel file locks

---

## 📦 What’s in the Solution

- **Power Automate Cloud Flow**: `Merge Excel Files Data Into Master`
- **Environment Variables** (EVs)
- **Connection References**: SharePoint, Excel Online (Business)

---

## 🧩 Prerequisites

- Access to a Power Platform environment (Developer/Production) to import solutions
- Permissions to import unmanaged solutions
- SharePoint library with:
  - A **Source** folder containing the input `.xlsx` workbooks  
  - A **Master** workbook that has a defined table (e.g., `Table1`)

> Each source workbook must contain at least one **Table** (Insert → Table). Ranges without tables are ignored.

---

## 🚀 Installation

1. **Download** the solution zip from this repo, for example: [ExcelMergerToMaster....zip](https://github.com/spashikanti/SunilP-ExcelMergerToMaster/releases/latest)
2. Go to:
**Power Apps → Solutions → Import**

3. **Upload** the ZIP file and complete the import wizard.

4. After import, go to:  
**Solutions → ExcelMergeToMaster → Environment Variables**

5. Set **Current Values** for each EV:

### 🔧 Environment Variables (Current Values to set)

| Display name              | Name                      | Example Current Value                                           | Notes                                                                                  |
|---------------------------|---------------------------|------------------------------------------------------------------|----------------------------------------------------------------------------------------|
| `EV_DestinationDocLib`    | `sp_EV_DestinationDocLib` | `Documents`                                                      | Internal library name for the **destination** (master) site (often `Documents`).       |
| `EV_DestinationSiteUrl`   | `sp_EV_DestinationSiteUrl`| `https://<tenant>.sharepoint.com/sites/site`                      | **No trailing slash**. SharePoint site that hosts the **master** workbook.             |
| `EV_MasterFilePath`       | `sp_EV_MasterFilePath`    | `/Shared Documents/MasterFiles/Master.xlsx`                      | Server‑relative path to the **master** workbook.                                       |
| `EV_MasterTableName`      | `sp_EV_MasterTableName`   | `Table1`                                                         | The **single** target table in the master workbook.                                    |
| `EV_SourceFolder`         | `sp_EV_SourceFolder`      | `/drives/b!abc123/root:/Shared Documents/SourceFiles`            | **Recommended:** use the **internal folder identifier** from *Peek code* (not plain path). |
| `EV_SrcDocLib`            | `sp_EV_SrcDocLib`         | `Documents`                                                      | Internal library name for the **source** site (often `Documents`).                     |
| `EV_SrcSiteUrl`           | `sp_EV_SrcSiteUrl`        | `https://<tenant>.sharepoint.com/sites/site`                      | **No trailing slash**. SharePoint site that hosts the **source** files.                |


6. Open the flow → **Turn On**.

7. Test by running manually or by waiting for the schedule (if configured).

---

## 🧪 Testing

1. Upload **2–3** `.xlsx` files into the **Source** folder.  
   - Each file must contain **at least one** Table.  
   - All tables should share the same **column structure** expected by the master table.
  
    ### 📑 Columns used in this solution 4 + 1):**
      `Date`, `Merchant`, `Amount`, `Category`, **`SourceFile`**
    
    > The flow **automatically populates** `SourceFile` with the originating file name while appending rows to the master table (`Table1`).  
    > **Note:** Current version maps these four columns explicitly. A future enhancement will make column mapping dynamic or support whole‑row insertion.
    ``
    
    This workflow **reads all tables** from each source workbook and **writes into one fixed table** in the master workbook.
    
    - **Source workbooks (each table) — expected columns (4):**  
      `Date`, `Merchant`, `Amount`, `Category`
    
2. Run the flow manually.

3. Open the master workbook and verify rows were appended to **`Table1`**.

---

## 🔧 How It Works (High Level)

1. **List folder** → gets files from the Source folder (filters to `.xlsx`, skips `~$` lock files)  
2. **For each file** → **Get tables**, then loop **all tables** found in that workbook  
3. **For each table** → **List rows present in a table** (pagination enabled)  
4. **For each row** → **Add a row into a table** (master workbook, fixed `Table1`)  

> The design ensures **all tables** in each source workbook are read, but **only one master table** receives the data.

---

## 🧰 Troubleshooting

**❗ 404 – “The response is not in a JSON format” on List folder**  
- The SharePoint connector needs the **internal folder identifier** (from **… → Peek code**) rather than a plain text folder path.  
  Add that internal ID to `EV_SourceFolderIdentifier`.

**❗ Excel actions stuck on “loading…”**  
- Repair or rebind the **Excel Online (Business)** connection reference inside the **Solution**.

**❗ Missing rows**  
- Ensure **Pagination** is enabled (e.g., threshold **5000**) in “List rows present in a table”.

**❗ Table not found**  
- Confirm each source workbook uses **Insert → Table** (named or default).  
- Confirm the master workbook contains the target table (e.g., `Table1`).

**❗ Locking / conflicts**  
- Keep concurrency at **1** for the loop that writes to the master workbook.  
- Avoid opening the master workbook while the flow runs.

---

## 🗺️ Roadmap

- [ ] Publish **bulk‑insert** variant using **Office Scripts** (single‑call append)  
- [ ] Add optional **GitHub Actions** (ALM) for export/unpack and managed imports  
- [ ] Enhanced diagnostics & logging  
- [ ] Step‑by‑step video tutorial
- [ ] Support dynamic columns

---

## 🤝 Contributing

Contributions are welcome!  
- File issues for bugs and enhancements  
- Submit PRs for improvements (bulk insert, error handling, docs)

---

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

