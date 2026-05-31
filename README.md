# 🛡️ Endpoint DLP for Microsoft Purview

> A collection of PowerShell scripts, KQL queries, and reference samples to help administrators **manage, audit, and report** on Microsoft Purview **Endpoint Data Loss Prevention (DLP)** policies.

![PowerShell](https://img.shields.io/badge/PowerShell-100%25-5391FE?logo=powershell&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Microsoft%20Purview-0078D4?logo=microsoft&logoColor=white)
![Maintained](https://img.shields.io/badge/Maintained-yes-brightgreen)

---

## 📑 Table of Contents

- [Overview](#-overview)
- [Repository Contents](#-repository-contents)
- [Prerequisites](#-prerequisites)
- [Getting Started](#-getting-started)
- [Usage](#-usage)
  - [1. List DLP Policies](#1-list-dlp-policies)
  - [2. Get Details for a Specific Policy](#2-get-details-for-a-specific-policy)
  - [3. Export DLP Rules with Additional Info](#3-export-dlp-rules-with-additional-info)
  - [4. Export Activity Explorer Data](#4-export-activity-explorer-data)
  - [5. Advanced Hunting Query](#5-advanced-hunting-query)
- [Test DLP Websites](#-test-dlp-websites)
- [Contributing](#-contributing)
- [Author](#-author)

---

## 📖 Overview

**Microsoft Endpoint DLP** extends Data Loss Prevention monitoring and protection to activities on Windows and macOS endpoints. This repository provides ready-to-use **PowerShell cmdlet samples** and **KQL hunting queries** to:

- Inventory all DLP compliance policies and rules.
- Export detailed rule configurations (restrictions, scopes, alert settings) to CSV.
- Export Activity Explorer data for reporting and analysis.
- Visualize DLP alert trends in Microsoft Defender / Advanced Hunting.

---

## 📂 Repository Contents

| File | Description |
| --- | --- |
| `Get-DlpRuleWithAdditionalInfoV1.ps1` | Exports all Endpoint DLP rules with their restrictions (CloudEgress, RemovableMedia, Print, etc.), scopes, and metadata to a timestamped CSV. |
| `ExportActivity_PowershellV1.ps1` | Exports **Activity Explorer** data from Microsoft Purview to CSV or JSON, with paging and filtering support. |
| `Advance hunting query` | A KQL query for Microsoft Defender Advanced Hunting that charts DLP alert counts over time. |
| `img/` | Screenshots used in this documentation. |

---

## ✅ Prerequisites

- **PowerShell 5.1+** (Windows PowerShell) or **PowerShell 7+**.
- The **Exchange Online Management module**:
  ```powershell
  Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
  ```
- An account with appropriate **Microsoft Purview / Compliance** permissions (for example, Compliance Administrator or DLP Compliance Management).
- Connection to the **Security & Compliance PowerShell** endpoint:
  ```powershell
  Connect-IPPSSession -UserPrincipalName admin@yourtenant.onmicrosoft.com
  ```

---

## 🚀 Getting Started

```powershell
# 1. Clone the repository
git clone https://github.com/Rajeshcbsa/Endpoint-DLP-Microsoft.git
cd Endpoint-DLP-Microsoft

# 2. Connect to Security & Compliance PowerShell
Connect-IPPSSession -UserPrincipalName admin@yourtenant.onmicrosoft.com

# 3. Run any of the scripts below
```

---

## 🔧 Usage

### 1. List DLP Policies

Run the following command to get all DLP compliance policies:

```powershell
Get-DlpCompliancePolicy
```

**Output:**

![Get-DlpCompliancePolicy output](https://github.com/Rajeshcbsa/Endpoint-DLP-Microsoft/blob/main/img/output1.png?raw=true)

---

### 2. Get Details for a Specific Policy

```powershell
Get-DlpCompliancePolicy "MyDLP-Test01-Confidentialtext"
```

To show only Endpoint DLP locations with selected properties:

```powershell
Get-DlpCompliancePolicy "MyDLP-Test01-Confidentialtext" |
    Where-Object -Property EndpointDlpLocation -NE "" |
    Select-Object Priority, DisplayName, Mode, EndpointDlpLocation, EndpointDlpLocationException
```

**Output:**

![Get-DlpCompliancePolicy details output](https://github.com/Rajeshcbsa/Endpoint-DLP-Microsoft/blob/main/img/Output2.png?raw=true)

---

### 3. Export DLP Rules with Additional Info

`Get-DlpRuleWithAdditionalInfoV1.ps1` iterates through every Endpoint DLP policy and rule, flattens the endpoint restrictions (CloudEgress, RemovableMedia, NetworkShare, UnallowedApps, Print, RemoteDesktopServices, UnallowedBluetoothTransferApps, PasteToBrowser), and exports the result to a timestamped CSV in the current folder.

```powershell
.\Get-DlpRuleWithAdditionalInfoV1.ps1
```

> Output file: `Rulesstatus_<MM-dd-yyyy HH-mm-ss>.csv`

---

### 4. Export Activity Explorer Data

`ExportActivity_PowershellV1.ps1` exports Activity Explorer data to **CSV** or **JSON** with paging and up to five filters.

```powershell
.\ExportActivity_PowershellV1.ps1 `
    -StartTime "01/13/2024 00:00:01" `
    -EndTime   "01/14/2024 23:59:59" `
    -OutputFormat "csv" `
    -PageSize 50 `
    -Filter1 @("Activity", "DLPRuleMatch") `
    -Filter2 @("Workload", "Endpoint")
```

**Key parameters:**

| Parameter | Required | Description |
| --- | --- | --- |
| `-StartTime` | ✅ | Start of the time window. |
| `-EndTime` | ✅ | End of the time window. |
| `-OutputFormat` | ✅ | `csv` or `json`. |
| `-PageSize` | ❌ | Records per page (default `5000`). |
| `-RecordsCount` | ❌ | Total records to export (must be a multiple of `PageSize`). |
| `-Filter1` … `-Filter5` | ❌ | Activity Explorer filters as `@("Field", "Value")`. |
| `-UserPrincipalName` | ❌ | Connects an IPPS session before exporting. |
| `-Details` | ❌ | Writes a query-execution details file. |

> Output file: `Export-ActivityExplorerData <timestamp>.csv` / `.json`

---

### 5. Advanced Hunting Query

Use the following **KQL** query in **Microsoft Defender -> Advanced Hunting** to chart DLP alerts over time:

```kusto
AlertInfo
| join kind=inner (
    AlertEvidence
    | where DetectionSource contains "Microsoft Data Loss Prevention"
) on AlertId
| summarize count() by bin(Timestamp, 1d)
| render timechart
```

---

## 🌐 Test DLP Websites

Use these sites to validate Endpoint DLP upload-to-cloud or browser scenarios:

- Test DLP Site 1: https://testdlp.thehpc.in/
- Test DLP Site 2: https://testdlp.azurewebsites.net/

---

## 🤝 Contributing

Contributions are welcome. Feel free to open an issue or submit a pull request with improvements, new scripts, or bug fixes.

---

## 👤 Author

**Developed by [Rajeshcbsa](https://github.com/Rajeshcbsa)**  
Repository: https://github.com/Rajeshcbsa/Endpoint-DLP-Microsoft
