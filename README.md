## 📖 Project Overview

This repository contains a UiPath-based automation project designed to streamline the HR offer letter process. The bot handles validation, offer letter generation, and email dispatch for selected candidates using Excel and Word integration, along with Outlook 365. By implementing Dispatcher and Performer logic, this automation separates data processing from transactional execution — aligning with enterprise-level best practices. It also includes robust data validation and exception handling to ensure high reliability and accuracy.

---

## 🎬 Demo Video

*Coming Soon...*

---

## 🚀 How to Run

Before running this automation, ensure the following are configured:

### ⚙️ Required Tools
- [UiPath Studio](https://www.uipath.com/)
- Microsoft Excel & Microsoft Word
- UiPath Orchestrator (for Assets and Queues)
- Outlook 365 (email account integration via UiPath M365 activities)

### 🔹 Dispatcher: `Dispatcher_OfferLetterAutomation/Main.xaml`

This workflow reads the candidate data, performs validation, and queues valid records.

#### ✅ Orchestrator Setup

**Assets:**

| Asset Name                 | Type | Example Value                                     |
|---------------------------|------|--------------------------------------------------|
| `ExcelPath_CandidateInfo` | Text | `C:\Users\UiPath\Documents\CandidateInfo.xlsx`     |
| `SheetName_CandidateInfo` | Text | `Status`                                         |

**Queue:**

| Queue Name               |
|--------------------------|
| `Queue_OfferLetter_Hired` |

#### ▶ Steps to Run Dispatcher

1. Ensure the assets and queue above are created in Orchestrator.
2. Open the Dispatcher project in UiPath Studio:  
   `Dispatcher_OfferLetterAutomation/Main.xaml`
3. Run or publish the Dispatcher process.
4. The bot will:
   - Read candidate records from Excel (via asset)
   - Filter for `Status = "Hired"`
   - Validate required fields and email format
   - Mark invalid records with `"Not Valid"` in the `Validation` column
   - Add valid candidates to the `Queue_OfferLetter_Hired` queue

### 🔸 Performer: `Performer_OfferLetterAutomation/Main.xaml`

This workflow processes the queued candidates and sends personalized offer letters.

#### ✅ Configuration Required

Update the `Config.xlsx` file inside the **`Data/`** folder with the following keys:

| Key Name                   | Description                             |
|---------------------------|-----------------------------------------|
| `EmployerName`            | Company or employer name                |
| `OfferLetterTemplateFilePath` | Full path to Word template file      |
| `OfferLetterFolderPath`   | Folder where generated letters will be saved |
| `EmailSubject`            | Subject line for offer email            |
| `EmailBody`               | Body message for offer email            |
| `HRName`, `HREmail`, etc. | HR contact info used in the offer letter/email |

#### ▶ Steps to Run Performer

1. Open the Performer project in UiPath Studio:  
   `Performer_OfferLetterAutomation/Main.xaml`
2. Ensure `Config.xlsx` is updated correctly.
3. Run or publish the Performer process.
4. The bot will:
   - Retrieve candidate data from `Queue_OfferLetter_Hired`
   - Fill in the offer letter template with candidate info
   - Generate and save DOCX and PDF versions
   - Send offer letter via Outlook 365
   - Mark transaction success/failure in Orchestrator

---

## 🔄 Process Flow

### Dispatcher
1. Read candidate data from Excel
2. Filter for `Status = "Hired"`
3. Validate required fields and email format
4. Write `"Not Valid"` in `Validation` column for invalid records
5. Add valid candidates to Orchestrator queue

### Performer
1. Get each transaction item from queue
2. Open Word template and fill in candidate info
3. Save personalized offer letter as DOCX and PDF
4. Send the offer letter via Outlook 365
5. Log and mark transaction as successful or failed
