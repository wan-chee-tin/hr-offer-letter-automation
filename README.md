# 🤖 HR Offer Letter Automation (UiPath)

## 📖 Project Overview

This repository contains a UiPath-based automation project designed to streamline the **HR offer letter process**. The bot handles validation, offer letter generation, and email dispatch for selected candidates using Excel and Word integration, along with Outlook 365. 

By implementing Dispatcher and Performer logic, this automation separates data processing from transactional execution — aligning with enterprise-level best practices. It also includes robust data validation and exception handling to ensure high reliability and accuracy.

This project serves as a personal portfolio example demonstrating proficiency in UiPath, REFramework, data validation, queue management, and document/email automation.

---

## 🎬 Demo Video

*Coming Soon...*  
📌 *(Optional: Add link to your YouTube or GitHub-hosted video when ready)*

---

## 🚀 How to Run

Before running the automation, ensure the following tools are installed:

1. **UiPath Studio** (Community or Enterprise): [Download here](https://www.uipath.com/)
2. **Microsoft Excel** (local install or web)
3. **Microsoft Word**
4. **Outlook 365 account** (for email integration via UiPath M365 activities)

### Steps:

#### 📦 Dispatcher

1. Open the project folder `HR_OfferLetter_Dispatcher` in UiPath Studio.
2. Configure the Excel file path in the workflow (e.g., `CandidateInfo.xlsx`).
3. Run the dispatcher. It will:
   - Read the Excel data
   - Filter for hired candidates
   - Validate required fields (e.g., Name, Email, Salary, etc.)
   - Mark invalid rows with `"Not Valid"` in a new `Validation` column
   - Add only valid rows to the Orchestrator queue

#### 🔁 Performer

1. Open the project folder `HR_OfferLetter_Performer` in UiPath Studio.
2. Ensure your Word offer letter template (`OfferLetterTemplate.docx`) is in the correct location.
3. Run the performer. It will:
   - Get queue items (hired, validated candidates)
   - Generate personalized DOCX and PDF offer letters
   - Send each offer letter via **Outlook 365**
   - Update transaction status in Orchestrator

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

---

## 🛠️ Technologies Used

- **UiPath Studio (with REFramework)**  
- **Microsoft Excel**  
- **Microsoft Word**  
- **Outlook 365 (via UiPath M365 Activities)**  
- **Orchestrator Queues (Dispatcher–Performer Architecture)**  

---
