# Career_Suite_AI_v4.5

# 🤖 Master Job Manager: An AI-Powered Job Application CRM in Google Sheets **REVISED NAME**

# Case Study: A Modular, Serverless AI Workflow Engine

This repository contains the source code for **CareerSuite.ai v4.5**, a project re-architected to demonstrate a modern, event-driven, and human-in-the-loop AI processing system. The application itself tailors resumes, but its primary purpose is to serve as a portfolio piece showcasing specific skills in software architecture, API orchestration, and user-centric system design.

The entire system is implemented using **Google Apps Script**, running serverlessly within a user's own Google Workspace ecosystem.

---

## 🚀 Architectural Highlights & Skills Demonstrated

This project was engineered to solve specific technical challenges and showcase the following capabilities:

*   **Decoupled, Staged Architecture:** The core of the project is the refactoring of a monolithic script into a **decoupled, 3-stage execution model** (`Stage 1: Analyze`, `Stage 2: Process`, `Stage 3: Assemble`). This design separates concerns, prevents cascading failures, and allows for asynchronous-like user interaction in a synchronous environment.

*   **Human-in-the-Loop (HITL) System Design:** A key design principle was to avoid a "black box" AI. The system externalizes its state and the results of its AI analysis to a **Google Sheet, which acts as a staging database and UI**. This allows a human user to review, validate, and selectively approve machine-generated results before triggering the next stage of processing—a critical pattern for building trustworthy AI systems.

*   **Multi-API Orchestration & Latency Optimization:** The system is built to orchestrate calls between multiple services (Google Workspace APIs) and third-party LLM providers. It specifically integrates **Groq's LPU™ Inference Engine** as the primary LLM to drastically reduce latency in the analysis phase, creating a near-instantaneous feedback loop for the user. It retains flexibility by also supporting Google's Gemini models, demonstrating the ability to abstract and switch between external services.

*   **Serverless, "Full-Stack" Implementation:** The entire application—logic, data storage, and document generation—operates within the Google Workspace serverless environment. This showcases the ability to build and deploy a complete, functional application with zero infrastructure management, using the available platform services as a cohesive stack (Apps Script as backend logic, Sheets as a database/UI, Docs as a templating engine).

*   **Configuration-Driven & Maintainable Code:** Readability and maintainability are emphasized through the use of a centralized `Constants.gs` configuration file and comprehensive JSDoc comments, demonstrating professional coding standards.

---

## 🛠️ System Design: The 3-Stage Workflow

The application operates as a state machine where the state is managed in a Google Sheet. The user initiates each stage transition.

**Diagram:**
`[User Input: Job Desc] -> [Stage 1: Analyze & Score] -> [Sheet: BulletScoringResults] -> [User Action: 'YES'] -> [Stage 2: Tailor] -> [Sheet: TailoredText] -> [User Action: Run] -> [Stage 3: Assemble & Generate] -> [Output: Google Doc]`

**Stage 1: Analyze & Score (`runStage1_AnalyzeAndScore`)**
*   **Trigger:** User provides a job description and executes the function.
*   **Process:**
    1.  Fetches master data from the `ResumeData` sheet.
    2.  Sends the job description to the configured LLM (Groq default) for structured analysis.
    3.  Iterates through every bullet point in the master resume.
    4.  Makes individual LLM calls to score each bullet's relevance against the analyzed job description.
*   **Output:** Writes all scores, AI justifications, and other metadata to the `BulletScoringResults` sheet. The system is now "paused," awaiting human input.

**Stage 2: Tailor Selected (`runStage2_TailorSelectedBullets`)**
*   **Trigger:** User reviews the `BulletScoringResults` sheet, marks desired rows with "YES", and executes the function.
*   **Process:**
    1.  Reads the entire `BulletScoringResults` sheet.
    2.  Filters for rows where the "SelectToTailor" column is marked "YES".
    3.  For each selected row, sends the original bullet point to the LLM with a prompt to rephrase and tailor it.
*   **Output:** Updates the `TailoredBulletText(Stage2)` column in-place for the corresponding rows.

**Stage 3: Assemble & Generate (`runStage3_BuildAndGenerateDocument`)**
*   **Trigger:** User executes the final stage function.
*   **Process:**
    1.  Constructs a final resume object in memory.
    2.  It uses the `BulletScoringResults` sheet as the "source of truth," selecting only user-approved bullets. It prioritizes tailored text from Stage 2 if available, otherwise falling back to the original.
    3.  Gathers the highest-scoring content to generate a new, tailored professional summary via an LLM call.
    4.  Merges the final object with a Google Doc template.
*   **Output:** A new, formatted Google Doc is created in the user's Drive.

---

### ⚙️ Core Technologies

*   **Backend & Orchestration:** Google Apps Script (JavaScript/ES5)
*   **AI Inference:** Groq API (Primary), Google Gemini API (Secondary)
*   **Data Persistence & State Management:** Google Sheets
*   **Templating & Output:** Google Docs API
*   **Development Principles:** Decoupled Architecture, Human-in-the-Loop, Serverless Computing, Configuration-as-Code.

---

### 📂 Repository & Setup

*(This section can be the same as the previous one, as it's the practical guide)*

1.  **Get Sheet & Doc Templates:**
    *   [Master Resume Sheet Template](https://docs.google.com/spreadsheets/d/1FmsiLee476IwW4atDs3I_E017MX4LFkGNvyoo1U7D0w/copy)
    *   [Resume Doc Output Template](https://docs.google.com/document/d/18eX765FWVBHpOZ2jzxwNdGVKSMOzmyPJmS_Kfzp8JSk/copy)
2.  **Setup the Script:**
    *   In your new Google Sheet, go to `Extensions > Apps Script`.
    *   Copy the code from the `.gs` files in this repository into corresponding files in the editor.
    *   In `Constants.gs`, update the `MASTER_RESUME_SPREADSHEET_ID` and `RESUME_TEMPLATE_DOC_ID` with the IDs from your copies.
3.  **Set API Keys:**
    *   Go to `Project Settings > Script Properties`.
    *   Add `GROQ_API_KEY` and `GEMINI_API_KEY` with your respective keys.
4.  **Execute:** Run the `runStage*` functions in order from the Apps Script editor.

---

### 🤝 Contact

I'm always open to connecting with other developers and technical leaders. Let's talk about system design, AI integration, or interesting engineering challenges.

*   **[Your LinkedIn Profile URL]**
*   **[Your Personal Portfolio Website URL]**
