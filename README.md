# Insurance Grievance Automation & Root Cause Analysis (RCA) Engine

An automated operations tool designed to streamline regulatory grievance logging, execute automated Root Cause Analysis (RCA), and instantly generate executive-ready compliance presentations.

## 📌 Project Overview
This repository hosts a production-ready Python solution that standardizes how customer grievances are audited, categorized, and presented. By translating raw service delivery breakdowns into data-driven analytical structures (using the 5-Why Framework), this tool helps insurance operations teams maintain compliance and meet tight regulatory Turnaround Times (TAT).

### Key Features
* **Automated PowerPoint Generation:** Generates a structured corporate 16:9 presentation slide deck matching institutional standards.
* **5-Why RCA Framework:** Programmatically structures the narrative from operational breakdown to systemic root cause.
* **Audit Registry Export:** Standardizes inputs into a flat grid format easily mapable to corporate governance logs.

---

## 📊 Core Case Analytics Map

| Parameter | Operational Registry Data |
| :--- | :--- |
| **Systemic Category** | Operational Failure / Intermediary Misconduct |
| **Monitored Vendor** | Shriram Life Insurance |
| **Regulatory TAT Window** | 14 Days Maximum Limit |
| **Core Delivery Deficit** | 60-Day physical policy print and dispatch delay |
| **Intermediary Issue** | Unresponsive field agent tracking / Communication block |

---

## 🧠 Systemic Root Cause Analysis (RCA Summary)
The engine traces operational failures using an structured investigative sequence:
1. **The Trigger:** Customer filed a formal grievance due to an unfulfilled physical delivery promise and absolute field agent non-responsiveness.
2. **The Field Deficit:** Absence of localized field-accountability mechanisms and intermediary code-of-conduct enforcement at the branch level.
3. **The Logistical Gap:** Manual hand-delivery promises by agents allowed the order to completely bypass standard automated central logistics tracking loops.
4. **The Monitoring Blindspot:** Lack of a centralized systemic data reconciliation tool matching newly issued active policies against physical courier tracking IDs.
5. **Root Cause:** Operational dependency on manual field-agent distribution paths without an automated, system-driven exception alert for aging unfulfilled dispatches.

---

## 🛡️ Corrective & Preventive Action (CAPA) Log

### Corrective Actions (Immediate Remediation)
* Intervened via central insurer SPOC to track, procure, and deliver the physical policy document.
* Secured direct policyholder sign-off and successfully closed the case file within the 14-day regulatory TAT limit.

### Preventive Actions (Long-Term Mitigations)
* **Intermediary Accountability:** Logged a formal misconduct infraction against the associated agent for internal disciplinary review.
* **Systemic Exception Filters:** Implemented a 20-day central systemic ageing trigger alert to flag any issued policy lacking a verified courier tracking identifier.
* **Direct Logistics:** Transitioned dispatch status notifications directly to customer endpoints (SMS/Email automated hooks), removing manual field agent dependency.

---

## 🚀 How to Execute via the Cloud
Since this tool is designed to work within managed network environments, it can be run completely in the cloud with zero local software installation.

1. Copy the source script from `generate_rca_report.py`.
2. Open a free, browser-based cloud environment like **Google Colab**.
3. Install the layout dependencies inside a code cell:
   ```bash
   !pip install python-pptx
   ```
4. Paste the script execution block and hit **Run**.
5. Download your generated `Grievance_RCA_Report.pptx` directly through the cloud browser platform.


