AI Security Maturity Assessment Agent


This document explains how to use the AI Security Maturity Assessment agent built for evaluating organisational controls against the CSA AI Security Maturity Model (AISMM), the NIST AI Risk Management Framework (AI RMF), and ISO/IEC 42001.


Purpose
The assessment agent is designed to help organisations evaluate their AI security controls by comparing them against a comprehensive mapping of AISMM controls to NIST AI RMF categories and ISO/IEC 42001 clauses.
It processes your organisation's self‑assessment, identifies which controls are fully implemented, partially implemented, missing evidence or not implemented, and generates targeted recommendations using guidance from the NIST AI RMF and ISO/IEC 42001.

Files Required
You will need:
•	• **Mapping workbook** – the cross‑walk spreadsheet that maps AISMM controls to NIST AI RMF and ISO/IEC 42001. Ensure it contains a sheet named “Mapping” with columns such as Control ID, Control, NIST AI RMF Control Language / Playbook Actions, ISO/IEC 42001 Control Language / Checklist Criteria and Frameworks Complied To.
•	• **Self‑assessment file** – your organisation’s controls and evidence. At a minimum this file should include the columns Control ID, Implementation Status and Evidence. You may include additional columns such as Owner, Business Unit, Target Maturity Score or Residual Risk.


Running the Agent
The agent is implemented in the Python script `ai_security_maturity_agent.py`. To run the assessment:
1.	1. Install the Python dependencies (pandas and openpyxl are required). These packages are available in most Python environments.
2. Run the script from the command line, specifying the mapping file and your assessment file:
2.	   python ai_security_maturity_agent.py --mapping-path <mapping.xlsx> --assessment-path <assessment.xlsx> --output-path <results.xlsx>
3.	3. After execution, the script will create an Excel workbook with two sheets:
   • **Assessment_Results** – detailed per‑control analysis including the compliance status, gaps and recommended next steps.
   • **Summary** – aggregated counts and percentages of compliant, partial, missing evidence and not implemented controls per AISMM domain.


Understanding Compliance Status
The agent determines compliance using simple, transparent rules:
•	• **Compliant** – the control is reported as “Implemented” (or “Complete”) and valid evidence is provided.
•	• **Partial** – the control is reported as “Partial” or has evidence but no clear implementation status. Additional work is needed.
•	• **Missing Evidence** – the control is reported as “Implemented” or “Partial” but no evidence has been provided. Documentation must be gathered.
•	• **Not Implemented** – the control is absent from your assessment or explicitly marked “Not Implemented”.


Recommendations
For each control that is not fully compliant, the agent generates a concise recommendation leveraging the NIST AI RMF Playbook language or ISO/IEC 42001 checklist criteria. These recommendations are meant to guide you toward concrete actions, such as implementing a specific process, documenting policies, or collecting evidence. Tailor them to your organisational context for maximum effectiveness.


Extending the Agent
The assessment agent is open for enhancement. Organisations may wish to:
•	• **Incorporate weighted maturity scoring** – adjust the logic to calculate maturity scores based on evidence quality or control criticality.
•	• **Integrate with GRC systems** – export the results to your governance, risk and compliance platform for continuous monitoring.
•	• **Customise recommendations** – tailor the recommendation templates to match your organisation’s policies, terminology and risk appetite.


Disclaimer
This tool is provided to assist with internal assessments and does not constitute legal or compliance advice. Refer to the official CSA AISMM, NIST AI RMF and ISO/IEC 42001 publications for authoritative guidance.
