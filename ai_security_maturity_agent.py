"""
AI Security Maturity Assessment Agent
====================================

This script implements an automated agent for assessing an organization’s
AI security maturity against the Cloud Security Alliance (CSA) AI Security
Maturity Model (AISMM), the NIST AI Risk Management Framework (AI RMF)
and ISO/IEC 42001 requirements.  It is designed to work with the
comprehensive mapping workbook created as part of this project.

Overview
--------

The agent compares an organization’s list of implemented controls and
supporting evidence against the control objectives defined in the mapping
workbook.  For each AISMM control it determines:

* Whether the organization claims to have implemented the control.
* Whether evidence has been provided to substantiate the claim.
* Which frameworks (CSA, NIST and/or ISO) the control maps to.
* A compliance status (Compliant, Partial, Missing Evidence or
  Not Implemented).
* The maturity gap and a tailored recommendation based on the
  corresponding NIST playbook language and ISO/IEC 42001 checklist
  criteria.

The script produces a detailed per‑control assessment as well as high‑level
summaries to help risk, compliance and engineering teams identify gaps and
prioritize remediation activities.

Usage
-----

Run the agent from the command line, providing the path to the mapping
workbook and the path to your organization’s self‑assessment spreadsheet.

Example:

```
python ai_security_maturity_agent.py \
    --mapping-path CSA_AISMM_NIST_RMF_ISO42001_Maturity_Assessment_Best_Version.xlsx \
    --assessment-path my_org_controls.xlsx \
    --output-path my_org_assessment_results.xlsx
```

Input Files
-----------

**Mapping workbook** (required):  Must contain a sheet named
`Mapping` with the following columns.  Additional columns are
permitted and will be ignored if not referenced below.

* ``Control ID`` (string): unique identifier for the AISMM control.
* ``Control`` (string): summary of the control objective.
* ``NIST AI RMF Control ID`` (string): mapped NIST category or subcategory.
* ``NIST AI RMF Control Language / Playbook Actions`` (string):
  explanatory language from the NIST AI RMF playbook.
* ``ISO/IEC 42001 Clause`` (string): mapped ISO/IEC 42001 clause
  identifier or annex reference.
* ``ISO/IEC 42001 Control Language / Checklist Criteria`` (string):
  plain‑language summary of the requirement from the ISO/IEC 42001
  checklist.
* ``Frameworks Complied To`` (string): comma‑separated list of
  frameworks applicable to this control (e.g. ``CSA, NIST, ISO``).

**Assessment file** (required):  This is your organization’s
self‑assessment.  It must contain at least the following columns:

* ``Control ID`` (string): must match the AISMM ``Control ID`` in
  the mapping workbook.
* ``Implementation Status`` (string): one of ``Implemented``,
  ``Partial``, ``Not Implemented`` or any values you choose; values
  are compared case‑insensitively and the script infers compliance as
  described below.
* ``Evidence`` (string): description of the artefact proving that the
  control is implemented (e.g. policy document name, audit log
  reference, process URL).  Leave empty if no evidence is available.

Additional columns (such as ``Owner``, ``Last Assessed`` or
``Residual Risk``) are preserved in the output but not used by the
algorithm.

Output
------

The agent writes a new Excel workbook containing the following sheets:

* **Assessment_Results** – a row for each control in the mapping
  workbook with compliance status, gap description and
  recommendation.  All input columns from both files are retained
  where applicable.
* **Summary** – high‑level counts of controls by domain and
  compliance status, along with percentages.

If ``--output-path`` is not specified the results are saved as
``assessment_results.xlsx`` in the current working directory.

Implementation Notes
--------------------

* Compliance logic:

  - A control is marked **Compliant** if the organisation lists it as
    ``Implemented`` (case‑insensitive) *and* provides non‑empty
    evidence.
  - A control is **Partial** if it is listed as ``Partial`` *or*
    ``Implemented`` but no evidence is provided.
  - A control is **Missing Evidence** if it is listed as
    ``Implemented`` or ``Partial`` with an empty evidence field.
  - A control is **Not Implemented** if it is absent from the
    assessment file or explicitly marked ``Not Implemented``.

* Gap and recommendation:  When a control is not fully compliant, the
  script constructs a simple recommendation using the NIST playbook
  language and ISO/IEC 42001 checklist criteria.  Users may wish to
  tailor these recommendations for their context.

* Domain support:  If a column named ``AISMM Domain`` or
  ``Domain`` exists in the mapping sheet it is copied to the
  results and used to produce the summary.  Otherwise all controls
  are treated as belonging to a single domain.

Copyright and Licensing
-----------------------

This script is provided for educational purposes and to assist
organisations in operationalising the CSA AISMM.  It does not
constitute professional legal or compliance advice.  Consult the
official framework publications when implementing or certifying
against them.
"""

from __future__ import annotations

import argparse
import os
from collections import defaultdict
from typing import Dict, List, Optional

import pandas as pd


def load_mapping(mapping_path: str) -> pd.DataFrame:
    """Load the mapping workbook and return the mapping DataFrame.

    Parameters
    ----------
    mapping_path : str
        Path to the AISMM mapping workbook.  The workbook must contain
        a sheet named 'Mapping'.

    Returns
    -------
    pd.DataFrame
        DataFrame of the mapping sheet with string column names.
    """
    if not os.path.exists(mapping_path):
        raise FileNotFoundError(f"Mapping workbook not found: {mapping_path}")
    excel = pd.ExcelFile(mapping_path)
    if 'Mapping' not in excel.sheet_names:
        raise ValueError("Mapping workbook must contain a sheet named 'Mapping'.")
    df = excel.parse('Mapping')
    # Normalize column names by stripping whitespace
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_assessment(assessment_path: str) -> pd.DataFrame:
    """Load the organisation's self‑assessment file.

    Accepts both Excel (.xlsx, .xls) and CSV files.  The file must
    contain at least the columns 'Control ID', 'Implementation Status'
    and 'Evidence'.
    """
    if not os.path.exists(assessment_path):
        raise FileNotFoundError(f"Assessment file not found: {assessment_path}")
    _, ext = os.path.splitext(assessment_path)
    if ext.lower() in ('.xlsx', '.xls'):  # Excel file
        assessment_excel = pd.ExcelFile(assessment_path)
        # use the first sheet by default
        df_assess = assessment_excel.parse(assessment_excel.sheet_names[0])
    else:
        # assume CSV
        df_assess = pd.read_csv(assessment_path)
    df_assess.columns = [str(c).strip() for c in df_assess.columns]
    required = {'Control ID', 'Implementation Status', 'Evidence'}
    missing = required - set(df_assess.columns)
    if missing:
        raise ValueError(f"Assessment file is missing required columns: {missing}")
    return df_assess


def determine_compliance(status: Optional[str], evidence: Optional[str]) -> str:
    """Determine compliance status based on implementation status and evidence.

    Parameters
    ----------
    status : str
        The organisation's self‑reported implementation status.
    evidence : str
        Evidence string provided by the organisation (may be empty).

    Returns
    -------
    str
        One of 'Compliant', 'Partial', 'Missing Evidence' or 'Not Implemented'.
    """
    status_normalized = (status or '').strip().lower()
    evidence_present = bool((evidence or '').strip())
    if status_normalized in ('implemented', 'complete', 'compliant'):
        if evidence_present:
            return 'Compliant'
        return 'Missing Evidence'
    if status_normalized in ('partial', 'partially implemented', 'in progress'):
        # partial implementation always requires evidence to be considered
        return 'Partial'
    if status_normalized == 'not implemented':
        return 'Not Implemented'
    # If status is missing or unrecognised and evidence is present, treat as partial
    if evidence_present:
        return 'Partial'
    return 'Not Implemented'


def build_recommendation(row: pd.Series, compliance: str) -> str:
    """Construct a recommendation for a given mapping row and compliance result.

    This function uses the NIST and ISO control language columns to
    formulate a simple recommendation.  Users are encouraged to refine
    the recommendations based on organisational context.
    """
    # Fetch guidance from NIST and ISO columns
    nist_lang = str(row.get('NIST AI RMF Control Language / Playbook Actions', '')).strip()
    iso_lang = str(row.get('ISO/IEC 42001 Control Language / Checklist Criteria', '')).strip()
    control_id = str(row.get('Control ID', '')).strip()
    control_name = str(row.get('Control', '')).strip()
    # choose the more descriptive language for recommendation
    guidance = nist_lang or iso_lang or control_name
    if compliance == 'Compliant':
        return f"No action required; continue maintaining evidence for {control_id}."
    # For missing evidence, emphasise documentation and proof
    if compliance == 'Missing Evidence':
        return f"Provide documented evidence demonstrating that the organisation meets the requirements of control {control_id}: {guidance}."
    # For partial implementation, emphasise completing implementation
    if compliance == 'Partial':
        return f"Complete implementation of control {control_id} and retain evidence that the control objective is met. Guidance: {guidance}."
    # Not implemented
    return f"Implement control {control_id} in accordance with the following guidance: {guidance}."


def assess_controls(mapping_df: pd.DataFrame, assessment_df: pd.DataFrame) -> pd.DataFrame:
    """Perform a row‑level assessment of controls.

    Parameters
    ----------
    mapping_df : pd.DataFrame
        DataFrame containing the AISMM mapping information.
    assessment_df : pd.DataFrame
        DataFrame containing the organisation’s self‑assessment.

    Returns
    -------
    pd.DataFrame
        DataFrame containing the original mapping columns plus
        compliance results, gaps and recommendations, and any
        assessment columns.
    """
    # Create lookup for assessment entries keyed by control ID
    assess_lookup: Dict[str, pd.Series] = {}
    # Remove duplicate control IDs by taking the first occurrence
    for _, arow in assessment_df.iterrows():
        cid = str(arow['Control ID']).strip()
        if cid not in assess_lookup:
            assess_lookup[cid] = arow

    results: List[pd.Series] = []
    for _, mrow in mapping_df.iterrows():
        cid = str(mrow.get('Control ID', '')).strip()
        assess_row = assess_lookup.get(cid)
        status: Optional[str] = None
        evidence: Optional[str] = None
        if assess_row is not None:
            status = str(assess_row.get('Implementation Status', '')).strip()
            evidence = str(assess_row.get('Evidence', '')).strip()
        compliance = determine_compliance(status, evidence)
        # Determine gap description
        if compliance == 'Compliant':
            gap = 'No gaps'
        elif compliance == 'Partial':
            gap = 'Control partially implemented'
        elif compliance == 'Missing Evidence':
            gap = 'Evidence missing'
        else:
            gap = 'Control not implemented'
        recommendation = build_recommendation(mrow, compliance)
        # Combine mapping row, assessment row and new fields
        combined = mrow.copy()
        # Append assessment fields if they exist
        if assess_row is not None:
            for col in assessment_df.columns:
                # Avoid overwriting mapping columns
                if col not in combined:
                    combined[col] = assess_row[col]
        combined['Compliance Status'] = compliance
        combined['Gap'] = gap
        combined['Recommendation'] = recommendation
        results.append(combined)
    results_df = pd.DataFrame(results)
    return results_df


def build_summary(results_df: pd.DataFrame) -> pd.DataFrame:
    """Build a summary table of compliance by domain.

    The function expects a column named either 'AISMM Domain' or
    'Domain' in the results.  If neither is present, all controls are
    grouped under a single domain labelled 'All'.
    """
    domain_col = None
    for col in results_df.columns:
        col_lower = col.strip().lower()
        if col_lower == 'aismm domain' or col_lower == 'domain':
            domain_col = col
            break
    if domain_col is None:
        results_df['_Domain'] = 'All'
        domain_col = '_Domain'
    summary_records = []
    for domain, group in results_df.groupby(domain_col):
        counts = group['Compliance Status'].value_counts().to_dict()
        total = len(group)
        summary_record = {
            'Domain': domain,
            'Total Controls': total,
            'Compliant': counts.get('Compliant', 0),
            'Partial': counts.get('Partial', 0),
            'Missing Evidence': counts.get('Missing Evidence', 0),
            'Not Implemented': counts.get('Not Implemented', 0)
        }
        # compute percentages
        for key in ['Compliant', 'Partial', 'Missing Evidence', 'Not Implemented']:
            summary_record[f'{key} %'] = (
                summary_record[key] / total * 100 if total > 0 else 0
            )
        summary_records.append(summary_record)
    return pd.DataFrame(summary_records)


def write_results(results_df: pd.DataFrame, summary_df: pd.DataFrame, output_path: str) -> None:
    """Write the assessment results and summary to an Excel file."""
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Save results sheet
        results_df.to_excel(writer, sheet_name='Assessment_Results', index=False)
        # Save summary
        summary_df.to_excel(writer, sheet_name='Summary', index=False)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Assess an organisation's AI security controls against the CSA"
            " AI Security Maturity Model (AISMM), NIST AI RMF and ISO/IEC 42001."
        )
    )
    parser.add_argument('--mapping-path', required=True,
                        help='Path to the AISMM mapping workbook.')
    parser.add_argument('--assessment-path', required=True,
                        help="Path to your organisation's self‑assessment file.")
    parser.add_argument('--output-path', default='assessment_results.xlsx',
                        help='Where to write the assessment results.')
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    mapping_df = load_mapping(args.mapping_path)
    assessment_df = load_assessment(args.assessment_path)
    results_df = assess_controls(mapping_df, assessment_df)
    summary_df = build_summary(results_df)
    write_results(results_df, summary_df, args.output_path)
    print(f"Assessment completed. Results saved to {args.output_path}")


if __name__ == '__main__':
    main()