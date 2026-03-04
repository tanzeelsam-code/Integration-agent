"""
AGENT ZEE — Jobs Registry
Central registry for all available processing jobs.
"""
from __future__ import annotations

from .proposal_development import process_proposal
from .document_analysis import process_analysis
from .comparison import process_comparison
from .project_management import process_project_management
from .report_writing import process_report
from .jis_mapping import process_jis_mapping
from .cv_rewriting import process_cv_rewrite
from .contract_management import process_contract
from input_adapter import supported_extensions_csv


ACCEPTED_UPLOADS = supported_extensions_csv()


JOB_REGISTRY = {
    "formatting": {
        "name": "Document Formatting",
        "description": "Reformat uploaded files to AGENT ZEE style guide with full compliance",
        "icon": "format",
        "accept": ACCEPTED_UPLOADS,
        "multi_file": False,
        "fields": [
            {"id": "report_name", "label": "Report Name", "default": "Report"},
            {"id": "year", "label": "Year", "default": "2026"},
            {"id": "project_number", "label": "Project Number", "default": "PRJ-001"},
        ],
        "processor": None,  # Uses the original formatter engine
    },
    "proposal": {
        "name": "Proposal Development",
        "description": "Generate structured proposals compliant with WB, EU, and ADB requirements",
        "icon": "proposal",
        "accept": ACCEPTED_UPLOADS,
        "multi_file": False,
        "fields": [
            {"id": "client_type", "label": "Client Type", "type": "select",
             "options": ["World Bank (WB)", "European Union (EU)", "Asian Development Bank (ADB)", "Other"],
             "default": "World Bank (WB)"},
            {"id": "project_title", "label": "Project Title", "default": ""},
            {"id": "country", "label": "Country / Region", "default": "Pakistan"},
        ],
        "processor": process_proposal,
    },
    "analysis": {
        "name": "Documentation Analysis",
        "description": "Comprehensive document analysis with readability metrics and quality assessment",
        "icon": "analysis",
        "accept": ACCEPTED_UPLOADS,
        "multi_file": False,
        "fields": [
            {"id": "analysis_depth", "label": "Analysis Depth", "type": "select",
             "options": ["Quick Overview", "Standard", "Deep Dive"],
             "default": "Standard"},
        ],
        "processor": process_analysis,
    },
    "comparison": {
        "name": "Document Comparison",
        "description": "Compare two documents side-by-side with structural and content gap analysis",
        "icon": "comparison",
        "accept": ACCEPTED_UPLOADS,
        "multi_file": True,
        "fields": [
            {"id": "comparison_mode", "label": "Comparison Mode", "type": "select",
             "options": ["Full Comparison", "Structure Only", "Content Only"],
             "default": "Full Comparison"},
        ],
        "processor": process_comparison,
    },
    "project_mgmt": {
        "name": "Project Management",
        "description": "Generate WBS, Gantt tables, RACI matrix, and risk register from project docs",
        "icon": "project",
        "accept": ACCEPTED_UPLOADS,
        "multi_file": False,
        "fields": [
            {"id": "project_name", "label": "Project Name", "default": ""},
            {"id": "duration_months", "label": "Duration (Months)", "default": "12"},
            {"id": "output_format", "label": "Output Format", "type": "select",
             "options": ["DOCX Report", "DOCX + Excel"],
             "default": "DOCX Report"},
        ],
        "processor": process_project_management,
    },
    "report": {
        "name": "Report Writing",
        "description": "Transform raw content into professionally structured reports for WB/EU/ADB",
        "icon": "report",
        "accept": ACCEPTED_UPLOADS,
        "multi_file": False,
        "fields": [
            {"id": "report_style", "label": "Report Style", "type": "select",
             "options": ["World Bank", "European Union", "ADB", "Generic Professional"],
             "default": "Generic Professional"},
            {"id": "report_title", "label": "Report Title", "default": ""},
            {"id": "include_exec_summary", "label": "Include Executive Summary", "type": "select",
             "options": ["Yes", "No"], "default": "Yes"},
        ],
        "processor": process_report,
    },
    "jis_mapping": {
        "name": "JIS Mapping",
        "description": "Generate Results Framework, LogFrame, and M&E mapping from project documents",
        "icon": "jis",
        "accept": ACCEPTED_UPLOADS,
        "multi_file": False,
        "fields": [
            {"id": "framework_type", "label": "Framework Type", "type": "select",
             "options": ["Results Framework", "Logical Framework (LogFrame)", "M&E Matrix", "Full Package"],
             "default": "Full Package"},
            {"id": "sector", "label": "Sector", "default": "Energy"},
        ],
        "processor": process_jis_mapping,
    },
    "cv_rewrite": {
        "name": "CV Reception & Rewriting",
        "description": "Reformat CVs to match WB, EU, or ADB client template requirements",
        "icon": "cv",
        "accept": ACCEPTED_UPLOADS,
        "multi_file": False,
        "fields": [
            {"id": "client_template", "label": "Client Template", "type": "select",
             "options": ["World Bank (WB)", "European Union (EU)", "Asian Development Bank (ADB)"],
             "default": "World Bank (WB)"},
            {"id": "position_title", "label": "Position Title", "default": ""},
            {"id": "years_experience_required", "label": "Min. Years Experience", "default": "10"},
        ],
        "processor": process_cv_rewrite,
    },
    "contract": {
        "name": "Contract Management",
        "description": "Extract and structure deliverables, payments, timelines, and key clauses from contracts",
        "icon": "contract",
        "accept": ACCEPTED_UPLOADS,
        "multi_file": False,
        "fields": [
            {"id": "client_name", "label": "Client Name", "default": ""},
            {"id": "contract_type", "label": "Contract Type", "type": "select",
             "options": ["Standard Consulting", "Framework Agreement", "Supply Contract", "Other"],
             "default": "Standard Consulting"},
        ],
        "processor": process_contract,
    },
}


def get_job(job_id: str) -> dict | None:
    """Return a job definition or None if not found."""
    return JOB_REGISTRY.get(job_id)


def list_jobs() -> list[dict]:
    """Return list of all jobs with metadata (no processor reference)."""
    result = []
    for jid, info in JOB_REGISTRY.items():
        result.append({
            "id": jid,
            "name": info["name"],
            "description": info["description"],
            "icon": info["icon"],
            "accept": info["accept"],
            "multi_file": info["multi_file"],
            "fields": info["fields"],
        })
    return result
