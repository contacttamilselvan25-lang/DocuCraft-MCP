"""Comprehensive Test Suite for DocuCraft MCP."""

import os
import sys
import pytest

# Add docucraft_mcp directory to sys.path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from themes import get_theme_config, hex_to_rgb, PRESET_THEMES
from generators.pptx_generator import build_presentation
from generators.docx_generator import build_document
from generators.excel_generator import build_spreadsheet
from generators.pdf_generator import build_pdf
from server import tool_create_suite, mcp


OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_outputs")


def test_theme_engine():
    # Test preset
    cfg = get_theme_config("emerald_growth")
    assert cfg["primary"] == "#065F46"
    assert cfg["primary_rgb"] == (6, 95, 70)
    assert cfg["font_heading"] == "Segoe UI"

    # Test custom override
    custom = get_theme_config("corporate_blue", primary_color="#FF5733", font_family="Trebuchet MS")
    assert custom["primary"] == "#FF5733"
    assert custom["primary_rgb"] == (255, 87, 51)
    assert custom["font_heading"] == "Trebuchet MS"


def test_pptx_generation():
    pptx_path = os.path.join(OUTPUT_DIR, "test_deck.pptx")
    slides = [
        {
            "layout": "stats",
            "title": "Clinical Milestone Metrics",
            "category": "Phase II Results",
            "stats": [
                {"value": "94.2%", "label": "Efficacy Rate", "description": "Primary endpoint achieved"},
                {"value": "< 1.5%", "label": "Adverse Events", "description": "Well-tolerated cohort"},
                {"value": "1,200+", "label": "Patients Screened", "description": "Across 14 trial sites"}
            ]
        },
        {
            "layout": "comparison",
            "title": "Formulation Comparison",
            "category": "Preclinical Benchmarking",
            "column_1": {
                "title": "Standard Suspension",
                "bullets": ["Poor dissolution rate (f2 < 50)", "High peak-to-trough fluctuations", "Twice daily dosing"]
            },
            "column_2": {
                "title": "Nanoparticle Tablet (F3)",
                "bullets": ["Sustained 24h release", "Bioequivalent to innovator", "Once daily oral administration"]
            }
        },
        {
            "layout": "table",
            "title": "Batch Assay Uniformity",
            "category": "Quality Control",
            "headers": ["Batch ID", "Hardness (kg/cm²)", "Friability (%)", "Assay (%)"],
            "rows": [
                ["Batch F1", "6.2", "0.45", "99.4"],
                ["Batch F2", "6.5", "0.38", "100.2"],
                ["Batch F3", "6.8", "0.29", "99.8"]
            ]
        },
        {
            "layout": "bullets",
            "title": "Executive Summary & Next Steps",
            "bullets": [
                "Submission of IND amendment to regulatory agencies scheduled for Q3.",
                "Commercial scale-up planned at GMP contract manufacturing facility.",
                "Target product profile meets all pre-specified clinical endpoints."
            ]
        }
    ]

    res = build_presentation(
        title="Project Apex: Oncology Therapeutic Review",
        subtitle="Phase II Preclinical & Formulation Progress Report",
        slides=slides,
        output_path=pptx_path,
        theme="corporate_blue"
    )

    assert os.path.exists(pptx_path)
    assert os.path.getsize(pptx_path) > 10000  # Non-trivial PPTX
    assert res["slide_count"] == 5  # Title + 4 content slides


def test_docx_generation():
    docx_path = os.path.join(OUTPUT_DIR, "test_document.docx")
    sections = [
        {
            "heading": "1. Executive Summary",
            "paragraphs": [
                "This project proposal details the engineering and validation strategy for the automated document synthesis platform. The goal is to provide end-to-end multi-format publishing capabilities across enterprise ecosystems."
            ],
            "callout": {
                "title": "STRATEGIC OBJECTIVE",
                "text": "Deliver multi-format office document generation directly through natural language agent prompts without requiring human formatting."
            }
        },
        {
            "heading": "2. Formulation Specifications & Release Criteria",
            "paragraphs": [
                "The table below outlines the analytical test parameters and acceptance criteria for commercial release testing:"
            ],
            "table": {
                "headers": ["Test Attribute", "Analytical Method", "Acceptance Criteria", "Result"],
                "rows": [
                    ["Identification", "HPLC-UV", "Retention time matches reference", "Conforms"],
                    ["Assay Content", "HPLC", "98.0% - 102.0%", "99.8%"],
                    ["Dissolution (8h)", "USP Apparatus II (Paddle)", "NLT 85% released", "94.2%"],
                    ["Related Substances", "RP-HPLC", "Individual impurity <= 0.2%", "0.08%"]
                ]
            },
            "bullets": [
                "All stability tests conform to ICH Q1A guidelines.",
                "Re-test period recommended at 24 months under standard storage conditions."
            ]
        }
    ]

    res = build_document(
        title="Commercial Manufacturing & Quality Protocol",
        subtitle="Analytical Chemistry & Quality Assurance Specification",
        sections=sections,
        output_path=docx_path,
        theme="emerald_growth"
    )

    assert os.path.exists(docx_path)
    assert os.path.getsize(docx_path) > 10000


def test_excel_generation():
    xlsx_path = os.path.join(OUTPUT_DIR, "test_spreadsheet.xlsx")
    sheets = [
        {
            "sheet_name": "Equipment Budget",
            "kpi_cards": [
                {"label": "Total Capex", "value": "$145,000"},
                {"label": "Allocated", "value": "$132,500"},
                {"label": "Variance", "value": "-$12,500"},
                {"label": "Approval", "value": "100% OK"}
            ],
            "headers": ["Equipment Item", "Department", "Vendor", "Unit Price", "Quantity", "Total Cost"],
            "rows": [
                ["HPLC System (Binary)", "Analytical Chem", "Waters", 65000.0, 1, 65000.0],
                ["Dissolution Tester (USP II)", "Pharmaceutics", "Electrolab", 24000.0, 1, 24000.0],
                ["Rotary Tablet Press", "Pilot Plant", "Cadmach", 38000.0, 1, 38000.0],
                ["Analytical Balance (0.01mg)", "Quality Control", "Mettler Toledo", 5500.0, 1, 5500.0]
            ],
            "column_formats": {
                "Unit Price": "currency",
                "Total Cost": "currency"
            },
            "include_totals": True
        }
    ]

    res = build_spreadsheet(
        title="Laboratory Capital Expenditure Model 2026",
        sheets=sheets,
        output_path=xlsx_path,
        theme="corporate_blue"
    )

    assert os.path.exists(xlsx_path)
    assert os.path.getsize(xlsx_path) > 4000


def test_pdf_generation():
    pdf_path = os.path.join(OUTPUT_DIR, "test_report.pdf")
    sections = [
        {
            "heading": "1. Clinical Trial Design & Cohort Demographics",
            "paragraphs": [
                "This randomized, double-blind study evaluated the pharmacokinetic profile and clinical safety of the investigational candidate. A total of 120 subjects were randomized across four therapeutic dosage arms."
            ],
            "callout": {
                "title": "KEY REGULATORY FINDING",
                "text": "The test formulation demonstrated bioequivalence across all primary PK endpoints with no severe adverse event reports."
            },
            "table": {
                "headers": ["Dosage Arm", "Subject Count (n)", "Mean Cmax (ng/mL)", "AUC0-inf (ng*h/mL)"],
                "rows": [
                    ["Placebo", "30", "0.0", "0.0"],
                    ["Low Dose (25 mg)", "30", "142.5", "840.2"],
                    ["Medium Dose (50 mg)", "30", "289.1", "1720.5"],
                    ["High Dose (100 mg)", "30", "594.3", "3480.0"]
                ]
            },
            "bullets": [
                "Dose proportionality confirmed across 25 mg to 100 mg range.",
                "Elimination half-life consistent with once-daily therapeutic regimen."
            ]
        }
    ]

    res = build_pdf(
        title="Phase II Pharmacokinetic Summary Report",
        subtitle="Investigational New Drug (IND) Annual Briefing",
        sections=sections,
        output_path=pdf_path,
        theme="warm_executive"
    )

    assert os.path.exists(pdf_path)
    assert os.path.getsize(pdf_path) > 2000


def test_suite_generation():
    suite_dir = os.path.join(OUTPUT_DIR, "test_suite")
    slides = [{"layout": "bullets", "title": "Overview", "bullets": ["Point A", "Point B"]}]
    sections = [{"heading": "Section 1", "paragraphs": ["Content 1"]}]
    sheets = [{"sheet_name": "Data", "headers": ["Col 1", "Col 2"], "rows": [["A", 10]]}]

    res = tool_create_suite(
        project_title="Project Horizon",
        output_dir=suite_dir,
        presentation_slides=slides,
        document_sections=sections,
        spreadsheet_sheets=sheets,
        theme="modern_dark"
    )

    assert res["status"] == "success"
    files = res["generated_files"]
    assert os.path.exists(files["presentation_pptx"])
    assert os.path.exists(files["document_docx"])
    assert os.path.exists(files["spreadsheet_xlsx"])
    assert os.path.exists(files["report_pdf"])
