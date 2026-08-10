"""Seed 3 POCs into the multi-POC Google Sheet tracker."""
import tomllib, gspread
from google.oauth2.service_account import Credentials

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

with open(".streamlit/secrets.toml", "rb") as f:
    s = tomllib.load(f)

gc = gspread.authorize(Credentials.from_service_account_info(s["gcp_service_account"], scopes=SCOPES))
ss = gc.open_by_key(s["spreadsheet_id"])

def ws(name): return ss.worksheet(name)

def reset(sheet_name, headers):
    w = ws(sheet_name)
    w.clear()
    w.append_row(headers)
    print(f"  Reset {sheet_name}")

# ── Headers (must match app.py exactly) ──────────────────────────────────────
REGISTRY_H = ["POC_ID","Customer","AE","SE","Status","Created"]
OVERVIEW_H = [
    "POC_ID","Snowflake AE","Snowflake SE",
    "Labcorp Champion","Champion Title","Champion Email",
    "Executive Sponsor","Sponsor Title",
    "Business Unit","Primary Use Case","Secondary Use Case",
    "Cloud Environment","Current Data Platform","Data Volume",
    "Compliance Requirements","POC Start Date","Target Completion Date","Status",
    "POC Objective","Technical Success Criteria","Business Success Criteria",
    "POC Budget ($)","Confirmed Spend ($)","Potential ARR ($)",
    "Procurement Contact","Budget Notes",
]
KPI_H    = ["POC_ID","KPI","Target","Current Value","Unit","Status","Notes"]
ACTION_H = ["POC_ID","Action","Owner","Category","Due Date","Status","Notes"]
UPDATE_H = ["POC_ID","Date","Update","Posted By"]

print("Resetting all sheets…")

# Ensure POC_Registry exists
existing = {w.title for w in ss.worksheets()}
if "POC_Registry" not in existing:
    new_ws = ss.add_worksheet("POC_Registry", rows=100, cols=10)
    new_ws.append_row(REGISTRY_H)
    print("  Created POC_Registry")

for name, headers in [
    ("POC_Registry", REGISTRY_H),
    ("Overview",     OVERVIEW_H),
    ("KPIs",         KPI_H),
    ("Action_Items", ACTION_H),
    ("Updates",      UPDATE_H),
]:
    reset(name, headers)

# ═════════════════════════════════════════════════════════════════════════════
# POC 1 — Labcorp Drug Development
# ═════════════════════════════════════════════════════════════════════════════
P1 = "labcorp-drugdev"

ws("POC_Registry").append_row([P1,"Labcorp","Sarah Mitchell","Geet Sukhmani","Active","2026-07-01"])

ws("Overview").append_row([
    P1,
    "Sarah Mitchell","Geet Sukhmani",
    "Brandon Lee","VP of Data & Analytics","brandon.lee@labcorp.com",
    "Dr. Jennifer Walsh","Chief Data Officer",
    "Drug Development (Biopharma Solutions)","Real-World Evidence (RWE)","Clinical Data Analytics",
    "AWS","Teradata","450 TB - 15M rows/day",
    "HIPAA, 21 CFR Part 11, GxP","2026-07-01","2026-09-30","Active",
    "Validate Snowflake capability to replace Teradata for clinical trial analytics and RWE generation. Demonstrate 10x performance improvement on key workloads.",
    "Query < 30s on 450TB. Snowpark pipeline 15M rows/hr. Zero data loss. HIPAA BAA executed.",
    "3 RWE use cases delivered. Exec sign-off on migration roadmap. Cost reduction > 30%. Team trained on Snowpark.",
    "250000","45000","2400000",
    "Michael Torres - procurement@labcorp.com",
    "POC budget approved Q3 FY26. Full migration budget pending POC success.",
])

for row in [
    [P1,"Query Response Time","< 30","18","seconds","Met","Validated on 450TB clinical dataset. Peak 24s on complex joins."],
    [P1,"Data Ingestion Rate","15000000","22000000","rows/hour","Met","Snowpipe + Kafka. Sustained over 72hr stress test."],
    [P1,"Cost Reduction vs Teradata","30","38","%","Met","3-year TCO: 38% savings including compute, storage, licensing."],
    [P1,"RWE Use Cases Validated","3","2","use cases","On Track","Oncology ✓  Cardiovascular ✓  Rare Disease — pipeline complete, validation pending."],
    [P1,"HIPAA Compliance Validation","Pass","In Review","status","On Track","External legal counsel engaged. BAA in place."],
    [P1,"Data Migration Accuracy","100","99.97","%","On Track","47/47 tables migrated. Minor rounding in 3 numeric columns under investigation."],
]:
    ws("KPIs").append_row(row)

for row in [
    [P1,"Provision HIPAA-compliant Snowflake environment","Snowflake","Technical","2026-07-15","Complete","AWS us-east-1. BAA executed."],
    [P1,"Execute Business Associate Agreement (BAA)","Joint","Compliance","2026-07-15","Complete","Signed by both parties July 14."],
    [P1,"Teradata schema migration — Phase 1 (47 tables)","Joint","Technical","2026-07-31","Complete","All 47 tables migrated. 99.97% accuracy."],
    [P1,"Validate RWE Oncology use case","Customer","Business","2026-08-15","Complete","Signed off by Dr. Walsh's team."],
    [P1,"Snowpark pipeline for 15M rows/hr ingestion","Snowflake","Technical","2026-08-15","Complete","Sustained 22M rows/hr in 72hr test."],
    [P1,"Validate RWE Cardiovascular use case","Customer","Business","2026-08-22","Complete","Sign-off received."],
    [P1,"Legal review of HIPAA BAA — external counsel","Customer","Compliance","2026-08-30","In Progress","Jones Day reviewing. On track."],
    [P1,"Snowpark training — data engineering team","Snowflake","Training","2026-09-12","In Progress","8 of 12 engineers certified."],
    [P1,"Validate RWE Rare Disease use case","Customer","Business","2026-09-10","Open","Pipeline complete. Validation workshop Sept 8."],
    [P1,"Executive business review — CDO presentation","Joint","Executive","2026-09-15","Open","Deck in progress."],
    [P1,"Finalise migration roadmap","Joint","Business","2026-09-25","Open","Pending Rare Disease sign-off."],
    [P1,"Submit POC results for CDO sign-off","Joint","Executive","2026-09-30","Open","Final gate before full migration budget."],
]:
    ws("Action_Items").append_row(row)

for row in [
    [P1,"2026-08-07 10:30","Week 10: Cardiovascular RWE validated. Query performance averaging 18s on 450TB dataset. Snowpipe sustaining 22M rows/hr. Legal engaged for BAA review (Aug 30). TCO shows 38% savings vs Teradata. Two of three use cases complete. On track for Q3 close.","Geet Sukhmani"],
    [P1,"2026-07-25 14:15","Week 8: Oncology RWE signed off. Snowpark at 22M rows/hr. Migration Phase 1 complete — 47 tables, 99.97% accuracy. Minor rounding discrepancy in 3 columns being investigated. Brandon confirmed exec review Sept 15.","Geet Sukhmani"],
    [P1,"2026-07-07 09:00","Kick-off complete. HIPAA environment live on AWS us-east-1. BAA executed. Schema analysis underway on 450TB Teradata environment. Oncology RWE pipeline design started. Training cohort scheduled July 21.","Geet Sukhmani"],
]:
    ws("Updates").append_row(row)

print("✓  Labcorp Drug Development seeded")

# ═════════════════════════════════════════════════════════════════════════════
# POC 2 — Pfizer Genomics
# ═════════════════════════════════════════════════════════════════════════════
P2 = "pfizer-genomics"

ws("POC_Registry").append_row([P2,"Pfizer","James Caldwell","Priya Nair","Active","2026-06-15"])

ws("Overview").append_row([
    P2,
    "James Caldwell","Priya Nair",
    "Dr. Marcus Chen","Head of Genomics Data Engineering","marcus.chen@pfizer.com",
    "Stephanie Rowe","SVP Global Research Technology",
    "Genomics","Genomics & Next-Gen Sequencing","Data Sharing with Pharma Partners",
    "Azure","Oracle Exadata","1.2 PB - 40M variants/day",
    "HIPAA, GxP, GDPR","2026-06-15","2026-10-31","Active",
    "Prove Snowflake can handle Pfizer's petabyte-scale NGS variant data pipeline — replacing Oracle Exadata — while enabling compliant data sharing with 12 external research partners across EU and US.",
    "NGS pipeline processes 40M variants/day. Cross-region data sharing to 3 test partners without data movement. Query on 1.2PB variant store < 45s.",
    "3 external research partners connected via Snowflake Data Sharing. GDPR-compliant data residency validated. Sign-off from SVP Research Technology.",
    "380000","92000","4100000",
    "David Kim - david.kim@pfizer.com",
    "Approved by Pfizer Research Technology council. Multi-year deal contingent on POC success.",
])

for row in [
    [P2,"NGS Variant Processing Rate","40000000","38500000","variants/day","On Track","Nearing target. Optimising Snowpark UDF for variant annotation step."],
    [P2,"Cross-Region Query Performance","< 45","41","seconds","Met","Azure East US ↔ West Europe. Tested on 800TB subset."],
    [P2,"External Partner Data Shares","3","3","partners","Met","Broad Institute, GSK, Novartis connected. Zero data movement."],
    [P2,"GDPR Data Residency Compliance","Pass","Pass","status","Met","EU data stays in Azure West Europe. Validated by DPO team."],
    [P2,"Storage Cost vs Oracle Exadata","25","31","%","Met","31% reduction vs Oracle licensing + hardware."],
    [P2,"Snowflake Cortex AI Adoption","2","1","use cases","On Track","Variant classification model deployed. Drug interaction model in progress."],
]:
    ws("KPIs").append_row(row)

for row in [
    [P2,"Azure Tenant provisioning + Snowflake account setup","Snowflake","Technical","2026-06-25","Complete","East US + West Europe accounts live."],
    [P2,"GDPR data residency architecture review","Joint","Compliance","2026-07-10","Complete","DPO approved. EU data stays in West Europe region."],
    [P2,"Oracle Exadata schema extraction and analysis","Customer","Technical","2026-07-20","Complete","1,247 tables catalogued. 340 priority tables identified."],
    [P2,"Connect Broad Institute as pilot data sharing partner","Joint","Technical","2026-08-01","Complete","Live. Broad accessing variant data via Snowflake Share."],
    [P2,"Deploy NGS variant annotation Snowpark pipeline","Snowflake","Technical","2026-08-20","Complete","Processing 38.5M variants/day. Optimisation in progress for final 1.5M gap."],
    [P2,"Connect GSK and Novartis as data sharing partners","Joint","Business","2026-09-05","Complete","Both partners live. Zero data movement confirmed."],
    [P2,"Cortex AI — drug interaction prediction model","Snowflake","Technical","2026-09-20","In Progress","Model training on variant + EHR dataset. 70% complete."],
    [P2,"Oracle Exadata parallel run and data validation","Joint","Technical","2026-10-01","In Progress","Row counts match on 298/340 tables. 42 tables in reconciliation."],
    [P2,"SVP Research Technology executive review","Joint","Executive","2026-10-15","Open","Priya to present cost, performance, and AI findings."],
    [P2,"Migration go/no-go decision","Customer","Executive","2026-10-31","Open","Pending exec review sign-off."],
]:
    ws("Action_Items").append_row(row)

for row in [
    [P2,"2026-08-20 11:00","Week 10: NGS pipeline processing 38.5M variants/day — 96% of target. Snowpark optimisation underway for variant annotation bottleneck (expected fix Sept 1). All 3 data sharing partners live. GDPR residency validated. 31% storage cost reduction confirmed vs Oracle.","Priya Nair"],
    [P2,"2026-07-22 09:30","Week 6: Oracle schema analysis complete — 1,247 tables, 340 priority tables for migration. Broad Institute data share live and actively used by their research team. Azure dual-region setup validated for GDPR residency.","Priya Nair"],
    [P2,"2026-06-17 14:00","Kick-off: Azure accounts provisioned in East US and West Europe. GDPR architecture reviewed and approved by Pfizer DPO. Schema extraction from Oracle Exadata started. First milestone: GDPR validation by July 10.","Priya Nair"],
]:
    ws("Updates").append_row(row)

print("✓  Pfizer Genomics seeded")

# ═════════════════════════════════════════════════════════════════════════════
# POC 3 — CVS Health Analytics
# ═════════════════════════════════════════════════════════════════════════════
P3 = "cvs-health"

ws("POC_Registry").append_row([P3,"CVS Health","Rachel Torres","Amir Shah","Planning","2026-08-01"])

ws("Overview").append_row([
    P3,
    "Rachel Torres","Amir Shah",
    "Kevin Park","Director of Enterprise Data Platform","kevin.park@cvshealth.com",
    "Linda Gomez","VP Technology & Digital",
    "Technology Solutions","Lab Results Platform Modernisation","Clinical Data Analytics",
    "AWS","Informatica + Redshift","320 TB - 8M lab results/day",
    "HIPAA, CLIA, SOC 2","2026-08-01","2026-11-30","Planning",
    "Modernise CVS Health's lab results platform by replacing Informatica + Redshift with Snowflake — enabling real-time lab result delivery, physician analytics, and patient-facing result notifications at scale.",
    "Ingest 8M lab results/day with < 5 min end-to-end latency. Physician analytics dashboard query < 10s. CLIA audit trail complete in Snowflake.",
    "Real-time patient notification pipeline live. 500 physician users onboarded to analytics dashboard. HIPAA + CLIA compliance validated. 25% cost reduction vs Redshift.",
    "175000","0","1800000",
    "Sandra Wu - sandra.wu@cvshealth.com",
    "Budget approved pending legal review. Formal SOW expected Sept 1.",
])

for row in [
    [P3,"Lab Result Ingestion Latency","< 5 min","Not started","minutes","Not Started","Baseline on Redshift: avg 18 min. Target is sub-5 min with Snowpipe Streaming."],
    [P3,"Physician Dashboard Query Time","< 10","Not started","seconds","Not Started","Current Redshift avg: 45s on standard queries."],
    [P3,"Daily Lab Results Processed","8000000","Not started","results/day","Not Started","Peak load during flu season: 12M results/day."],
    [P3,"Cost Reduction vs Redshift","25","Not started","%","Not Started","Redshift + Informatica current annual cost: $1.4M."],
    [P3,"CLIA Audit Trail Completeness","100","Not started","%","Not Started","Full chain of custody for every result required."],
]:
    ws("KPIs").append_row(row)

for row in [
    [P3,"Legal review and SOW execution","Joint","Compliance","2026-09-01","Open","Sandra Wu (procurement) leading. Snowflake legal engaged."],
    [P3,"Snowflake account provisioning on AWS","Snowflake","Technical","2026-09-08","Open","Pending SOW. US-East-1 selected."],
    [P3,"Informatica pipeline discovery and mapping","Customer","Technical","2026-09-20","Open","Kevin's team to document 340 active Informatica jobs."],
    [P3,"HIPAA + CLIA compliance architecture review","Joint","Compliance","2026-09-30","Open","External compliance consultant engaged by CVS."],
    [P3,"Snowpipe Streaming POC — 8M results/day","Snowflake","Technical","2026-10-15","Open","Target: < 5 min end-to-end latency."],
    [P3,"Physician analytics dashboard prototype","Joint","Business","2026-10-31","Open","500-user cohort for pilot."],
    [P3,"Patient notification pipeline prototype","Snowflake","Technical","2026-11-10","Open","Integration with CVS patient portal."],
    [P3,"VP Technology executive review","Joint","Executive","2026-11-20","Open","Rachel + Amir to present with Linda Gomez."],
]:
    ws("Action_Items").append_row(row)

for row in [
    [P3,"2026-08-01 15:00","Kick-off call completed. Kevin confirmed pain points: 18 min lab result latency (vs 5 min target), Informatica cost ($600K/year), and Redshift hitting performance limits during peak season. Legal review starting Sept 1. Formal SOW expected by end of August. Amir to send technical questionnaire this week.","Amir Shah"],
]:
    ws("Updates").append_row(row)

print("✓  CVS Health seeded")
print("\nAll 3 POCs seeded successfully. Refresh the Streamlit app.")
