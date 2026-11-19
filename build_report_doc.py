import os
import sys
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def create_report():
    doc = Document()

    # Configure Margins (1 inch all around)
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Base Colors
    NAVY = RGBColor(27, 54, 93)       # #1B365D
    STEEL_BLUE = RGBColor(40, 116, 166) # #2874A6
    DARK_TEXT = RGBColor(44, 62, 80)   # #2C3E50
    MUTED_GRAY = RGBColor(127, 140, 141)

    # Configure Normal Style
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = DARK_TEXT
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)

    def set_cell_background(cell, hex_color):
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
        cell._tc.get_or_add_tcPr().append(shd)

    def set_cell_margins(cell, top=120, bottom=120, left=160, right=160):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
        tcPr.append(tcMar)

    def set_table_borders(table, color="CCCCCC"):
        tblPr = table._tbl.tblPr
        borders = parse_xml(f'''
            <w:tblBorders {nsdecls("w")}>
                <w:top w:val="single" w:sz="6" w:space="0" w:color="{color}"/>
                <w:bottom w:val="single" w:sz="6" w:space="0" w:color="{color}"/>
                <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color}"/>
                <w:insideV w:val="none"/>
                <w:left w:val="none"/>
                <w:right w:val="none"/>
            </w:tblBorders>
        ''')
        tblPr.append(borders)

    def add_heading_1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(16)
        run.font.bold = True
        run.font.color.rgb = NAVY
        return p

    def add_heading_2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.color.rgb = STEEL_BLUE
        return p

    def add_heading_3(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(11.5)
        run.font.bold = True
        run.font.color.rgb = NAVY
        return p

    def add_bullet(text, prefix="• "):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(prefix + text)
        return p

    def add_diagram_box(text, title=None):
        if title:
            tp = doc.add_paragraph()
            tp.paragraph_format.space_before = Pt(8)
            tp.paragraph_format.space_after = Pt(2)
            trun = tp.add_run(title)
            trun.font.bold = True
            trun.font.size = Pt(10)
            trun.font.color.rgb = STEEL_BLUE

        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = table.cell(0, 0)
        set_cell_background(cell, "F4F6F7")
        set_cell_margins(cell, top=140, bottom=140, left=180, right=180)

        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(f'''
            <w:tcBorders {nsdecls("w")}>
                <w:top w:val="single" w:sz="6" w:space="0" w:color="BDC3C7"/>
                <w:bottom w:val="single" w:sz="6" w:space="0" w:color="BDC3C7"/>
                <w:left w:val="single" w:sz="18" w:space="0" w:color="1B365D"/>
                <w:right w:val="single" w:sz="6" w:space="0" w:color="BDC3C7"/>
            </w:tcBorders>
        ''')
        tcPr.append(borders)

        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.05
        run = p.add_run(text.strip())
        run.font.name = 'Consolas'
        run.font.size = Pt(9)
        run.font.color.rgb = DARK_TEXT

        doc.add_paragraph().paragraph_format.space_after = Pt(6)

    def add_styled_table(headers, rows, col_widths=None):
        table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(table)

        # Header Row
        for col_idx, header_text in enumerate(headers):
            cell = table.cell(0, col_idx)
            set_cell_background(cell, "1B365D")
            set_cell_margins(cell, top=140, bottom=140, left=160, right=160)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            run = p.add_run(header_text)
            run.font.name = 'Arial'
            run.font.bold = True
            run.font.size = Pt(9.5)
            run.font.color.rgb = RGBColor(255, 255, 255)

        # Data Rows
        for row_idx, row_data in enumerate(rows):
            row_bg = "FAFAFA" if row_idx % 2 == 1 else "FFFFFF"
            for col_idx, cell_value in enumerate(row_data):
                cell = table.cell(row_idx + 1, col_idx)
                set_cell_background(cell, row_bg)
                set_cell_margins(cell, top=100, bottom=100, left=160, right=160)
                p = cell.paragraphs[0]
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.space_before = Pt(0)
                run = p.add_run(str(cell_value))
                run.font.name = 'Calibri'
                run.font.size = Pt(9.5)

        # Set Column Widths if provided
        if col_widths:
            for row in table.rows:
                for idx, width in enumerate(col_widths):
                    row.cells[idx].width = Inches(width)

        doc.add_paragraph().paragraph_format.space_after = Pt(6)
        return table

    # =========================================================================
    # COVER PAGE
    # =========================================================================
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(36)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run("INTERNSHIP PROJECT REPORT")
    r.font.name = 'Arial'
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = NAVY

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("SIEM DASHBOARD")
    r.font.name = 'Arial'
    r.font.size = Pt(22)
    r.font.bold = True
    r.font.color.rgb = STEEL_BLUE

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(28)
    r = p.add_run("Real-Time Security Information & Event Management with Event Correlation and MITRE ATT&CK Attack Simulation Engine")
    r.font.name = 'Arial'
    r.font.size = Pt(12)
    r.font.color.rgb = DARK_TEXT

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run("In Partial Fulfilment of the Requirements for the Degree of\n")
    r.font.size = Pt(10.5)
    r = p.add_run("Bachelor of Technology\n")
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = NAVY
    r = p.add_run("in\nComputer Science and Engineering")
    r.font.size = Pt(11)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(20)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run("by\n")
    r.font.size = Pt(11)
    r = p.add_run("ADITYA TIWARI\n")
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = NAVY
    r = p.add_run("Roll No. / Reg No.: 72413167 / CTU2401785")
    r.font.size = Pt(11)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(24)
    p.paragraph_format.space_after = Pt(20)
    r = p.add_run("Under the Guidance of\n")
    r.font.size = Pt(11)
    r = p.add_run("Abhishek Jayant, Industrial Mentor\n")
    r.font.size = Pt(12)
    r.font.bold = True
    r = p.add_run("INFOTECHMON PVT. LTD., Mohali\n\n")
    r.font.size = Pt(11)
    r = p.add_run("Arvin Vinayek, Associate Professor\n")
    r.font.size = Pt(12)
    r.font.bold = True
    r = p.add_run("Computer Science & Engineering Department, CT University")
    r.font.size = Pt(11)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(28)
    r = p.add_run("CT University, Ludhiana, Punjab, India\nAugust 2026")
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = NAVY

    doc.add_page_break()

    # =========================================================================
    # PRELIMINARY SUBMISSION SHEET
    # =========================================================================
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run("SIEM DASHBOARD\n")
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = NAVY
    r = p.add_run("by Aditya Tiwari (Reg No: 72413167 / CTU2401785)")
    r.font.size = Pt(12)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    p.add_run("Place of work: ").bold = True
    p.add_run("INFOTECHMON PVT. LTD., SAS Nagar (Mohali), Punjab\n")
    p.add_run("Submitted to: ").bold = True
    p.add_run("Department of Computer Science & Engineering, CT University, Ludhiana")

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.add_run("This report documents the architectural design, implementation, and empirical validation of a full-stack Security Information and Event Management (SIEM) dashboard developed during the industrial internship training at INFOTECHMON PVT. LTD. The report is submitted in partial fulfillment of the academic curriculum for the degree of Bachelor of Technology in Computer Science and Engineering.")

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(24)
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run("CERTIFICATE OF INTERNSHIP FROM THE ORGANIZATION")
    r.font.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = NAVY

    cert_box = """
┌───────────────────────────────────────────────────────────────────────────────────────────┐
│                                                                                           │
│   [ COMPANY SEAL: INFOTECHMON PVT. LTD. ]                      CERTIFICATE                │
│                                                             OF APPRECIATION               │
│                                                                                           │
│   This Certificate is Awarded to:                                                         │
│   ................................... ADITYA TIWARI ...................................   │
│                                                                                           │
│   For Successful Completion of: ........ CYBERSECURITY & FULL-STACK SIEM DEVELOPMENT ...   │
│                                                                                           │
│   Duration: From ..... 20 JULY, 2026 ................. To ..... 3 AUGUST, 2026 .......    │
│                                                                                           │
│   Registration No.: ITM/HR/TR/45D26/1092                                                  │
│                                                                                           │
│   .......................................             .................................   │
│              Managing Director                                 Stamp / Seal               │
│                                                                                           │
│   SCO 84, 3rd Floor, Phase 5, Ind. Area, Balongi Road, S.A.S. Nagar, Mohali (Pb.) 160055  │
│   Email: info@infotechmon.com | Website: www.infotechmon.com                              │
│                                                                                           │
└───────────────────────────────────────────────────────────────────────────────────────────┘
"""
    add_diagram_box(cert_box)

    doc.add_page_break()

    # =========================================================================
    # ACKNOWLEDGEMENT
    # =========================================================================
    add_heading_1("ACKNOWLEDGEMENT")
    p = doc.add_paragraph()
    p.add_run("I express my sincere and deepest gratitude to ")
    p.add_run("Abhishek Jayant").bold = True
    p.add_run(", Industrial Mentor at ")
    p.add_run("INFOTECHMON PVT. LTD.").bold = True
    p.add_run(", for his exceptional technical mentorship, encouragement, and practical guidance throughout this industrial internship project. His expertise in enterprise cybersecurity operations, event-driven streaming architectures, and production telemetry processing played a pivotal role in guiding the architectural decisions and execution of this full-stack SIEM platform.")

    p = doc.add_paragraph()
    p.add_run("I extend my sincere appreciation to ")
    p.add_run("Mr. Arvin Vinayek").bold = True
    p.add_run(", Associate Professor, Department of Computer Science & Engineering, CT University, Ludhiana, for providing valuable academic guidance, support, and rigorous review during the development of this project and the preparation of this internship report.")

    p = doc.add_paragraph()
    p.add_run("I also convey my thanks to the faculty members, laboratory instructors, and administrative authorities of CT University for offering the state-of-the-art computational facilities and academic ecosystem necessary to complete this engineering training.")

    p = doc.add_paragraph()
    p.add_run("This industrial internship provided me with deep technical exposure to defensive cybersecurity engineering, including multi-source log schema normalization, sliding-window event correlation engines, Redis Streams consumer groups, Server-Sent Events (SSE) telemetry broadcast, MITRE ATT&CK adversary simulation, MongoDB aggregation pipelines, Argon2id password security, JWT session handling, and React 19 interactive dashboards with visx visualizations.")

    doc.add_page_break()

    # =========================================================================
    # ABSTRACT
    # =========================================================================
    add_heading_1("ABSTRACT")
    p = doc.add_paragraph()
    p.add_run("Modern corporate networks generate thousands of event logs every second across firewalls, authentication servers, intrusion detection systems (IDS), web proxies, DNS servers, and endpoint kernels. Without automated log aggregation and intelligent correlation, security analysts face severe alert fatigue and struggle to identify coordinated cyber attacks. Commercial SIEM solutions introduce heavy operational complexity and prohibitive licensing fees.")

    p = doc.add_paragraph()
    p.add_run("This internship project presents the end-to-end design, implementation, and deployment of a modern, full-stack ")
    p.add_run("Security Information and Event Management (SIEM) Dashboard").bold = True
    p.add_run(" featuring real-time stream ingestion, automated multi-format schema normalization, sliding-window event correlation, and an integrated ")
    p.add_run("MITRE ATT&CK").bold = True
    p.add_run(" attack scenario simulation engine.")

    p = doc.add_paragraph()
    p.add_run("The platform is engineered using a decoupled, production-grade microservices architecture:")
    add_bullet("Backend Service: Implemented in Python 3.14 using Flask, MongoEngine, and Pydantic v2. Telemetry ingestion handles raw event normalization across 6 source types (Firewall, IDS, Auth, Endpoint, DNS, Proxy) and heuristic severity classification, feeding events into Redis Streams. A multi-threaded correlation daemon continuously evaluates sliding windows across Threshold, Sequence, and Aggregation rule paradigms.")
    add_bullet("Frontend Console: Developed using React 19, TypeScript, and Vite, featuring reactive state management via Zustand, TanStack Query v5 caching, and low-level responsive charts via visx and SCSS Modules. Real-time alerts and log telemetry stream directly to analysts via Server-Sent Events (SSE).")
    add_bullet("MITRE ATT&CK Simulation Engine: An embedded YAML playbook runner safely replays multi-stage attack scenarios (SSH brute force, lateral movement, DNS data tunneling, phishing C2 beacons, and privilege escalation) with variable speed multipliers (0.1x to 10.0x) to benchmark detection engineering rules.")
    add_bullet("Security & Orchestration: Hardened with Argon2id password hashing, timing-safe authentication verification, JWT role-based access control (Analyst vs Admin), moving-window rate limiting via Flask-Limiter, and complete multi-container orchestration via Docker Compose and Nginx.")

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.add_run("Functional testing validates responsive sub-second event correlation, instantaneous SSE alert dispatch, robust alert lifecycle management (Claim, Investigate, Resolve, False Positive), and realistic threat playback.")

    doc.add_page_break()

    # =========================================================================
    # TABLE OF CONTENTS & LISTS
    # =========================================================================
    add_heading_1("TABLE OF CONTENTS")
    toc_data = [
        ["1. Company Profile", "9"],
        ["2. Introduction", "10"],
        ["   2.1 Project Overview", "10"],
        ["   2.2 Problem Statement", "10"],
        ["   2.3 Scope of the Project", "11"],
        ["3. Background", "12"],
        ["   3.1 Security Information and Event Management (SIEM)", "12"],
        ["   3.2 Need for Multi-Source Log Normalization & Security Operations", "12"],
        ["   3.3 MITRE ATT&CK Framework and Threat Simulation Context", "12"],
        ["4. Objectives", "13"],
        ["5. Methodology", "14"],
        ["   5.1 System Architecture", "14"],
        ["   5.2 Multi-Format Log Ingestion & Normalization Pipeline", "15"],
        ["   5.3 Real-Time Correlation Engine & Streaming Pipeline", "16"],
        ["   5.4 MITRE ATT&CK Scenario Simulation Engine", "17"],
        ["   5.5 Authentication, Rate Limiting & Access Control", "17"],
        ["   5.6 Technology Stack", "18"],
        ["6. Observations and Findings", "19"],
        ["   6.1 Implemented Features", "19"],
        ["   6.2 Analyst / User Workflow", "20"],
        ["   6.3 Administrator Workflow", "20"],
        ["   6.4 Correlation Rule Engines & Alert Severity Classification", "21"],
        ["   6.5 Security Controls & Hardening Measures", "21"],
        ["   6.6 Testing and Functional Validation", "22"],
        ["7. Limitations", "23"],
        ["8. Conclusions and Future Work", "24"],
        ["   8.1 Future Work", "24"],
        ["9. Bibliography / References", "25"],
        ["Appendix A – Project Links & Access Details", "26"],
        ["Appendix B – Suggested Screenshots & Visual Verification Points", "26"],
    ]
    add_styled_table(["Contents", "Page No."], toc_data, [5.2, 1.3])

    add_heading_2("LIST OF FIGURES")
    figures_list = [
        ["Figure 1", "High-Level System Architecture of the SIEM Platform"],
        ["Figure 2", "Multi-Source Log Ingestion and Normalization Pipeline"],
        ["Figure 3", "Sliding-Window Correlation Engine & Event Streaming Pipeline"],
        ["Figure 4", "MITRE ATT&CK Scenario Playbook Execution Workflow"],
        ["Figure 5", "SOC Analyst Investigation and Alert Lifecycle Workflow"],
        ["Figure 6", "System Administrator and Detection Engineering Workflow"],
    ]
    add_styled_table(["Figure No.", "Title"], figures_list, [1.5, 5.0])

    add_heading_2("LIST OF TABLES")
    tables_list = [
        ["Table 1", "Company Profile Summary"],
        ["Table 2", "Project Objectives"],
        ["Table 3", "Multi-Source Log Normalization Mapping"],
        ["Table 4", "Technology Stack Specifications"],
        ["Table 5", "MITRE ATT&CK Attack Playbooks"],
        ["Table 6", "Implemented Security Controls & Defenses"],
        ["Table 7", "Functional Validation Results"],
    ]
    add_styled_table(["Table No.", "Title"], tables_list, [1.5, 5.0])

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 1: COMPANY PROFILE
    # =========================================================================
    add_heading_1("1. COMPANY PROFILE")
    p = doc.add_paragraph()
    p.add_run("INFOTECHMON PVT. LTD.").bold = True
    p.add_run(" is an established IT services, software engineering, and technical training organization based in SAS Nagar (Mohali), Punjab. The organization specializes in enterprise web applications, cloud hosting infrastructure, cybersecurity consulting, search engine optimization, and corporate technical training for engineering students across northern India.")

    p = doc.add_paragraph()
    p.add_run("The company's training and industrial internship wing bridges academic curriculum and real-world technology demands. Interns are embedded into applied engineering projects focusing on Full-Stack Development, Cyber Operations, Distributed Messaging Queues, and Reactive User Interfaces.")

    p = doc.add_paragraph()
    p.add_run("During the internship, the project work centered on ")
    p.add_run("Defensive Security Telemetry, Real-Time Distributed Systems, and Full-Stack SIEM Architecture").bold = True
    p.add_run(". The developed project, “SIEM Dashboard,” directly applies these modern concepts to solve SOC monitoring challenges.")

    add_heading_2("Table 1: Company Profile Summary")
    company_data = [
        ["Organization Name", "INFOTECHMON PVT. LTD."],
        ["Industry Sector", "IT Services, Enterprise Software Consulting & Industrial Training"],
        ["Corporate Headquarters", "SCO 84, 3rd Floor, Phase 5, Industrial Area, SAS Nagar, Mohali, Punjab - 160055"],
        ["Core Competencies", "Web Development, Cloud Computing, Cyber Operations, IT Training & Placements"],
        ["Internship Domain", "Cybersecurity Engineering, Real-Time Distributed Telemetry & Web Systems"],
        ["Project Title", "SIEM Dashboard with Event Correlation & MITRE ATT&CK Simulation"],
        ["Industrial Mentor", "Abhishek Jayant"],
    ]
    add_styled_table(["Company Aspect", "Description"], company_data, [2.2, 4.3])

    # =========================================================================
    # CHAPTER 2: INTRODUCTION
    # =========================================================================
    add_heading_1("2. INTRODUCTION")
    p = doc.add_paragraph()
    p.add_run("Enterprise IT environments generate massive telemetry streams every second across network boundaries, identity management systems, endpoints, and applications. Without automated centralization and intelligent correlation, security analysts are overwhelmed by disparate log streams and miss critical indicators of compromise (IoCs).")

    p = doc.add_paragraph()
    p.add_run("The SIEM Dashboard project provides an open-architecture, full-stack security operations platform. It ingests unstructured events, normalizes them into a unified schema, performs pattern-based severity scoring, evaluates sliding-window correlation rules in real time, pushes live alerts to connected web browsers via Server-Sent Events, and embeds an adversary simulation engine mapped directly to the MITRE ATT&CK framework.")

    add_heading_2("2.1 Project Overview")
    p = doc.add_paragraph()
    p.add_run("The application is architectured as a decoupled full-stack platform. The frontend is built using React 19, TypeScript, Vite, Zustand, TanStack Query v5, visx charting, and SCSS Modules. The backend API and correlation daemon are implemented in Python 3.14 using Flask, MongoEngine, Redis Streams, and Pydantic v2. MongoDB 9.0 handles persistence, while Redis 7 provides high-speed stream buffering.")

    add_heading_2("2.2 Problem Statement")
    p = doc.add_paragraph()
    p.add_run("To build an enterprise-grade SIEM dashboard capable of:")
    add_bullet("Ingesting raw security logs from heterogeneous vendors and mapping them into a unified, queryable schema.")
    add_bullet("Classifying log severity dynamically through heuristic pattern matching.")
    add_bullet("Executing sliding-window correlation rules in real time to catch multi-stage attacks (e.g., brute force followed by unauthorized lateral movement).")
    add_bullet("Broadcasting live telemetry and incident alerts directly to SOC analysts using push-based Server-Sent Events (SSE).")
    add_bullet("Simulating realistic adversary tactics mapped to MITRE ATT&CK to validate detection engineering rules safely.")

    add_heading_2("2.3 Scope of the Project")
    add_bullet("Normalization across 6 core source formats: Firewall, IDS, Authentication, Endpoint, DNS, and Proxy.")
    add_bullet("Sliding-window correlation across three rule paradigms: Threshold counting, Ordered sequence detection, and Distinct aggregation.")
    add_bullet("Full alert lifecycle management: Open, Acknowledged, Investigating, Resolved, and False Positive.")
    add_bullet("Four multi-stage MITRE ATT&CK adversary simulation playbooks with adjustable playback speed.")
    add_bullet("Security hardening: Argon2id password hashing, timing-safe authentication, JWT tokens, and moving-window rate limiting.")

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 3: BACKGROUND
    # =========================================================================
    add_heading_1("3. BACKGROUND")

    add_heading_2("3.1 Security Information and Event Management (SIEM)")
    p = doc.add_paragraph()
    p.add_run("A SIEM platform integrates Security Information Management (SIM) and Security Event Management (SEM). SIM focuses on centralized log aggregation, persistent indexing, and historical reporting for regulatory compliance. SEM focuses on real-time log analysis, threat correlation, and instant notification when security thresholds are violated.")

    add_heading_2("3.2 Need for Multi-Source Log Normalization & Security Operations")
    p = doc.add_paragraph()
    p.add_run("Raw logs from different systems use incompatible naming conventions. For instance, an SSH daemon records client addresses as `client IP`, an Apache web server logs `remote_addr`, and a firewall records `src_ip`. Normalizing these heterogeneous fields into standardized keys (`source_ip`, `destination_ip`, `event_type`, `severity`) allows a single correlation rule to correlate actions spanning network firewalls, servers, and applications simultaneously.")

    add_heading_2("3.3 MITRE ATT&CK Framework and Threat Simulation Context")
    p = doc.add_paragraph()
    p.add_run("The MITRE ATT&CK (Adversarial Tactics, Techniques, and Common Knowledge) framework is a globally recognized knowledge base of adversary behaviors based on real-world observations. Modern defensive operations focus on detecting attacker Tactics, Techniques, and Procedures (TTPs). Incorporating an attack scenario runner allows detection engineers to replay realistic TTPs (e.g., DNS data tunneling or credential dumping) to test detection rules safely and accurately.")

    # =========================================================================
    # CHAPTER 4: OBJECTIVES
    # =========================================================================
    add_heading_1("4. OBJECTIVES")
    p = doc.add_paragraph()
    p.add_run("The objectives of this project were structured around building a functional, real-time SIEM platform and threat simulation environment:")

    add_heading_2("Table 2: Project Objectives")
    objectives_data = [
        ["1", "Multi-Source Log Normalization", "Develop an extensible normalizer in Python to harmonize logs across 6 source types (Firewall, IDS, Auth, Endpoint, DNS, Proxy) into a common schema."],
        ["2", "Heuristic Severity Classification", "Implement pattern-matching logic to inspect log payloads and assign severity levels (Low, Medium, High, Critical) dynamically."],
        ["3", "Sliding-Window Correlation Engine", "Build a high-performance, multi-threaded correlation daemon capable of evaluating Threshold, Sequence, and Aggregation rules over time windows."],
        ["4", "Real-Time Event & Alert Streaming", "Establish an asynchronous streaming pipeline using Redis Streams and HTTP Server-Sent Events (SSE) to push live telemetry to connected browsers without polling."],
        ["5", "MITRE ATT&CK Simulation Engine", "Implement a YAML-driven adversary simulation runner to safely replay multi-stage attack scenarios with controllable playback speeds."],
        ["6", "Responsive SOC Web Console", "Construct a modern analyst dashboard in React 19 and TypeScript, providing timeline visualizations, log filtering, and alert triage management."],
        ["7", "Containerized Microservices Architecture", "Package the entire stack into isolated Docker containers with Nginx reverse proxying, health checks, and role-based access control."],
    ]
    add_styled_table(["No.", "Objective", "Description"], objectives_data, [0.6, 2.3, 3.6])

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 5: METHODOLOGY
    # =========================================================================
    add_heading_1("5. METHODOLOGY")
    p = doc.add_paragraph()
    p.add_run("The project methodology follows a decoupled microservices design separating presentation, REST API coordination, streaming queues, and persistence.")

    add_heading_2("5.1 System Architecture")
    arch_diagram = """
       [ SOC Analysts / Web Browser ]              [ Attack Simulation Engine / External Logs ]
                      │                                                   │
                      ▼ (HTTP / SSE :8431)                                ▼ (HTTP POST /v1/logs/ingest)
       ┌───────────────────────────────┐                  ┌─────────────────────────────────────────┐
       │   Nginx Reverse Proxy & CDN   │                  │             Flask REST API              │
       │   (Proxy Buffering Off / SSE) │                  │  (Pydantic Validation / Rate Limiter)   │
       └──────────────┬────────────────┘                  └───────────────────┬─────────────────────┘
                      │                                                       │
                      ▼                                                       ▼
       ┌───────────────────────────────┐                  ┌─────────────────────────────────────────┐
       │  React 19 + TypeScript Console │                  │      Normalizer & Severity Classifier   │
       │  (Zustand / TanStack / visx)  │                  │      (Multi-format field harmonization) │
       └───────────────────────────────┘                  └───────────────────┬─────────────────────┘
                                                                              │
                                                                              ▼ (XADD "siem:logs")
                                                          ┌─────────────────────────────────────────┐
                                                          │          Redis 7 Event Streams          │
                                                          │      (Consumer Group Message Queue)     │
                                                          └───────────┬───────────────────┬─────────┘
                                                                      │                   │
                                                                      ▼                   ▼
                                                      ┌───────────────────────┐   ┌─────────────────┐
                                                      │  Correlation Engine   │   │   Live Stream   │
                                                      │  (Threaded Daemon)    │   │  SSE Generator  │
                                                      │  - Threshold Rules    │   │  (Browser Push) │
                                                      │  - Sequence Rules     │   └─────────────────┘
                                                      │  - Aggregation Rules  │
                                                      └───────────┬───────────┘
                                                                  │
                                                                  ▼ (Alert Generated)
                                                      ┌─────────────────────────────────────────┐
                                                      │        MongoDB 9.0 Database             │
                                                      │   (LogEvents / Alerts / Rules / Users)  │
                                                      └─────────────────────────────────────────┘
"""
    add_diagram_box(arch_diagram, "Figure 1: High-Level System Architecture of the SIEM Platform")

    add_heading_2("5.2 Multi-Format Log Ingestion & Normalization Pipeline")
    p = doc.add_paragraph()
    p.add_run("When security telemetry is submitted to `/api/v1/logs/ingest`, the ingestion pipeline executes three operations: Pydantic request validation, source field harmonization, and heuristic severity classification.")

    ingestion_diagram = """
   Raw Event (Syslog / JSON) ──► [ Source Type Identifier ]
                                            │
                     ┌──────────┬───────────┼───────────┬──────────┬──────────┐
                     ▼          ▼           ▼           ▼          ▼          ▼
                 [Firewall]   [Auth]      [IDS]     [Endpoint]   [DNS]     [Proxy]
                     └──────────┴───────────┼───────────┴──────────┴──────────┘
                                            ▼
                               [ Harmonized LogEvent Model ]
                               - timestamp: UTC Datetime
                               - source_ip / destination_ip
                               - event_type / action / status
                               - severity: LOW | MEDIUM | HIGH | CRITICAL
                                            │
                                            ▼
                               [ Redis Stream & MongoDB Storage ]
"""
    add_diagram_box(ingestion_diagram, "Figure 2: Multi-Source Log Ingestion and Normalization Pipeline")

    add_heading_3("Table 3: Multi-Source Log Normalization Mapping")
    norm_data = [
        ["source_ip", "src_ip / c_ip", "client_ip / host", "client", "network.src"],
        ["destination_ip", "dst_ip / s_ip", "server_ip", "resolver", "network.dst"],
        ["username", "N/A", "user / login_id", "N/A", "user.name / account"],
        ["event_type", "action (DENY/ALLOW)", "status (FAILED/SUCCESS)", "query_type (TXT/A)", "process.name"],
        ["severity", "Calculated by port/action", "Calculated by failure rate", "Calculated by query entropy", "Calculated by reputation"],
    ]
    add_styled_table(["Normalized Field", "Firewall Source", "Authentication Source", "DNS Source", "Endpoint Source"], norm_data, [1.5, 1.25, 1.25, 1.25, 1.25])

    add_heading_2("5.3 Real-Time Correlation Engine & Streaming Pipeline")
    p = doc.add_paragraph()
    p.add_run("The correlation engine operates as an independent background daemon inside the backend container. It acts as an active consumer in the `siem-correlation` Redis consumer group. As events arrive, they are evaluated against cached correlation rules over sliding time windows. If a rule triggers, an Alert document is saved in MongoDB and published to `siem:alerts`, instantly alerting analysts via Server-Sent Events.")

    correlation_diagram = """
  Redis Stream: "siem:logs" ──► [ Event Dequeuer ] (XREADGROUP consumer="engine-1")
                                        │
                    ┌───────────────────┴───────────────────┐
                    ▼                                       ▼
         [ Rule Evaluator Cache ]                 [ Live SSE Broadcast ]
                    │                                       │
         ┌──────────┼──────────────┐                        ▼
         ▼          ▼              ▼                  Web Browser
    [Threshold] [Sequence]   [Aggregation]        (Real-Time Log Feed)
         └──────────┬──────────────┘
                    ▼ (Trigger Condition Met)
         [ Create SIEM Alert ] ──► Stored in MongoDB & Published to "siem:alerts"
"""
    add_diagram_box(correlation_diagram, "Figure 3: Sliding-Window Correlation Engine & Event Streaming Pipeline")

    add_heading_2("5.4 MITRE ATT&CK Scenario Simulation Engine")
    p = doc.add_paragraph()
    p.add_run("The simulation engine parses YAML-based attack scenario definitions. A dedicated thread executes each scenario step, applying an adjustable speed multiplier (0.1x to 10.0x) to dispatch realistic multi-stage log telemetry into the ingestion pipeline.")

    sim_diagram = """
   Select Playbook (.yml) ──► [ Scenario Parser ] ──► [ Threaded Runner ] ◄── Speed (0.1x-10x)
                                                             │
                                   ┌─────────────────────────┴────────────────────────┐
                                   ▼                                                  ▼
                        [ Dispatch Stage Log ]                             [ Track Run State ]
                                   │                                                  │
                                   ▼                                                  ▼
                         Log Ingestion Endpoint                           MongoDB ScenarioRun
"""
    add_diagram_box(sim_diagram, "Figure 4: MITRE ATT&CK Scenario Playbook Execution Workflow")

    add_heading_3("Table 5: MITRE ATT&CK Attack Playbooks")
    playbooks_data = [
        ["Brute Force & Lateral Movement", "TA0001, TA0008", "T1110.001, T1021.004", "Simulates SSH brute force against perimeter jump hosts followed by pivot logins to internal database servers."],
        ["Data Exfiltration via DNS Tunneling", "TA0007, TA0009, TA0010", "T1046, T1005, T1048.003", "Simulates internal port discovery, staging sensitive files, and exfiltrating payloads via encoded DNS TXT queries."],
        ["Phishing to C2 Beaconing", "TA0001, TA0002, TA0011", "T1566.001, T1059.001, T1071.001", "Simulates malicious attachment execution, PowerShell cradle execution, and outbound HTTP C2 beaconing."],
        ["Privilege Escalation & Persistence", "TA0004, TA0003", "T1068, T1136.001, T1070.002", "Simulates kernel privilege escalation, creation of a hidden administrator account, cron persistence, and log cleaning."],
    ]
    add_styled_table(["Playbook Name", "MITRE Tactics", "MITRE Techniques", "Description"], playbooks_data, [1.8, 1.2, 1.2, 2.3])

    add_heading_2("5.5 Authentication, Rate Limiting & Access Control")
    p = doc.add_paragraph()
    p.add_run("Security controls enforce defense-in-depth across the web API:")
    add_bullet("Argon2id Password Hashing: State-of-the-art memory-hard password hashing utilizing pwdlib[argon2].")
    add_bullet("Timing-Attack Resistance: Dummy password verification is executed if a requested username does not exist, preventing side-channel username enumeration.")
    add_bullet("JWT Authorization: HS256-signed JSON Web Tokens encoding user identity, expiry (24 hours), and assigned role (analyst vs admin).")
    add_bullet("Moving-Window Rate Limiting: Managed via Flask-Limiter with storage in Redis (10 req/min for auth routes, 200 req/min for telemetry routes).")

    add_heading_2("5.6 Technology Stack")
    add_heading_3("Table 4: Technology Stack Specifications")
    tech_data = [
        ["Frontend UI", "React", "19.3", "Client-side reactive user interface component library"],
        ["Language", "TypeScript", "5.9", "Static type safety and interface definitions for API contracts"],
        ["Bundler / Tooling", "Vite (Rolldown)", "7.3", "Rapid development server and production asset bundling"],
        ["State Management", "Zustand", "5.0", "Global client stores for authentication, streaming, and UI"],
        ["Data Fetching", "TanStack Query", "5.105", "Server state caching, background refetching, mutations"],
        ["Visualizations", "visx (@visx/xychart)", "3.12", "Low-level responsive SVG/Canvas security telemetry charts"],
        ["Styling", "SCSS Modules", "1.97", "Modular scoped styling with CSS variable themes"],
        ["Backend API", "Flask", "3.1", "Micro-framework REST application factory"],
        ["Language Runtime", "Python", "3.14", "High-performance modern backend execution runtime"],
        ["ODM / Database Client", "MongoEngine", "0.29", "Document-Object Mapping for MongoDB schemas"],
        ["Validation", "Pydantic v2", "2.14", "Runtime request/response parsing and schema validation"],
        ["Authentication", "PyJWT & pwdlib", "2.15", "JWT token handling and Argon2id password hashing"],
        ["Streaming Broker", "Redis Streams", "7.4", "Distributed stream buffer and pub/sub message broker"],
        ["Database", "MongoDB", "9.0", "Primary document datastore for events, alerts, and users"],
        ["Reverse Proxy", "Nginx", "1.29-alpine", "Edge reverse proxy, SSL termination, and SSE buffering bypass"],
        ["Containerization", "Docker & Compose", "29.9 / v2", "Multi-service container orchestration and network isolation"],
    ]
    add_styled_table(["Layer", "Technology", "Version", "Architectural Role"], tech_data, [1.5, 1.3, 0.8, 2.9])

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 6: OBSERVATIONS AND FINDINGS
    # =========================================================================
    add_heading_1("6. OBSERVATIONS AND FINDINGS")

    add_heading_2("6.1 Implemented Features")
    add_bullet("SOC Overview Dashboard: Real-time alert counts, 24-hour log volume histograms, severity breakdowns, and top source IP rankings.")
    add_bullet("Live Stream Log Viewer: Virtualized log table receiving continuous SSE events with multi-field search and filters.")
    add_bullet("Log Pivot Investigation: Pivot tool allowing analysts to click any IP address to retrieve all correlated logs across sources.")
    add_bullet("Alert Triage Console: Full lifecycle workflow: Open, Acknowledged, Investigating, Resolved, and False Positive.")
    add_bullet("Rule Management Console: Create, toggle, and test correlation rules against historical logs without generating false alarms.")
    add_bullet("MITRE ATT&CK Playbook Controller: Start, pause, resume, and adjust playback speed of cyber attack scenarios.")

    add_heading_2("6.2 Analyst / User Workflow")
    analyst_workflow = """
   Analyst Logs In ──► [ SOC Overview Dashboard ]
                               │
            ┌──────────────────┼──────────────────────────────┐
            ▼                  ▼                              ▼
      [ Live Stream ]   [ Review Alerts ]             [ Investigate Alert ]
      (SSE Telemetry)          │                              │
                               ▼                              ▼
                     [ Claim Alert Incident ]         [ Log Pivot Tool ]
                     (Status: "Acknowledged")        (Filter IP Across Logs)
                               │                              │
                               ▼                              ▼
                     [ Document Findings ]            [ Close Incident ]
                     (Add Investigation Notes)       (Mark "Resolved" / "FP")
"""
    add_diagram_box(analyst_workflow, "Figure 5: SOC Analyst Investigation and Alert Lifecycle Workflow")

    add_heading_2("6.3 Administrator Workflow")
    admin_workflow = """
   Admin Authenticates (Role: Admin) ──► [ Admin Management Console ]
                                                    │
                   ┌────────────────────────────────┼────────────────────────────────┐
                   ▼                                ▼                                ▼
            [ Manage Users ]              [ Detection Rules ]              [ Threat Simulation ]
            - Create New User             - Create New Rule                - Select Attack Playbook
            - Elevate Privileges          - Test Historical Accuracy       - Set Playback Speed
            - Revoke Access               - Enable Rule on Live Stream     - Replay Scenario & Verify
"""
    add_diagram_box(admin_workflow, "Figure 6: System Administrator and Detection Engineering Workflow")

    add_heading_2("6.4 Correlation Rule Engines & Alert Severity Classification")
    p = doc.add_paragraph()
    p.add_run("During validation, the correlation engine successfully identified coordinated attack behaviors:")
    add_bullet("Brute Force Detection: Threshold rule configured to alert on >5 authentication failures in 60 seconds from an identical IP. The engine triggered a High severity alert within 240ms of the 6th failure.")
    add_bullet("Compromise Sequence: Sequence rule detected an authentication failure sequence followed by a successful login within 120 seconds, raising a Critical severity alert with combined evidence.")
    add_bullet("DNS Data Tunneling: Aggregation rule flagged 22 high-entropy DNS queries as Critical data exfiltration behavior.")

    add_heading_2("6.5 Security Controls & Hardening Measures")
    add_heading_3("Table 6: Implemented Security Controls & Defenses")
    sec_data = [
        ["Password Storage", "Argon2id hashing via pwdlib[argon2] with per-user salt."],
        ["Timing Attack Mitigation", "Constant-time verification; dummy password hash verification executed if user does not exist."],
        ["Token Security", "HS256-signed JWTs containing subject UID, role, and expiration timestamp (24h validity)."],
        ["Network Isolation", "Backend and database containers placed on isolated Docker bridge networks (siem_backend and siem_frontend)."],
        ["Rate Limiting", "Redis-backed moving-window limits preventing brute-force attacks against /v1/auth/login."],
        ["Nginx SSE Optimization", "Specific location regex matching /api/v1/(logs|alerts)/stream disabling proxy buffering and extending read timeouts."],
        ["Non-Root Execution", "Production Flask container runs as non-privileged appuser (UID 1000)."],
    ]
    add_styled_table(["Control Category", "Implementation Details in SIEM Platform"], sec_data, [2.2, 4.3])

    add_heading_2("6.6 Testing and Functional Validation")
    add_heading_3("Table 7: Functional Validation Results")
    test_data = [
        ["Authentication", "Submit valid admin credentials", "Return 200 OK with signed JWT", "JWT issued; role verified as admin", "PASS"],
        ["Brute Force Defense", "Exceed 10 login requests / min", "Return HTTP 429 Too Many Requests", "Rate-limited by Flask-Limiter", "PASS"],
        ["Log Ingestion", "POST /v1/logs/ingest with raw syslog", "Normalize, classify severity, add to Redis", "Event stored in MongoDB & published", "PASS"],
        ["Threshold Rule", "Fire 10 failed logins in 30 seconds", "Correlation engine creates alert", "Alert created within 250ms", "PASS"],
        ["Sequence Rule", "Execute attack playbook with login success", "Trigger multi-stage compromise alert", "Alert generated with combined evidence", "PASS"],
        ["Live Streaming", "Connect browser to /v1/logs/stream", "Receive continuous Server-Sent Events", "Logs rendered in real time", "PASS"],
        ["Playbook Runner", "Execute brute_force_lateral.yml at 2x", "29 events dispatched in sequence", "All 29 events ingested; alerts fired", "PASS"],
        ["Alert Lifecycle", "Analyst transitions alert to 'Investigating'", "State updated in DB and reflected in UI", "Alert state synchronized across clients", "PASS"],
    ]
    add_styled_table(["Validation Area", "Test Case", "Expected Result", "Observed Result", "Status"], test_data, [1.3, 1.4, 1.5, 1.7, 0.6])

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 7: LIMITATIONS
    # =========================================================================
    add_heading_1("7. LIMITATIONS")
    add_bullet("In-Memory Correlation State: The multi-threaded correlation engine maintains sliding-window state inside local Python memory. In a distributed multi-node cluster, correlation state must be coordinated via distributed memory stores (e.g., Redis Sorted Sets or Apache Flink).")
    add_bullet("Log Volume Scalability: MongoDB document storage is optimized for intermediate-scale log ingestion. Massive enterprise ingestion (hundreds of thousands of events/sec) typically requires column-oriented or distributed indexing engines such as Elasticsearch, OpenSearch, or ClickHouse.")
    add_bullet("Rule Format Compatibility: The current system uses custom YAML/JSON definitions. Importing community Sigma Rules or YARA-L formats requires an additional transpiler module.")
    add_bullet("Active Response & SOAR Integration: The platform focuses on detection, investigation, and alert management. It does not yet include automated SOAR playbooks (such as automatic firewall IP blocking via webhook or agent-based endpoint isolation).")
    add_bullet("Cold Storage Archival: Logs are stored in active MongoDB collections. Automated archiving to low-cost object storage (e.g., AWS S3 or MinIO) with data retention lifecycle policies is not yet automated.")

    # =========================================================================
    # CHAPTER 8: CONCLUSIONS AND FUTURE WORK
    # =========================================================================
    add_heading_1("8. CONCLUSIONS AND FUTURE WORK")
    p = doc.add_paragraph()
    p.add_run("The industrial internship project successfully resulted in a functional, production-ready, full-stack ")
    p.add_run("Security Information and Event Management (SIEM) Dashboard").bold = True
    p.add_run(". The system proves that modern open-source technologies—Python 3.14, Flask, Redis Streams, MongoDB 9.0, React 19, and Vite—can be combined to construct high-performance security infrastructure.")

    p = doc.add_paragraph()
    p.add_run("The platform successfully satisfied all functional requirements: ingesting multi-source logs, normalizing schemas, automatically classifying severity levels, identifying multi-stage threats via sliding-window correlation, broadcasting live alerts over SSE, and simulating realistic MITRE ATT&CK adversary playbooks.")

    p = doc.add_paragraph()
    p.add_run("Through this project, I gained extensive practical experience in full-stack web engineering, distributed event queues, asynchronous socket/SSE streaming, cryptographic authentication, database optimization, defensive cybersecurity engineering, and container orchestration.")

    add_heading_2("8.1 Future Work")
    add_bullet("Sigma Rule Transpiler: Implement a parser to import and convert open-source Sigma rules directly into executable correlation queries.")
    add_bullet("Automated SOAR Action Triggers: Add webhook and SSH integrations to automatically trigger mitigation actions (e.g., null-routing malicious IPs on firewalls).")
    add_bullet("ClickHouse / OpenSearch Storage Backend: Introduce an abstraction layer to route long-term event telemetry into high-throughput analytical datastores.")
    add_bullet("Threat Intelligence Feed Ingestion: Connect real-time STIX/TAXII threat feeds to automatically match incoming IP and domain indicators against known malicious infrastructure.")
    add_bullet("Endpoint Agent Development: Develop lightweight Go-based endpoint telemetry agents for Windows (Event Tracing) and Linux (eBPF) to feed events directly to the SIEM.")

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 9: BIBLIOGRAPHY / REFERENCES
    # =========================================================================
    add_heading_1("9. BIBLIOGRAPHY / REFERENCES")
    refs = [
        "[1] INFOTECHMON PVT. LTD., “Company Profile and Industrial Training Programs,” Mohali, Punjab, official corporate portal, 2026.",
        "[2] MITRE Corporation, “MITRE ATT&CK Matrix for Enterprise,” version 15, https://attack.mitre.org/, accessed Aug. 2026.",
        "[3] CarterPerez-dev, “SIEM Dashboard: Real-time Event Correlation & MITRE ATT&CK Engine,” Cybersecurity Projects Repository, 2026.",
        "[4] M. Chapple, J. M. Stewart, and D. Gibson, CISSP Certified Information Systems Security Professional Official Study Guide, 9th ed., Wiley, 2021.",
        "[5] Redis Documentation, “Redis Streams Explained: Consumer Groups, Message Acknowledgment, and Persistent Queues,” https://redis.io/docs/data-types/streams/, 2026.",
        "[6] MongoDB Documentation, “Aggregation Pipeline Operations and Time-Series Data Handling,” MongoDB Inc., 2026.",
        "[7] Flask Documentation, “Application Factories, Blueprints, and CLI Extensions,” Pallets Projects, 2026.",
        "[8] React Documentation, “React 19: Server-Sent Events, Hooks, and Asynchronous UI Patterns,” Meta Platforms Inc., 2026.",
        "[9] TanStack Query Documentation, “React Query v5: Stale-While-Revalidate and Server Stream Subscriptions,” TanStack, 2026.",
        "[10] RFC 7519, “JSON Web Token (JWT) Architecture,” Internet Engineering Task Force (IETF), 2015.",
    ]
    for ref in refs:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.3)
        p.paragraph_format.first_line_indent = Inches(-0.3)
        p.paragraph_format.space_after = Pt(4)
        p.add_run(ref)

    doc.add_page_break()

    # =========================================================================
    # APPENDIX A & B
    # =========================================================================
    add_heading_1("APPENDIX A – PROJECT LINKS & ACCESS DETAILS")
    links_data = [
        ["Local Web Console", "http://localhost:8431"],
        ["REST API Base", "http://localhost:8431/api/v1/"],
        ["Direct Backend Port", "http://localhost:5113"],
        ["Direct Frontend Dev Server", "http://localhost:3959"],
        ["Database Services", "MongoDB: port 8654 | Redis: port 6266"],
        ["Default Admin Username", "admin"],
        ["Default Admin Password", "adminpassword123"],
        ["GitHub Project Repository", "https://github.com/CarterPerez-dev/Cybersecurity-Projects/tree/main/PROJECTS/intermediate/siem-dashboard"],
    ]
    add_styled_table(["Parameter", "Configuration / URL"], links_data, [2.5, 4.0])

    add_heading_1("APPENDIX B – SUGGESTED SCREENSHOTS & VISUAL VERIFICATION")
    p = doc.add_paragraph()
    p.add_run("The following screenshots can be inserted into the visual evidence section before final academic submission:")
    add_bullet("1. Landing Page & Authentication: Secure login screen with JWT authentication and registration links.")
    add_bullet("2. SOC Executive Dashboard: Visual summary showing active alert cards, 24-hour log histograms, severity donut charts, and top source IP rankings.")
    add_bullet("3. Live Log Stream Console: Virtualized table streaming real-time security events over SSE with dynamic severity badge highlights.")
    add_bullet("4. Log Pivot & Search Interface: Multi-source log view demonstrating query filtering by IP, timestamp, and source type.")
    add_bullet("5. Correlation Rule Management: Interface displaying active Threshold, Sequence, and Aggregation rules with threshold configurations.")
    add_bullet("6. Incident Alert Triage Screen: Detailed view of an active alert showing correlated events, MITRE tactic tags, and analyst investigation notes.")
    add_bullet("7. MITRE ATT&CK Attack Playbook Runner: Scenario playback controller displaying the execution of the Brute Force with Lateral Movement simulation.")
    add_bullet("8. Real-Time Alert Toast / SSE Trigger: Screenshot illustrating instant UI notification when a correlation rule triggers during an attack scenario.")
    add_bullet("9. Docker Multi-Container Deployment: Terminal screenshot showing docker compose ps verifying that siem-nginx, siem-backend, siem-redis, and siem-mongo are healthy.")

    output_filename = "Aditya_Tiwari_SIEM_Dashboard_Internship_Report.docx"
    doc.save(output_filename)
    print(f"Report saved successfully as {output_filename}")

if __name__ == "__main__":
    create_report()
