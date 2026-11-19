import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def create_deck():
    prs = Presentation()
    # 16:9 Widescreen dimensions
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6] # Blank slide layout

    # Color Palette - Professional SOC Cybersecurity Theme
    BG_DARK = RGBColor(11, 25, 44)         # #0B192C (Deep Midnight Navy)
    BG_CARD = RGBColor(26, 43, 76)         # #1A2B4C (Card Container)
    BG_CARD_LIGHT = RGBColor(35, 58, 98)   # #233A62
    ACCENT_CYAN = RGBColor(0, 210, 211)    # #00D2D3 (Bright Cyan / Electric)
    ACCENT_BLUE = RGBColor(56, 189, 248)   # #38BDF8 (Sky Blue)
    ACCENT_RED = RGBColor(239, 68, 68)     # #EF4444 (Critical / Threat Red)
    ACCENT_AMBER = RGBColor(245, 158, 11)  # #F59E0B (Warning Amber)
    ACCENT_GREEN = RGBColor(16, 185, 129)  # #10B981 (Success Green)
    TEXT_WHITE = RGBColor(255, 255, 255)
    TEXT_MUTED = RGBColor(186, 199, 213)   # Soft silver-gray
    LINE_BORDER = RGBColor(45, 75, 120)

    def add_bg(slide):
        # Background rectangle covering full slide
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_DARK
        bg.line.color.rgb = BG_DARK
        return bg

    def add_header(slide, title_text, category="SIEM DASHBOARD | INTERNSHIP PROJECT DEFENSE"):
        # Header banner category
        tb_cat = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.35))
        tf_cat = tb_cat.text_frame
        tf_cat.word_wrap = True
        tf_cat.margin_left = tf_cat.margin_top = tf_cat.margin_right = tf_cat.margin_bottom = 0
        p_cat = tf_cat.paragraphs[0]
        p_cat.text = category.upper()
        p_cat.font.name = "Arial"
        p_cat.font.size = Pt(10)
        p_cat.font.bold = True
        p_cat.font.color.rgb = ACCENT_CYAN

        # Main slide title
        tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.75), Inches(11.7), Inches(0.65))
        tf_title = tb_title.text_frame
        tf_title.word_wrap = True
        tf_title.margin_left = tf_title.margin_top = tf_title.margin_right = tf_title.margin_bottom = 0
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.font.name = "Arial"
        p_title.font.size = Pt(22)
        p_title.font.bold = True
        p_title.font.color.rgb = TEXT_WHITE

        # Subtle divider line
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.45), Inches(11.7), Inches(0.02))
        line.fill.solid()
        line.fill.fore_color.rgb = LINE_BORDER
        line.line.color.rgb = LINE_BORDER

    def add_card(slide, left, top, width, height, bg_color=BG_CARD, border_color=LINE_BORDER):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1)
        return card

    # =========================================================================
    # SLIDE 1: TITLE SLIDE (COVER)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    add_bg(s1)

    # Accent top border strip
    strip = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.12))
    strip.fill.solid()
    strip.fill.fore_color.rgb = ACCENT_CYAN
    strip.line.color.rgb = ACCENT_CYAN

    # Main Card Box
    add_card(s1, Inches(1.0), Inches(1.0), Inches(11.333), Inches(5.5), bg_color=RGBColor(16, 32, 58), border_color=ACCENT_CYAN)

    tb = s1.shapes.add_textbox(Inches(1.5), Inches(1.3), Inches(10.333), Inches(1.5))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = "INTERNSHIP PROJECT DEFENSE & TECHNICAL PRESENTATION"
    p.font.name = "Arial"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = ACCENT_CYAN

    p2 = tf.add_paragraph()
    p2.text = "SIEM DASHBOARD"
    p2.font.name = "Arial"
    p2.font.size = Pt(36)
    p2.font.bold = True
    p2.font.color.rgb = TEXT_WHITE
    p2.space_before = Pt(4)

    p3 = tf.add_paragraph()
    p3.text = "Real-Time Security Information & Event Management with Multi-Format Log Normalization, Sliding-Window Event Correlation, and MITRE ATT&CK Attack Simulation Engine"
    p3.font.name = "Calibri"
    p3.font.size = Pt(14)
    p3.font.color.rgb = ACCENT_BLUE
    p3.space_before = Pt(6)

    # Candidate & Mentor Details (Two Cards)
    # Left Card: Candidate
    c1 = add_card(s1, Inches(1.5), Inches(3.6), Inches(4.9), Inches(2.4), bg_color=BG_CARD)
    tb_c1 = s1.shapes.add_textbox(Inches(1.7), Inches(3.75), Inches(4.5), Inches(2.1))
    tf_c1 = tb_c1.text_frame
    tf_c1.word_wrap = True
    
    p = tf_c1.paragraphs[0]
    p.text = "SUBMITTED BY"
    p.font.bold = True
    p.font.size = Pt(10)
    p.font.color.rgb = ACCENT_CYAN

    p = tf_c1.add_paragraph()
    p.text = "ADITYA TIWARI"
    p.font.bold = True
    p.font.size = Pt(16)
    p.font.color.rgb = TEXT_WHITE
    p.space_before = Pt(4)

    p = tf_c1.add_paragraph()
    p.text = "Roll No: 72413167 | Reg No: CTU2401785\nB.Tech - Computer Science and Engineering\nCT University, Ludhiana, Punjab"
    p.font.size = Pt(11)
    p.font.color.rgb = TEXT_MUTED
    p.space_before = Pt(4)

    # Right Card: Mentors & Organization
    c2 = add_card(s1, Inches(6.8), Inches(3.6), Inches(5.0), Inches(2.4), bg_color=BG_CARD)
    tb_c2 = s1.shapes.add_textbox(Inches(7.0), Inches(3.75), Inches(4.6), Inches(2.1))
    tf_c2 = tb_c2.text_frame
    tf_c2.word_wrap = True
    
    p = tf_c2.paragraphs[0]
    p.text = "UNDER THE GUIDANCE OF"
    p.font.bold = True
    p.font.size = Pt(10)
    p.font.color.rgb = ACCENT_CYAN

    p = tf_c2.add_paragraph()
    p.text = "Abhishek Jayant, Industrial Mentor"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = TEXT_WHITE
    p.space_before = Pt(4)

    p = tf_c2.add_paragraph()
    p.text = "INFOTECHMON PVT. LTD., Mohali\n"
    p.font.size = Pt(11)
    p.font.color.rgb = TEXT_MUTED

    p = tf_c2.add_paragraph()
    p.text = "Arvin Vinayek, Associate Professor"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = TEXT_WHITE
    p.space_before = Pt(4)

    p = tf_c2.add_paragraph()
    p.text = "Department of CSE, CT University, Ludhiana"
    p.font.size = Pt(11)
    p.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 2: EXECUTIVE SUMMARY & THREAT LANDSCAPE
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_bg(s2)
    add_header(s2, "Executive Summary: The Modern SOC Challenge", "PROJECT BACKGROUND & INDUSTRY NEED")

    # 3 Column Pillar Cards
    pillars = [
        ("The Enterprise Dilemma", ACCENT_RED, [
            "Thousands of telemetry logs per second across firewalls, auth servers, and endpoints.",
            "Data Isolation: Disparate log formats prevent cross-source correlation.",
            "Alert Fatigue: Floods of raw unclassified logs overwhelm SOC analysts.",
            "Commercial SIEMs (Splunk, Sentinel) introduce prohibitive per-gigabyte costs."
        ]),
        ("The SIEM Dashboard Solution", ACCENT_CYAN, [
            "Open, self-hosted, full-stack defensive monitoring platform.",
            "Multi-format ingestion normalizing 6 log source types into a unified schema.",
            "Threaded sliding-window engine for real-time complex threat correlation.",
            "Push-based Server-Sent Events (SSE) for sub-second alert streaming."
        ]),
        ("MITRE ATT&CK Playbook Engine", ACCENT_GREEN, [
            "Embedded adversary simulator executing real-world TTP playbooks.",
            "Replays multi-stage attacks: SSH brute force, lateral pivot, DNS exfiltration.",
            "Enables safe detection engineering benchmarking without risk to production.",
            "Complete alert lifecycle management: Claim, Investigate, Resolve, False Positive."
        ])
    ]

    for i, (title, accent, bullets) in enumerate(pillars):
        left = Inches(0.8 + i * 3.95)
        add_card(s2, left, Inches(1.8), Inches(3.8), Inches(5.0))
        
        # Accent top bar for card
        bar = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, Inches(1.8), Inches(3.8), Inches(0.08))
        bar.fill.solid()
        bar.fill.fore_color.rgb = accent
        bar.line.color.rgb = accent

        tb = s2.shapes.add_textbox(left + Inches(0.25), Inches(2.05), Inches(3.3), Inches(4.5))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = "Arial"
        p.font.size = Pt(15)
        p.font.bold = True
        p.font.color.rgb = TEXT_WHITE

        for b in bullets:
            p = tf.add_paragraph()
            p.text = "• " + b
            p.font.name = "Calibri"
            p.font.size = Pt(11)
            p.font.color.rgb = TEXT_MUTED
            p.space_before = Pt(8)
            p.line_spacing = 1.15

    # =========================================================================
    # SLIDE 3: PROBLEM STATEMENT & CORE OBJECTIVES
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_bg(s3)
    add_header(s3, "Problem Statement & Project Objectives", "ENGINEERING GOALS")

    # Left: Problem Statement Card
    add_card(s3, Inches(0.8), Inches(1.8), Inches(5.2), Inches(5.0))
    tb_prob = s3.shapes.add_textbox(Inches(1.05), Inches(2.0), Inches(4.7), Inches(4.6))
    tf_prob = tb_prob.text_frame
    tf_prob.word_wrap = True
    
    p = tf_prob.paragraphs[0]
    p.text = "PROBLEM STATEMENT"
    p.font.name = "Arial"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = ACCENT_RED

    problems = [
        "Vendor Format Fragmentation: Conflicting field names (e.g., c_ip vs remote_addr vs src_ip) prevent unified rule creation.",
        "Missed Multi-Stage Attacks: Isolated log searches fail to correlate separate events, such as failed logins followed by lateral movement.",
        "Excessive Alert Noise: Lack of heuristic severity classification overwhelms security responders.",
        "Inability to Test Detections: Junior analysts lack safe environments to trigger realistic adversary tactics against detection rules."
    ]
    for prob in problems:
        p = tf_prob.add_paragraph()
        p.text = "▸ " + prob
        p.font.name = "Calibri"
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_MUTED
        p.space_before = Pt(8)
        p.line_spacing = 1.15

    # Right: Project Objectives Cards
    objectives = [
        ("Multi-Source Ingestion & Normalization", "Harmonize Firewall, IDS, Auth, Endpoint, DNS, and Proxy logs into a normalized LogEvent model."),
        ("Heuristic Severity Scoring", "Pattern-match signatures to automatically assign Low, Medium, High, and Critical severities."),
        ("Sliding-Window Correlation Engine", "Implement multi-threaded Threshold, Sequence, and Aggregation rule evaluation in real time."),
        ("Asynchronous Push Streaming", "Leverage Redis Streams and Server-Sent Events (SSE) for live browser telemetry without polling."),
        ("MITRE ATT&CK Playbook Engine", "Develop a threaded YAML-based attack scenario runner with variable speed multipliers (0.1x to 10x).")
    ]

    for j, (obj_title, obj_desc) in enumerate(objectives):
        top_pos = Inches(1.8 + j * 0.98)
        c = add_card(s3, Inches(6.3), top_pos, Inches(6.2), Inches(0.88), bg_color=BG_CARD_LIGHT)
        
        # Left tag bar
        tag = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.3), top_pos, Inches(0.08), Inches(0.88))
        tag.fill.solid()
        tag.fill.fore_color.rgb = ACCENT_CYAN
        tag.line.color.rgb = ACCENT_CYAN

        tb_obj = s3.shapes.add_textbox(Inches(6.5), top_pos + Inches(0.08), Inches(5.9), Inches(0.72))
        tf_obj = tb_obj.text_frame
        tf_obj.word_wrap = True
        
        p = tf_obj.paragraphs[0]
        p.text = f"{j+1}. {obj_title}"
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = TEXT_WHITE

        p2 = tf_obj.add_paragraph()
        p2.text = obj_desc
        p2.font.name = "Calibri"
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = TEXT_MUTED
        p2.space_before = Pt(1)

    # =========================================================================
    # SLIDE 4: SYSTEM ARCHITECTURE
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_bg(s4)
    add_header(s4, "End-to-End System Architecture", "SYSTEM DESIGN & DATA FLOW")

    # 4 Architecture Horizontal Layer Cards
    layers = [
        ("1. Presentation Layer (SOC Console)", ACCENT_CYAN, [
            "React 19 & TypeScript frontend built with Vite (Rolldown) for fast rendering.",
            "Zustand global stores managing authentication state, streaming connections, and active UI filters.",
            "TanStack Query v5 providing optimistic caching, mutations, and automatic background refetching.",
            "visx (@visx/xychart) low-level SVG/Canvas visualizations for 24-hour log volume and threat distributions."
        ]),
        ("2. Gateway & Application Layer (Flask + Nginx)", ACCENT_BLUE, [
            "Nginx 1.29 edge reverse proxy with special buffering bypass (`proxy_buffering off`) for continuous SSE streams.",
            "Flask 3.1 application factory in Python 3.14 with structured blueprints and error hierarchies.",
            "Pydantic v2 validation ensuring strict request/response schema parsing and fast execution.",
            "Gunicorn WSGI multi-worker server orchestrating concurrent RESTful API endpoints."
        ]),
        ("3. Telemetry Stream & Correlation Engine", ACCENT_AMBER, [
            "Redis 7 Streams (`siem:logs`, `siem:alerts`) acting as high-throughput distributed event message queues.",
            "Consumer Groups (`siem-correlation`): Engine thread executes `XREADGROUP` to process events reliably.",
            "Threaded sliding-window daemon evaluating Threshold, Sequence, and Aggregation rules concurrently.",
            "Automatic Alert Generation: Dispatches alerts to MongoDB and broadcasts instant SSE notifications."
        ]),
        ("4. Persistent Storage & Identity Layer", ACCENT_GREEN, [
            "MongoDB 9.0 document datastore with indexes for fast IP pivot queries and aggregation pipelines.",
            "MongoEngine ORM managing LogEvent, Alert, CorrelationRule, ScenarioRun, and User models.",
            "Argon2id password hashing and HS256-signed JWT tokens enforcing Analyst vs Admin authorization."
        ])
    ]

    for k, (l_title, l_accent, l_bullets) in enumerate(layers):
        top_k = Inches(1.75 + k * 1.3)
        add_card(s4, Inches(0.8), top_k, Inches(11.7), Inches(1.18))
        
        # Left accent block
        l_bar = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), top_k, Inches(0.12), Inches(1.18))
        l_bar.fill.solid()
        l_bar.fill.fore_color.rgb = l_accent
        l_bar.line.color.rgb = l_accent

        tb_lay = s4.shapes.add_textbox(Inches(1.1), top_k + Inches(0.08), Inches(11.2), Inches(1.0))
        tf_lay = tb_lay.text_frame
        tf_lay.word_wrap = True
        
        p = tf_lay.paragraphs[0]
        p.text = l_title
        p.font.name = "Arial"
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = l_accent

        for b in l_bullets:
            p = tf_lay.add_paragraph()
            p.text = "• " + b
            p.font.name = "Calibri"
            p.font.size = Pt(9.5)
            p.font.color.rgb = TEXT_WHITE
            p.space_before = Pt(1)

    # =========================================================================
    # SLIDE 5: MULTI-SOURCE INGESTION & NORMALIZATION
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_bg(s5)
    add_header(s5, "Multi-Format Log Normalization Pipeline", "DATA INGESTION & TELEMETRY STANDARDIZATION")

    # Ingestion steps card (Top half)
    add_card(s5, Inches(0.8), Inches(1.8), Inches(11.7), Inches(1.8))
    tb_steps = s5.shapes.add_textbox(Inches(1.0), Inches(1.9), Inches(11.3), Inches(1.5))
    tf_steps = tb_steps.text_frame
    tf_steps.word_wrap = True

    p = tf_steps.paragraphs[0]
    p.text = "THREE-STAGE INGESTION & ENRICHMENT PIPELINE"
    p.font.name = "Arial"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = ACCENT_CYAN

    stages = [
        "1. Schema Parsing & Validation: Pydantic schemas validate JSON/Syslog fields, datatypes, and timestamps at /v1/logs/ingest.",
        "2. Source-Specific Harmonization: Dedicated normalizer transforms divergent source fields into standard LogEvent schema.",
        "3. Heuristic Severity Scoring: Pattern matching assigns severity (LOW, MEDIUM, HIGH, CRITICAL) based on regex attack signatures.",
        "4. Dual Dispatch: Normalized event is synchronously stored in MongoDB and published to Redis Stream ('siem:logs') via XADD."
    ]
    for stg in stages:
        p = tf_steps.add_paragraph()
        p.text = stg
        p.font.name = "Calibri"
        p.font.size = Pt(10.5)
        p.font.color.rgb = TEXT_WHITE
        p.space_before = Pt(3)

    # Normalization Comparison Table (Bottom half)
    table_shape = s5.shapes.add_table(5, 5, Inches(0.8), Inches(3.8), Inches(11.7), Inches(2.9))
    table = table_shape.table

    norm_headers = ["Normalized Key", "Firewall Source", "Authentication Source", "DNS Source", "Endpoint Source"]
    norm_rows = [
        ["source_ip", "src_ip / c_ip", "client_ip / host", "client", "network.src"],
        ["destination_ip", "dst_ip / s_ip", "server_ip", "resolver", "network.dst"],
        ["username", "N/A", "user / login_id", "N/A", "user.name / account"],
        ["event_type", "action (DENY/ALLOW)", "status (FAILED/SUCCESS)", "query_type (TXT/A)", "process.name"]
    ]

    # Format Table Headers
    for c_idx, h_text in enumerate(norm_headers):
        cell = table.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(16, 32, 58)
        p = cell.text_frame.paragraphs[0]
        p.text = h_text
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = ACCENT_CYAN

    # Format Table Rows
    for r_idx, row in enumerate(norm_rows):
        for c_idx, val in enumerate(row):
            cell = table.cell(r_idx + 1, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = BG_CARD if r_idx % 2 == 0 else BG_CARD_LIGHT
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = "Calibri"
            p.font.size = Pt(10)
            p.font.color.rgb = TEXT_WHITE

    # =========================================================================
    # SLIDE 6: CORRELATION RULE ENGINE
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_bg(s6)
    add_header(s6, "Sliding-Window Correlation Engine", "REAL-TIME DETECTION LOGIC")

    rules_info = [
        ("Threshold Rules", ACCENT_AMBER, "Frequency Burst Detection", [
            "Evaluates occurrence count of identical event types from the same origin.",
            "Example: > 5 failed logins within 60 seconds from an identical IP.",
            "Sliding-window counter tracks state in-memory with automatic expiry.",
            "Generates High severity 'Brute Force Attempt' alert with full audit trail."
        ]),
        ("Sequence Rules", ACCENT_RED, "Multi-Stage Attack Correlation", [
            "Detects ordered progression across multiple attack stages.",
            "Stage 1: Repeated authentication failures from IP 'X'.",
            "Stage 2: Successful login from identical IP 'X' within 120 seconds.",
            "Immediately raises Critical severity 'Account Compromise Following Brute Force'."
        ]),
        ("Aggregation Rules", ACCENT_CYAN, "Cardinality & Anomaly Grouping", [
            "Identifies behavioral anomalies based on distinct field cardinality.",
            "Example: Single client querying > 50 distinct subdomains via DNS in 3 minutes.",
            "Detects data exfiltration and C2 tunneling (e.g., OilRig / DNSpionage).",
            "Triggers automated incident alert and flags IP for analyst pivot investigation."
        ])
    ]

    for m, (r_name, r_color, r_sub, r_bullets) in enumerate(rules_info):
        left_m = Inches(0.8 + m * 3.95)
        add_card(s6, left_m, Inches(1.8), Inches(3.8), Inches(5.0))
        
        bar_m = s6.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_m, Inches(1.8), Inches(3.8), Inches(0.08))
        bar_m.fill.solid()
        bar_m.fill.fore_color.rgb = r_color
        bar_m.line.color.rgb = r_color

        tb_m = s6.shapes.add_textbox(left_m + Inches(0.2), Inches(2.0), Inches(3.4), Inches(4.5))
        tf_m = tb_m.text_frame
        tf_m.word_wrap = True
        
        p = tf_m.paragraphs[0]
        p.text = r_name
        p.font.name = "Arial"
        p.font.size = Pt(15)
        p.font.bold = True
        p.font.color.rgb = TEXT_WHITE

        p_sub = tf_m.add_paragraph()
        p_sub.text = r_sub
        p_sub.font.name = "Calibri"
        p_sub.font.size = Pt(10)
        p_sub.font.color.rgb = r_color
        p_sub.space_before = Pt(2)

        for b in r_bullets:
            p = tf_m.add_paragraph()
            p.text = "• " + b
            p.font.name = "Calibri"
            p.font.size = Pt(10)
            p.font.color.rgb = TEXT_MUTED
            p.space_before = Pt(6)
            p.line_spacing = 1.15

    # =========================================================================
    # SLIDE 7: MITRE ATT&CK PLAYBOOK SIMULATION ENGINE
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_bg(s7)
    add_header(s7, "MITRE ATT&CK Attack Scenario Simulation", "THREAT PLAYBOOK EMULATION ENGINE")

    # Simulation engine overview card
    add_card(s7, Inches(0.8), Inches(1.75), Inches(11.7), Inches(1.4))
    tb_sim = s7.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(11.3), Inches(1.2))
    tf_sim = tb_sim.text_frame
    tf_sim.word_wrap = True
    
    p = tf_sim.paragraphs[0]
    p.text = "PURPOSE & THREADED PLAYBOOK RUNNER ARCHITECTURE"
    p.font.bold = True
    p.font.size = Pt(11)
    p.font.color.rgb = ACCENT_CYAN

    p = tf_sim.add_paragraph()
    p.text = "• Detection Rule Verification: Allows security teams to validate correlation rules against realistic multi-stage attacks safely.\n• Threaded Scenario Runner: Dispatches staged events into the ingestion API with controllable playback speed (0.1x to 10.0x).\n• Dynamic Playbook Parsing: Scenarios defined in YAML specify step-by-step payloads, MITRE tactics, timestamps, and target hosts."
    p.font.size = Pt(10)
    p.font.color.rgb = TEXT_WHITE
    p.space_before = Pt(3)

    # 4 Playbook Cards
    playbooks = [
        ("1. Brute Force & Lateral Movement", "TA0001 / TA0008", "T1110.001 (Password Guessing)\nT1021.004 (SSH Pivot)", "Simulates SSH credential brute-forcing against external jump host, followed by lateral pivot to internal production database."),
        ("2. DNS Tunneling Data Exfiltration", "TA0007 / TA0010", "T1046 (Network Service Discovery)\nT1048.003 (DNS Exfiltration)", "Internal host port reconnaissance, staging sensitive customer records, and exfiltrating data via high-entropy DNS TXT queries."),
        ("3. Phishing to C2 Beaconing", "TA0001 / TA0011", "T1566.001 (Spearphishing Attachment)\nT1071.001 (HTTP C2 Web Protocols)", "User opens malicious macro attachment, executing PowerShell download cradle and establishing continuous HTTP C2 beaconing."),
        ("4. Privilege Escalation & Persistence", "TA0004 / TA0003", "T1068 (Kernel Exploitation)\nT1136.001 (Backdoor Account Creation)", "Exploitation of kernel vulnerability for root, creation of rogue admin user, cron persistence, and log wiping.")
    ]

    for n, (p_title, p_tactics, p_tech, p_desc) in enumerate(playbooks):
        col = n % 2
        row = n // 2
        l_pos = Inches(0.8 + col * 5.95)
        t_pos = Inches(3.35 + row * 1.85)

        add_card(s7, l_pos, t_pos, Inches(5.75), Inches(1.7), bg_color=BG_CARD)
        
        tb_pb = s7.shapes.add_textbox(l_pos + Inches(0.2), t_pos + Inches(0.1), Inches(5.35), Inches(1.5))
        tf_pb = tb_pb.text_frame
        tf_pb.word_wrap = True

        p = tf_pb.paragraphs[0]
        p.text = p_title
        p.font.name = "Arial"
        p.font.size = Pt(11.5)
        p.font.bold = True
        p.font.color.rgb = TEXT_WHITE

        p = tf_pb.add_paragraph()
        p.text = f"Tactics: {p_tactics} | Techniques: {p_tech.replace(chr(10), ', ')}"
        p.font.size = Pt(9)
        p.font.color.rgb = ACCENT_AMBER
        p.space_before = Pt(2)

        p = tf_pb.add_paragraph()
        p.text = p_desc
        p.font.size = Pt(9.5)
        p.font.color.rgb = TEXT_MUTED
        p.space_before = Pt(3)

    # =========================================================================
    # SLIDE 8: REAL-TIME STREAMING & SSE ARCHITECTURE
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_bg(s8)
    add_header(s8, "Real-Time Telemetry & SSE Streaming Architecture", "ASYNCHRONOUS PUSH MESSAGING")

    # Left Card: Traditional Polling vs Push
    add_card(s8, Inches(0.8), Inches(1.8), Inches(5.6), Inches(5.0))
    tb_poll = s8.shapes.add_textbox(Inches(1.05), Inches(2.0), Inches(5.1), Inches(4.5))
    tf_poll = tb_poll.text_frame
    tf_poll.word_wrap = True

    p = tf_poll.paragraphs[0]
    p.text = "POLLING VS ASYNCHRONOUS SSE PUSH"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = ACCENT_CYAN

    diffs = [
        ("The Polling Bottleneck:", "Traditional SIEM dashboards query the database every 5-10 seconds. This creates server query spikes, high database CPU usage, and high alert latency."),
        ("Server-Sent Events (SSE) Solution:", "Unidirectional HTTP streaming connection kept open between Nginx and browser. Telemetry pushed immediately as events occur."),
        ("Sub-Second Latency:", "Alerts generated by the correlation engine reach the analyst's screen in less than 250 milliseconds."),
        ("Zero Overhead on Ingestion:", "Browser connections subscribe to Redis Streams; database write operations remain completely decoupled from stream reads.")
    ]
    for h, d in diffs:
        p = tf_poll.add_paragraph()
        p.text = h
        p.font.bold = True
        p.font.size = Pt(10.5)
        p.font.color.rgb = TEXT_WHITE
        p.space_before = Pt(6)
        
        p2 = tf_poll.add_paragraph()
        p2.text = d
        p2.font.size = Pt(10)
        p2.font.color.rgb = TEXT_MUTED

    # Right Card: Stream Queues & Consumer Groups
    add_card(s8, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.0))
    tb_red = s8.shapes.add_textbox(Inches(7.05), Inches(2.0), Inches(5.2), Inches(4.5))
    tf_red = tb_red.text_frame
    tf_red.word_wrap = True

    p = tf_red.paragraphs[0]
    p.text = "REDIS STREAMS & NGINX PROXY TUNING"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = ACCENT_BLUE

    stream_points = [
        ("Redis Stream Queues:", "`siem:logs` buffers raw ingested telemetry.\n`siem:alerts` buffers newly generated correlation alerts."),
        ("Consumer Groups (`siem-correlation`):", "Enables reliable message consumption. Consumer `engine-1` reads in batches (`XREADGROUP`) with offset tracking to ensure zero lost events."),
        ("Nginx Proxy Buffering Bypass:", "Nginx defaults to buffering upstream HTTP responses, which breaks real-time SSE streaming."),
        ("Configured Proxy Directives:", "`proxy_buffering off;`\n`proxy_cache off;`\n`proxy_read_timeout 3600s;`\n`chunked_transfer_encoding off;`")
    ]
    for h, d in stream_points:
        p = tf_red.add_paragraph()
        p.text = h
        p.font.bold = True
        p.font.size = Pt(10.5)
        p.font.color.rgb = TEXT_WHITE
        p.space_before = Pt(6)
        
        p2 = tf_red.add_paragraph()
        p2.text = d
        p2.font.size = Pt(10)
        p2.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 9: SECURITY ENGINEERING & CONTROLS
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_bg(s9)
    add_header(s9, "Security Engineering & Defensive Hardening", "AUTHENTICATION, RATE LIMITING & ACCESS CONTROL")

    sec_cards = [
        ("Argon2id Password Security", ACCENT_CYAN, [
            "State-of-the-art memory-hard hashing via pwdlib[argon2].",
            "Protects credentials against GPU and ASIC cracking attacks.",
            "Per-user unique cryptographic salt generated automatically."
        ]),
        ("Anti-Timing Attack Protection", ACCENT_BLUE, [
            "Constant-time authentication verification.",
            "If user is not found, dummy Argon2 hash verification runs anyway.",
            "Completely eliminates side-channel user enumeration timing attacks."
        ]),
        ("JWT Role-Based Access Control", ACCENT_GREEN, [
            "HS256 signed JSON Web Tokens carrying UID, expiry, and user role.",
            "Analyst Role: Live stream monitoring, log search, alert triage.",
            "Admin Role: Correlation rule creation, user management, playbook runs."
        ]),
        ("Moving-Window Rate Limiting", ACCENT_RED, [
            "Powered by Flask-Limiter with Redis backend storage.",
            "Auth Endpoints (/v1/auth/*): Strict 10 req/min limit prevents brute-force.",
            "Telemetry & API Endpoints: 200 req/min general protection."
        ])
    ]

    for idx, (title, color, bullets) in enumerate(sec_cards):
        x = Inches(0.8 + (idx % 2) * 5.95)
        y = Inches(1.8 + (idx // 2) * 2.55)

        add_card(s9, x, y, Inches(5.75), Inches(2.35), bg_color=BG_CARD)
        
        tag = s9.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, Inches(5.75), Inches(0.06))
        tag.fill.solid()
        tag.fill.fore_color.rgb = color
        tag.line.color.rgb = color

        tb = s9.shapes.add_textbox(x + Inches(0.25), y + Inches(0.15), Inches(5.25), Inches(2.0))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = title
        p.font.name = "Arial"
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = TEXT_WHITE

        for b in bullets:
            p = tf.add_paragraph()
            p.text = "• " + b
            p.font.name = "Calibri"
            p.font.size = Pt(10)
            p.font.color.rgb = TEXT_MUTED
            p.space_before = Pt(4)

    # =========================================================================
    # SLIDE 10: COMPLETE TECHNOLOGY STACK
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_bg(s10)
    add_header(s10, "Technology Stack Specifications", "FULL-STACK TOOLS & FRAMEWORKS")

    t_shape = s10.shapes.add_table(8, 4, Inches(0.8), Inches(1.8), Inches(11.7), Inches(5.0))
    t = t_shape.table

    stack_headers = ["Layer", "Technology", "Version", "Key Function & Responsibility"]
    stack_data = [
        ["Frontend UI", "React & TypeScript", "19.3 / 5.9", "Reactive analyst console, component state, and static type safety"],
        ["State & Caching", "Zustand & TanStack Query", "5.0 / 5.105", "Client session stores, real-time stream state, and API query caching"],
        ["Visualizations", "visx (@visx/xychart)", "3.12", "Low-level responsive charts for 24h event volume & severity breakdown"],
        ["Backend API", "Python & Flask", "3.14 / 3.1", "Micro-framework application factory with Pydantic request validation"],
        ["Streaming Broker", "Redis 7 Streams", "7.4", "Distributed event bus (`siem:logs`, `siem:alerts`) with consumer groups"],
        ["Database", "MongoDB", "9.0", "Persistent document datastore for LogEvents, Alerts, Rules, and Users"],
        ["Reverse Proxy", "Nginx", "1.29-alpine", "Edge routing, SSL termination, and SSE proxy buffer disablement"]
    ]

    for col_i, h_txt in enumerate(stack_headers):
        c = t.cell(0, col_i)
        c.fill.solid()
        c.fill.fore_color.rgb = RGBColor(16, 32, 58)
        p = c.text_frame.paragraphs[0]
        p.text = h_txt
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = ACCENT_CYAN

    for row_i, row_vals in enumerate(stack_data):
        for col_i, val in enumerate(row_vals):
            c = t.cell(row_i + 1, col_i)
            c.fill.solid()
            c.fill.fore_color.rgb = BG_CARD if row_i % 2 == 0 else BG_CARD_LIGHT
            p = c.text_frame.paragraphs[0]
            p.text = val
            p.font.name = "Calibri"
            p.font.size = Pt(10)
            p.font.color.rgb = TEXT_WHITE

    # =========================================================================
    # SLIDE 11: FUNCTIONAL VALIDATION & TEST RESULTS
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    add_bg(s11)
    add_header(s11, "Empirical Validation & Test Results", "SYSTEM VERIFICATION & BENCHMARKING")

    v_shape = s11.shapes.add_table(8, 5, Inches(0.8), Inches(1.8), Inches(11.7), Inches(5.0))
    vt = v_shape.table

    val_headers = ["Validation Area", "Test Action Performed", "Expected Outcome", "Observed Result", "Status"]
    val_data = [
        ["Authentication", "Submit valid admin credentials", "200 OK + Signed JWT Token", "JWT issued; role verified as admin", "PASS"],
        ["Rate Limiting", "Exceed 10 login requests / min", "HTTP 429 Too Many Requests", "Rate limit triggered by Flask-Limiter", "PASS"],
        ["Log Ingestion", "POST /v1/logs/ingest raw syslog", "Normalize, classify, push to Redis", "Event stored in Mongo & published", "PASS"],
        ["Threshold Rule", "Send 10 failed logins in 30 sec", "Trigger Brute Force Alert", "Alert generated in < 250ms", "PASS"],
        ["Sequence Rule", "Run playbook: failure -> success", "Detect Account Compromise", "Alert raised with combined evidence", "PASS"],
        ["SSE Streaming", "Connect browser to /v1/logs/stream", "Stream events continuously", "Zero dropped events; instant UI render", "PASS"],
        ["Playbook Runner", "Execute brute_force_lateral.yml at 2x", "Dispatch 29 staged attack events", "All 29 events processed; alerts fired", "PASS"]
    ]

    for col_i, h_txt in enumerate(val_headers):
        c = vt.cell(0, col_i)
        c.fill.solid()
        c.fill.fore_color.rgb = RGBColor(16, 32, 58)
        p = c.text_frame.paragraphs[0]
        p.text = h_txt
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = ACCENT_CYAN

    for row_i, row_vals in enumerate(val_data):
        for col_i, val in enumerate(row_vals):
            c = vt.cell(row_i + 1, col_i)
            c.fill.solid()
            c.fill.fore_color.rgb = BG_CARD if row_i % 2 == 0 else BG_CARD_LIGHT
            p = c.text_frame.paragraphs[0]
            p.text = val
            p.font.name = "Calibri"
            p.font.size = Pt(10)
            if col_i == 4: # Status column
                p.font.bold = True
                p.font.color.rgb = ACCENT_GREEN
            else:
                p.font.color.rgb = TEXT_WHITE

    # =========================================================================
    # SLIDE 12: LIMITATIONS & FUTURE ROADMAP
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    add_bg(s12)
    add_header(s12, "Project Limitations & Future Enhancements", "ENGINEERING BOUNDARIES & ROADMAP")

    # Left: Limitations Card
    add_card(s12, Inches(0.8), Inches(1.8), Inches(5.6), Inches(5.0))
    tb_lim = s12.shapes.add_textbox(Inches(1.05), Inches(2.0), Inches(5.1), Inches(4.5))
    tf_lim = tb_lim.text_frame
    tf_lim.word_wrap = True

    p = tf_lim.paragraphs[0]
    p.text = "CURRENT SYSTEM BOUNDARIES"
    p.font.name = "Arial"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = ACCENT_AMBER

    limits = [
        ("In-Memory Sliding Window State:", "Correlation state is tracked in local process memory. Multi-node clusters require distributed state management via Redis Sorted Sets or Apache Flink."),
        ("Document Storage Scalability:", "MongoDB document collections are suitable for moderate throughput. High-volume ingestion (>100k events/sec) requires analytical stores like ClickHouse or OpenSearch."),
        ("Custom Rule Syntax:", "Rules currently use proprietary YAML/JSON definitions. Community Sigma rule integration requires an automated transpiler module."),
        ("No Automated SOAR Remediation:", "The platform focuses on detection and investigation; automated firewall blocking or endpoint isolation via webhooks is not yet implemented.")
    ]
    for h, d in limits:
        p = tf_lim.add_paragraph()
        p.text = "• " + h
        p.font.bold = True
        p.font.size = Pt(10.5)
        p.font.color.rgb = TEXT_WHITE
        p.space_before = Pt(6)
        
        p2 = tf_lim.add_paragraph()
        p2.text = d
        p2.font.size = Pt(10)
        p2.font.color.rgb = TEXT_MUTED

    # Right: Future Work Card
    add_card(s12, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.0))
    tb_fut = s12.shapes.add_textbox(Inches(7.05), Inches(2.0), Inches(5.2), Inches(4.5))
    tf_fut = tb_fut.text_frame
    tf_fut.word_wrap = True

    p = tf_fut.paragraphs[0]
    p.text = "FUTURE ENGINEERING ROADMAP"
    p.font.name = "Arial"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = ACCENT_CYAN

    futures = [
        ("Sigma & YARA-L Rule Transpiler:", "Enable automated import of public Sigma detection rules, translating them into real-time correlation queries."),
        ("SOAR Action Playbooks:", "Incorporate active response webhooks to automatically null-route attacking IPs on network firewalls and isolate compromised hosts."),
        ("ClickHouse Telemetry Datastore:", "Introduce an analytical column-store backend to support sub-second aggregations across billions of historical logs."),
        ("Real-Time Threat Intel Feeds:", "Integrate STIX/TAXII threat feeds to automatically match ingested source IPs and domains against known malicious infrastructure."),
        ("Custom eBPF Telemetry Agents:", "Develop lightweight Linux eBPF and Windows ETW endpoint agents to stream kernel telemetry directly to the SIEM.")
    ]
    for h, d in futures:
        p = tf_fut.add_paragraph()
        p.text = "▸ " + h
        p.font.bold = True
        p.font.size = Pt(10.5)
        p.font.color.rgb = TEXT_WHITE
        p.space_before = Pt(6)
        
        p2 = tf_fut.add_paragraph()
        p2.text = d
        p2.font.size = Pt(10)
        p2.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 13: CONCLUSION & KEY LEARNINGS
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    add_bg(s13)
    add_header(s13, "Conclusion & Professional Learnings", "PROJECT SUMMARY")

    add_card(s13, Inches(0.8), Inches(1.8), Inches(11.7), Inches(5.0))
    tb_con = s13.shapes.add_textbox(Inches(1.1), Inches(2.0), Inches(11.1), Inches(4.5))
    tf_con = tb_con.text_frame
    tf_con.word_wrap = True

    p = tf_con.paragraphs[0]
    p.text = "KEY INDUSTRIAL & TECHNICAL ACHIEVEMENTS"
    p.font.name = "Arial"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = ACCENT_CYAN

    achievements = [
        ("Complete Full-Stack SIEM Delivery:", "Successfully designed, containerized, and deployed a self-hosted SIEM platform integrating React 19, Python 3.14, Redis Streams, MongoDB 9.0, and Nginx."),
        ("Multi-Stage Threat Correlation:", "Proved that open-source sliding-window engines can detect sophisticated threats (brute force bursts, lateral movement, DNS tunneling) in sub-second timeframes."),
        ("Defensive Security Engineering:", "Hardened web services against timing attacks, brute-force bursts (rate-limiting), credential theft (Argon2id), and unauthorized access (RBAC JWT)."),
        ("Realistic Adversary Playbook Emulation:", "Implemented 4 MITRE ATT&CK attack scenarios providing repeatable detection benchmarking for cybersecurity analysts."),
        ("Professional Growth:", "Gained invaluable experience in distributed event streaming, microservices orchestration, defensive cyber operations, and reactive frontend visualization.")
    ]
    for h, d in achievements:
        p = tf_con.add_paragraph()
        p.text = "✔ " + h
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_WHITE
        p.space_before = Pt(8)
        
        p2 = tf_con.add_paragraph()
        p2.text = d
        p2.font.size = Pt(10.5)
        p2.font.color.rgb = TEXT_MUTED
        p2.space_before = Pt(1)

    # =========================================================================
    # SLIDE 14: Q&A / THANK YOU
    # =========================================================================
    s14 = prs.slides.add_slide(blank_layout)
    add_bg(s14)

    # Title Card
    add_card(s14, Inches(1.5), Inches(1.2), Inches(10.333), Inches(5.1), bg_color=RGBColor(16, 32, 58), border_color=ACCENT_CYAN)

    tb_end = s14.shapes.add_textbox(Inches(2.0), Inches(1.6), Inches(9.333), Inches(4.3))
    tf_end = tb_end.text_frame
    tf_end.word_wrap = True

    p = tf_end.paragraphs[0]
    p.text = "THANK YOU"
    p.alignment = PP_ALIGN.CENTER
    p.font.name = "Arial"
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = ACCENT_CYAN

    p = tf_end.add_paragraph()
    p.text = "Questions & Discussion"
    p.alignment = PP_ALIGN.CENTER
    p.font.name = "Calibri"
    p.font.size = Pt(18)
    p.font.color.rgb = TEXT_WHITE
    p.space_before = Pt(6)

    # Divider line
    p = tf_end.add_paragraph()
    p.text = "──────────────────────────────────────────────────────"
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = LINE_BORDER
    p.space_before = Pt(10)

    p = tf_end.add_paragraph()
    p.text = "SIEM DASHBOARD PROJECT"
    p.alignment = PP_ALIGN.CENTER
    p.font.bold = True
    p.font.size = Pt(14)
    p.font.color.rgb = ACCENT_BLUE
    p.space_before = Pt(8)

    p = tf_end.add_paragraph()
    p.text = "Candidate: Aditya Tiwari (Roll No: 72413167 / CTU2401785)\nIndustrial Mentor: Abhishek Jayant (INFOTECHMON PVT. LTD.)\nAcademic Guide: Arvin Vinayek (CT University, Ludhiana)\nLocal Console: http://localhost:8431 | Direct API: http://localhost:5113"
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(11)
    p.font.color.rgb = TEXT_MUTED
    p.space_before = Pt(8)

    output_ppt = "Aditya_Tiwari_SIEM_Dashboard_Presentation.pptx"
    prs.save(output_ppt)
    print(f"Presentation saved successfully as {output_ppt}")

if __name__ == "__main__":
    create_deck()
