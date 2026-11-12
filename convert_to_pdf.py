"""
Convert Road_Accident_AI_Hackathon_2026.pptx to PDF
Uses python-pptx to extract content and reportlab to render a clean PDF.
"""
from pptx import Presentation
from pptx.util import Pt
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfgen import canvas
from reportlab.platypus import PageBreak
import os, sys

PPTX_PATH = os.path.join(os.path.dirname(__file__), "Road_Accident_AI_Hackathon_2026.pptx")
PDF_PATH  = os.path.join(os.path.dirname(__file__), "Road_Accident_AI_Hackathon_2026.pdf")

# ── Colours ────────────────────────────────────────────────────────────────
NAVY   = colors.HexColor("#174173")
BLUE   = colors.HexColor("#2E75B6")
WHITE  = colors.white
LTBLUE = colors.HexColor("#EFF3FB")
GRAY   = colors.HexColor("#333333")
BGRAY  = colors.HexColor("#F5F5F5")
BLACK  = colors.black

PW, PH = landscape(A4)   # 841 x 595 pt

# ── Styles ─────────────────────────────────────────────────────────────────
styles = getSampleStyleSheet()

def S(name, **kw):
    return ParagraphStyle(name, **kw)

HEADER   = S("Header",   fontName="Helvetica-Bold", fontSize=9,  textColor=WHITE,  alignment=TA_LEFT,   leading=11)
TITLE    = S("Title",    fontName="Helvetica-Bold", fontSize=26, textColor=NAVY,   alignment=TA_CENTER, leading=32, spaceAfter=6)
SUBTITLE = S("Subtitle", fontName="Helvetica",      fontSize=14, textColor=BLUE,   alignment=TA_CENTER, leading=18, spaceAfter=4)
SLIDE_T  = S("SlideT",   fontName="Helvetica-Bold", fontSize=16, textColor=NAVY,   alignment=TA_LEFT,   leading=20, spaceAfter=6)
BODY     = S("Body",     fontName="Helvetica",      fontSize=10, textColor=GRAY,   alignment=TA_LEFT,   leading=14, spaceAfter=3)
BODY_SM  = S("BodySm",   fontName="Helvetica",      fontSize=9,  textColor=GRAY,   alignment=TA_LEFT,   leading=12, spaceAfter=2)
BODY_WH  = S("BodyWh",   fontName="Helvetica",      fontSize=10, textColor=WHITE,  alignment=TA_LEFT,   leading=14, spaceAfter=2)
BOLD_WH  = S("BoldWh",   fontName="Helvetica-Bold", fontSize=11, textColor=WHITE,  alignment=TA_CENTER, leading=14, spaceAfter=2)
BOLD_NV  = S("BoldNv",   fontName="Helvetica-Bold", fontSize=11, textColor=NAVY,   alignment=TA_LEFT,   leading=14, spaceAfter=2)
BOLD_BL  = S("BoldBl",   fontName="Helvetica-Bold", fontSize=11, textColor=BLUE,   alignment=TA_LEFT,   leading=14, spaceAfter=2)
CAPTION  = S("Caption",  fontName="Helvetica",      fontSize=8,  textColor=BLUE,   alignment=TA_CENTER, leading=10)
CODE     = S("Code",     fontName="Courier",        fontSize=8,  textColor=GRAY,   alignment=TA_LEFT,   leading=11, spaceAfter=1)
FOOTER   = S("Footer",   fontName="Helvetica",      fontSize=7,  textColor=WHITE,  alignment=TA_CENTER, leading=9)
THANKYOU = S("TY",       fontName="Helvetica-Bold", fontSize=13, textColor=NAVY,   alignment=TA_CENTER, leading=16, spaceAfter=4)

# ── Page template with header/footer ───────────────────────────────────────
def draw_slide_frame(c, doc, slide_num, total=6):
    """Draw navy top bar, blue left accent, bottom bar, hackathon label on every page."""
    # Top bar
    c.setFillColor(NAVY)
    c.rect(0, PH - 34, PW, 34, fill=1, stroke=0)
    # Header text
    c.setFillColor(WHITE)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(14, PH - 22, "Edunet Hackathon in collaboration with IBM SkillsBuild on IBM Bob")
    # Slide number (right)
    c.setFont("Helvetica", 8)
    c.drawRightString(PW - 14, PH - 22, f"{slide_num} / {total}")
    # Left blue accent bar
    c.setFillColor(BLUE)
    c.rect(0, 18, 8, PH - 52, fill=1, stroke=0)
    # Bottom bar
    c.setFillColor(NAVY)
    c.rect(0, 0, PW, 18, fill=1, stroke=0)
    c.setFillColor(WHITE)
    c.setFont("Helvetica", 7)
    c.drawCentredString(PW / 2, 5, "AI-Powered Road Accident Prediction & Emergency Response System  |  Built with IBM Bob")

class SlideCanvas(canvas.Canvas):
    def __init__(self, filename, slide_num_map):
        super().__init__(filename, pagesize=landscape(A4))
        self._slide_num_map = slide_num_map
        self._page = 1

    def showPage(self):
        draw_slide_frame(self, None, self._page)
        self._page += 1
        super().showPage()

    def save(self):
        draw_slide_frame(self, None, self._page)
        super().save()


# ── Helper to build coloured box rows for tables ───────────────────────────
def box_cell(text_lines, bg, style=BODY_WH):
    paras = [Paragraph(t, style) for t in text_lines if t]
    return paras

def navy_box(title, body_lines):
    """Return a navy Table cell with title bold and body lines."""
    content = [Paragraph(title, BOLD_WH)]
    for line in body_lines:
        content.append(Paragraph(line, BODY_WH))
    return content

def blue_box(title, body_lines):
    content = [Paragraph(title, BOLD_WH)]
    for line in body_lines:
        content.append(Paragraph(line, BODY_WH))
    return content

def tbl(data, col_widths, row_styles=None):
    t = Table(data, colWidths=col_widths)
    base = [
        ("TOPPADDING",    (0,0), (-1,-1), 6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
        ("LEFTPADDING",   (0,0), (-1,-1), 8),
        ("RIGHTPADDING",  (0,0), (-1,-1), 8),
        ("VALIGN",        (0,0), (-1,-1), "TOP"),
    ]
    if row_styles:
        base += row_styles
    t.setStyle(TableStyle(base))
    return t

# ── Slide content builders ──────────────────────────────────────────────────
MARGIN_L = 22   # left margin (pt) to clear accent bar
W = PW - MARGIN_L - 14  # usable width

def slide1():
    """Title Slide"""
    story = []
    story.append(Spacer(1, 28))
    story.append(Paragraph("AI-Powered Road Accident Prediction<br/>&amp; Emergency Response System", TITLE))
    story.append(Paragraph("Multi-Agent Agentic AI Platform for Public Road Safety", SUBTITLE))
    story.append(HRFlowable(width=W, thickness=1.5, color=NAVY, spaceAfter=10))
    need_data = [[
        [
            Paragraph("Need of Project", S("NT", fontName="Helvetica-Bold", fontSize=12, textColor=NAVY, leading=14, spaceAfter=4)),
            Paragraph("Over 1.35 million people die in road accidents globally every year (WHO, 2023). "
                      "India records 4.61 lakh accidents and 1.68 lakh fatalities annually (MoRTH 2023). "
                      "Accident-prone zones are not identified in advance, and emergency services face delays "
                      "due to traffic congestion and lack of coordinated AI intelligence. "
                      "This platform predicts accident risks before they occur — shifting from reactive to proactive road safety.", BODY),
        ]
    ]]
    need_t = tbl(need_data, [W], [("BACKGROUND",(0,0),(-1,-1),LTBLUE), ("BOX",(0,0),(-1,-1),1,NAVY)])
    story.append(need_t)
    story.append(Spacer(1, 10))
    info_data = [[
        Paragraph("<b>Team:</b> Road Safety AI", BODY),
        Paragraph("<b>Institute:</b> Bengaluru Hackathon 2026", BODY),
        Paragraph("<b>GitHub:</b> github.com/jharanabisoyi8050-alt/ibm", BODY),
    ]]
    story.append(tbl(info_data, [W/3]*3, [("BACKGROUND",(0,0),(-1,-1),NAVY),
        ("TEXTCOLOR",(0,0),(-1,-1),WHITE), ("ALIGN",(0,0),(-1,-1),"CENTER")]))
    return story

def slide2():
    """Project Overview & Team"""
    story = []
    story.append(Spacer(1, 8))
    story.append(Paragraph("Project Overview &amp; Team", SLIDE_T))
    story.append(Paragraph("<b>Project:</b> AI-Powered Road Accident Prediction &amp; Emergency Response System",
                            S("PT", fontName="Helvetica-Bold", fontSize=11, textColor=NAVY, leading=14, spaceAfter=6)))
    story.append(HRFlowable(width=W, thickness=1, color=NAVY, spaceAfter=6))

    need_col = [
        Paragraph("<b>Need of Project</b>", BOLD_NV),
        Paragraph("• 1.35M deaths/year globally (WHO 2023)", BODY_SM),
        Paragraph("• India: 4.61L accidents, 1.68L fatalities (MoRTH 2023)", BODY_SM),
        Paragraph("• Emergency response delayed by traffic congestion", BODY_SM),
        Paragraph("• Accident hotspots not identified until repeated tragedies", BODY_SM),
        Paragraph("• No unified AI system for prediction + monitoring", BODY_SM),
        Paragraph("• Manual systems lack speed, accuracy &amp; real-time awareness", BODY_SM),
    ]
    meth_col = [
        Paragraph("<b>Methodology</b>", BOLD_NV),
        Paragraph("1. Collect historical accident, traffic &amp; weather datasets", BODY_SM),
        Paragraph("2. Train XGBoost risk prediction model (0–100 score)", BODY_SM),
        Paragraph("3. DBSCAN clustering to identify accident hotspots", BODY_SM),
        Paragraph("4. Build 7-agent LangGraph system with IBM Bob", BODY_SM),
        Paragraph("5. Integrate 3 LLM providers via IBM Bob Model Gateway", BODY_SM),
        Paragraph("6. Deploy custom MCP server (traffic/weather/geo/ambulance)", BODY_SM),
        Paragraph("7. Build Streamlit dashboard: heatmaps, risk gauges, HITL", BODY_SM),
        Paragraph("8. Implement Human-in-the-Loop for HIGH/CRITICAL dispatch", BODY_SM),
    ]
    two_col = tbl([[need_col, meth_col]], [W*0.46, W*0.54],
                  [("BACKGROUND",(0,0),(-1,-1),LTBLUE), ("BOX",(0,0),(-1,-1),1,BLUE)])
    story.append(two_col)
    story.append(Spacer(1, 6))

    # IBM Tools Used
    story.append(Paragraph("<b>IBM TOOLS USED</b>", S("IT", fontName="Helvetica-Bold", fontSize=11, textColor=NAVY, alignment=TA_CENTER, leading=14, spaceAfter=4)))
    tools = [
        [navy_box("IBM Granite (watsonx)", ["Orchestrator &", "Hotspot Detection Agent"]),
         blue_box("IBM Bob", ["Agentic AI development,", "code generation &amp; MCP"]),
         navy_box("IBM Bob Model Gateway", ["Multi-LLM routing:", "watsonx + Bedrock + Vertex"]),
         blue_box("Agentic AI (LangGraph)", ["Multi-agent orchestration,", "HITL &amp; memory"])],
    ]
    tools_t = tbl(tools, [W/4]*4,
                  [("BACKGROUND",(0,0),(0,-1),NAVY), ("BACKGROUND",(1,0),(1,-1),BLUE),
                   ("BACKGROUND",(2,0),(2,-1),NAVY), ("BACKGROUND",(3,0),(3,-1),BLUE),
                   ("ALIGN",(0,0),(-1,-1),"CENTER"), ("ROWBACKGROUNDS",(0,0),(-1,-1),[NAVY, BLUE])])
    story.append(tools_t)
    story.append(Spacer(1, 6))

    # Conclusion
    conc = tbl([[
        [Paragraph("<b>Conclusion &amp; Demonstration</b>", BOLD_NV),
         Paragraph("Impact: AI-predicted risk alerts enable proactive traffic management, reducing high-risk "
                   "zone incidents by up to 30% through early intervention. The Streamlit dashboard demonstrates: "
                   "(1) Real-time risk scoring &amp; heatmaps, (2) AI-generated emergency response plans, "
                   "(3) Ambulance coordination with ETA, (4) Admin panel with LLM cost monitoring &amp; HITL approval queue.", BODY_SM)]
    ]], [W], [("BACKGROUND",(0,0),(-1,-1),LTBLUE), ("BOX",(0,0),(-1,-1),1,NAVY)])
    story.append(conc)
    return story

def slide3():
    """Development with IBM Bob"""
    story = []
    story.append(Spacer(1, 8))
    story.append(Paragraph("Development with IBM Bob", SLIDE_T))
    story.append(HRFlowable(width=W, thickness=1, color=NAVY, spaceAfter=6))

    steps_row1 = [
        [navy_box("1  REQUIREMENTS", ["Capture project goals, agent specs,", "DB schema &amp; acceptance criteria."]),
         blue_box("2  BOB CREATES PLAN", ["Multi-agent architecture,", "LangGraph state machine, file structure."]),
         navy_box("3  REVIEW &amp; REFINE", ["Review agent workflow, HITL gates,", "LLM routing &amp; data labeling policy."]),
         blue_box("4  CODE GENERATION", ["FastAPI backend, 7 LangGraph agents,", "MCP server, Streamlit dashboard, tests."]),
         navy_box("5  REVIEW OUTPUT", ["Review HITL enforcement,", "AI_PREDICTED labels &amp; SIMULATED flags."])],
    ]
    row1_t = tbl(steps_row1, [W/5]*5,
                 [("BACKGROUND",(0,0),(0,-1),NAVY), ("BACKGROUND",(1,0),(1,-1),BLUE),
                  ("BACKGROUND",(2,0),(2,-1),NAVY), ("BACKGROUND",(3,0),(3,-1),BLUE),
                  ("BACKGROUND",(4,0),(4,-1),NAVY), ("ALIGN",(0,0),(-1,-1),"CENTER")])
    story.append(row1_t)
    story.append(Spacer(1, 4))

    steps_row2 = [
        [blue_box("6  RUN &amp; VALIDATE", ["Streamlit app end-to-end:", "risk, heatmaps, ambulance dispatch."]),
         navy_box("7  BOB EXPLAINS", ["LangGraph transitions, MCP JSON-RPC", "calls &amp; LLM routing decisions."]),
         blue_box("8  QUALITY &amp; TESTS", ["pytest: agent unit tests, API", "integration &amp; ML accuracy tests."]),
         navy_box("9  DOCUMENTATION", ["README.md, AGENTS.md,", ".bob/custom_modes.yaml, API docs,", "inline code comments."]),
         blue_box("GitHub", ["84 files  |  6,607 lines", "All generated with IBM Bob"])],
    ]
    row2_t = tbl(steps_row2, [W/5]*5,
                 [("BACKGROUND",(0,0),(0,-1),BLUE), ("BACKGROUND",(1,0),(1,-1),NAVY),
                  ("BACKGROUND",(2,0),(2,-1),BLUE), ("BACKGROUND",(3,0),(3,-1),NAVY),
                  ("BACKGROUND",(4,0),(4,-1),BLUE), ("ALIGN",(0,0),(-1,-1),"CENTER")])
    story.append(row2_t)
    story.append(Spacer(1, 6))

    mode_data = [[
        [Paragraph("<b>IBM Bob Custom Mode: road-safety-ai</b>", BOLD_NV),
         Paragraph("Workspace-scoped custom mode in <b>.bob/custom_modes.yaml</b> — road safety expert persona "
                   "with Agent + MCP + execute permissions. <b>AGENTS.md</b> gives Bob full project architecture "
                   "context on every task.", BODY_SM)],
        [Paragraph("<b>GitHub Repository</b>", BOLD_NV),
         Paragraph("github.com/jharanabisoyi8050-alt/ibm", BODY_SM),
         Paragraph("Streamlit Dashboard: port 8501  |  API Docs: port 8000/docs", BODY_SM)],
    ]]
    story.append(tbl(mode_data, [W*0.6, W*0.4],
                     [("BACKGROUND",(0,0),(0,-1),LTBLUE), ("BACKGROUND",(1,0),(1,-1),LTBLUE),
                      ("BOX",(0,0),(-1,-1),1,NAVY)]))
    return story

def slide4():
    """IBM Bob Files & Code Generated"""
    story = []
    story.append(Spacer(1, 8))
    story.append(Paragraph("IBM Bob — Files &amp; Code Generated", SLIDE_T))
    story.append(HRFlowable(width=W, thickness=1, color=NAVY, spaceAfter=6))

    fs_lines = [
        "road-accident-ai/",
        "├── .bob/custom_modes.yaml    # IBM Bob custom mode",
        "├── AGENTS.md                 # Project context for Bob",
        "├── backend/app/",
        "│   ├── agents/               # 7 LangGraph AI agents",
        "│   ├── api/routes/           # 6 FastAPI route files",
        "│   ├── db/models.py          # PostgreSQL ORM models",
        "│   ├── llm_gateway/          # Multi-LLM routing",
        "│   ├── mcp_server/           # Custom MCP server",
        "│   ├── ml/                   # XGBoost + DBSCAN",
        "│   ├── observability/        # LangSmith + Prometheus",
        "│   └── governance/hitl.py    # HITL approval queue",
        "├── streamlit_app/            # 7-page Streamlit dashboard",
        "└── infra/                    # Docker + K8s + CI/CD",
    ]
    agents_col = [
        Paragraph("<b>IBM Bob — Agents &amp; Modules Generated</b>", BOLD_NV),
        Spacer(1, 4),
        tbl([[[Paragraph("<b>orchestrator.py</b>", BOLD_WH),
               Paragraph("LangGraph state machine, fan-out to 6 agents + HITL gate. Model: watsonx Granite-3", BODY_WH)]]], [W*0.54-16],
            [("BACKGROUND",(0,0),(-1,-1),NAVY)]),
        Spacer(1, 3),
        tbl([[[Paragraph("<b>traffic_agent.py | road_agent.py | ambulance_agent.py</b>", BOLD_WH),
               Paragraph("Gemini Flash — traffic density, road condition &amp; ambulance coordination", BODY_WH)]]], [W*0.54-16],
            [("BACKGROUND",(0,0),(-1,-1),BLUE)]),
        Spacer(1, 3),
        tbl([[[Paragraph("<b>risk_agent.py | emergency_agent.py</b>", BOLD_WH),
               Paragraph("Claude Sonnet (Bedrock) — XGBoost risk scoring + HITL emergency recommendations", BODY_WH)]]], [W*0.54-16],
            [("BACKGROUND",(0,0),(-1,-1),NAVY)]),
        Spacer(1, 3),
        tbl([[[Paragraph("<b>hotspot_agent.py + hotspot_model.py</b>", BOLD_WH),
               Paragraph("watsonx Granite + DBSCAN — generates heatmap data for Plotly density_mapbox", BODY_WH)]]], [W*0.54-16],
            [("BACKGROUND",(0,0),(-1,-1),BLUE)]),
    ]

    fs_paras = [Paragraph("<b>IBM Bob — Project File Structure (84 files, 6,607 lines)</b>", BOLD_NV)] + \
               [Paragraph(l, CODE) for l in fs_lines]

    layout = tbl([[fs_paras, agents_col]], [W*0.44, W*0.56],
                 [("BACKGROUND",(0,0),(0,-1),BGRAY), ("BOX",(0,0),(0,-1),1,NAVY),
                  ("VALIGN",(0,0),(-1,-1),"TOP")])
    story.append(layout)
    story.append(Spacer(1, 6))

    mcp_data = [[
        [Paragraph("<b>traffic_mcp.py</b>", BOLD_WH), Paragraph("get_traffic_density\nget_congestion_level\nget_vehicle_count", BODY_WH)],
        [Paragraph("<b>weather_mcp.py</b>", BOLD_WH), Paragraph("get_current_weather\nget_weather_risk\nget_weather_forecast", BODY_WH)],
        [Paragraph("<b>geo_mcp.py</b>", BOLD_WH), Paragraph("geocode_address\nget_route\nget_nearby_hospitals", BODY_WH)],
        [Paragraph("<b>gateway_config.yaml</b>", BOLD_WH), Paragraph("watsonx Granite-3\nClaude Sonnet (Bedrock)\nGemini Flash (Vertex)", BODY_WH)],
    ]]
    mcp_t = tbl(mcp_data, [W/4]*4,
                [("BACKGROUND",(0,0),(0,-1),NAVY), ("BACKGROUND",(1,0),(1,-1),BLUE),
                 ("BACKGROUND",(2,0),(2,-1),NAVY), ("BACKGROUND",(3,0),(3,-1),BLUE),
                 ("ALIGN",(0,0),(-1,-1),"CENTER")])
    story.append(Paragraph("<b>MCP Server Tool Groups &amp; LLM Gateway</b>",
                            S("MCPt", fontName="Helvetica-Bold", fontSize=10, textColor=NAVY, leading=13, spaceAfter=3)))
    story.append(mcp_t)
    return story

def slide5():
    """Features of the Project"""
    story = []
    story.append(Spacer(1, 8))
    story.append(Paragraph("Features of the Project", SLIDE_T))
    story.append(HRFlowable(width=W, thickness=1, color=NAVY, spaceAfter=6))

    f_data = [[
        navy_box("MULTI-AGENT AI SYSTEM",
                 ["7 LangGraph agents: Orchestrator, Traffic,", "Risk Prediction, Road Condition,",
                  "Hotspot Detection, Emergency Response,", "Ambulance Coordination."]),
        blue_box("XGBoost RISK PREDICTION",
                 ["Predicts risk score 0–100 using traffic density,", "weather, road quality &amp; time-of-day.",
                  "Labeled AI_PREDICTED — never confused", "with confirmed accidents."]),
        navy_box("ACCIDENT HOTSPOT HEATMAPS",
                 ["DBSCAN clustering on historical accident", "coordinates generates hotspot zones.",
                  "Interactive Plotly density_mapbox", "heatmaps on Streamlit dashboard."]),
    ]]
    story.append(tbl(f_data, [W/3]*3,
                     [("BACKGROUND",(0,0),(0,-1),NAVY), ("BACKGROUND",(1,0),(1,-1),BLUE),
                      ("BACKGROUND",(2,0),(2,-1),NAVY), ("ALIGN",(0,0),(-1,-1),"LEFT")]))
    story.append(Spacer(1, 4))

    f_data2 = [[
        blue_box("HUMAN-IN-THE-LOOP (HITL)",
                 ["HIGH &amp; CRITICAL emergency dispatch requires", "human officer approval before any action.",
                  "HITL queue in Admin Panel.", "Every decision logged for audit."]),
        navy_box("LLM GATEWAY &amp; COST MONITOR",
                 ["IBM Bob Model Gateway routes agents to", "watsonx Granite, Claude Sonnet, Gemini Flash.",
                  "Token usage &amp; USD cost per agent", "logged in real time. 7-day cost trend."]),
        blue_box("STREAMLIT DASHBOARD",
                 ["7-page interactive dashboard:", "Traffic Monitor, Hotspot Heatmap,",
                  "Risk Gauge, Emergency Report,", "Ambulance Coordination, Admin Panel."]),
    ]]
    story.append(tbl(f_data2, [W/3]*3,
                     [("BACKGROUND",(0,0),(0,-1),BLUE), ("BACKGROUND",(1,0),(1,-1),NAVY),
                      ("BACKGROUND",(2,0),(2,-1),BLUE), ("ALIGN",(0,0),(-1,-1),"LEFT")]))
    story.append(Spacer(1, 5))

    extra = tbl([[
        [Paragraph("<b>Additional Features</b>", BOLD_NV),
         Paragraph("• MCP Tool Server (traffic, weather, geo, ambulance)  "
                   "• Agent Memory: Redis session + ChromaDB pgvector  "
                   "• Guardrails: injection protection + Bedrock guardrailConfig  "
                   "• LangSmith Tracing + Prometheus + Grafana  "
                   "• Data Labeling: AI_PREDICTED / CONFIRMED / SIMULATED enforced at DB + UI  "
                   "• Kubernetes + GitHub Actions CI/CD Pipeline", BODY_SM)]
    ]], [W], [("BACKGROUND",(0,0),(-1,-1),LTBLUE), ("BOX",(0,0),(-1,-1),1,NAVY)])
    story.append(extra)
    return story

def slide6():
    """Novelty & Future Scope"""
    story = []
    story.append(Spacer(1, 8))
    story.append(Paragraph("Novelty &amp; Future Scope", SLIDE_T))
    story.append(HRFlowable(width=W, thickness=1, color=NAVY, spaceAfter=6))

    nov_col = [
        tbl([[[Paragraph("<b>NOVELTY &amp; UNIQUENESS</b>", BOLD_WH)]]], [W*0.46-4],
            [("BACKGROUND",(0,0),(-1,-1),NAVY), ("ALIGN",(0,0),(-1,-1),"CENTER")]),
        Spacer(1, 4),
        tbl([[[Paragraph("<b>Predictive, Not Reactive</b>", BOLD_NV),
               Paragraph("XGBoost predicts accident risk BEFORE incidents — enabling proactive intervention. "
                         "Unlike traditional systems that only respond after accidents.", BODY_SM)]]], [W*0.46-4],
            [("BACKGROUND",(0,0),(-1,-1),LTBLUE), ("BOX",(0,0),(-1,-1),1,NAVY)]),
        Spacer(1, 3),
        tbl([[[Paragraph("<b>Multi-LLM Agent Orchestration</b>", BOLD_NV),
               Paragraph("7 specialist agents each use their own optimized LLM (watsonx / Bedrock / Vertex) — "
                         "first road safety system combining IBM Bob Agentic AI with multi-LLM orchestration.", BODY_SM)]]], [W*0.46-4],
            [("BACKGROUND",(0,0),(-1,-1),LTBLUE), ("BOX",(0,0),(-1,-1),1,BLUE)]),
        Spacer(1, 3),
        tbl([[[Paragraph("<b>MCP-Integrated Live Tool Calls</b>", BOLD_NV),
               Paragraph("Custom JSON-RPC 2.0 MCP server with 4 tool groups (traffic, weather, geo, ambulance) "
                         "enables agents to call live external tools — not just static knowledge.", BODY_SM)]]], [W*0.46-4],
            [("BACKGROUND",(0,0),(-1,-1),LTBLUE), ("BOX",(0,0),(-1,-1),1,NAVY)]),
        Spacer(1, 3),
        tbl([[[Paragraph("<b>HITL + Responsible AI Governance</b>", BOLD_NV),
               Paragraph("HIGH/CRITICAL dispatch requires human approval. AI predictions labeled AI_PREDICTED. "
                         "Simulated data labeled is_simulated=True. Full audit log.", BODY_SM)]]], [W*0.46-4],
            [("BACKGROUND",(0,0),(-1,-1),LTBLUE), ("BOX",(0,0),(-1,-1),1,BLUE)]),
    ]
    fs_col = [
        tbl([[[Paragraph("<b>FUTURE SCOPE</b>", BOLD_WH)]]], [W*0.54-4],
            [("BACKGROUND",(0,0),(-1,-1),BLUE), ("ALIGN",(0,0),(-1,-1),"CENTER")]),
        Spacer(1, 4),
    ]
    fs_items = [
        ("<b>Live Traffic Integration</b>", "Connect to HERE Maps / TomTom / government APIs to replace simulated traffic with real-time feeds."),
        ("<b>IoT Sensor Integration</b>",   "Integrate speed cameras, weather stations &amp; pothole detectors for real-time road monitoring."),
        ("<b>Real CAD Ambulance Dispatch</b>", "Replace simulated fleet with live CAD system integration for actual emergency dispatch."),
        ("<b>Mobile App for First Responders</b>", "Smartphone app delivering risk alerts, hotspot navigation &amp; AI response plans to field officers."),
        ("<b>Advanced ML &amp; Computer Vision</b>", "CCTV/dashcam vision models for automatic accident detection. Upgrade to deep learning ensemble."),
    ]
    for bold, body in fs_items:
        fs_col.append(tbl([[[Paragraph(bold, BOLD_NV), Paragraph(body, BODY_SM)]]], [W*0.54-4],
                          [("BACKGROUND",(0,0),(-1,-1),LTBLUE), ("BOX",(0,0),(-1,-1),1,BLUE)]))
        fs_col.append(Spacer(1, 3))

    layout = tbl([[nov_col, fs_col]], [W*0.46, W*0.54],
                 [("VALIGN",(0,0),(-1,-1),"TOP")])
    story.append(layout)
    story.append(Spacer(1, 6))

    story.append(tbl([[
        Paragraph("<b>Thank You  |  AI-Powered Road Accident Prediction &amp; Emergency Response System  |  Built with IBM Bob</b>",
                  S("TY2", fontName="Helvetica-Bold", fontSize=11, textColor=WHITE, alignment=TA_CENTER, leading=14)),
    ]], [W], [("BACKGROUND",(0,0),(-1,-1),NAVY), ("ALIGN",(0,0),(-1,-1),"CENTER")]))
    return story

# ── Build PDF ──────────────────────────────────────────────────────────────
SLIDES = [slide1, slide2, slide3, slide4, slide5, slide6]

def build_pdf():
    story = []
    for i, slide_fn in enumerate(SLIDES):
        story += slide_fn()
        if i < len(SLIDES) - 1:
            story.append(PageBreak())

    doc = SimpleDocTemplate(
        PDF_PATH,
        pagesize=landscape(A4),
        leftMargin=MARGIN_L + 4,
        rightMargin=14,
        topMargin=42,
        bottomMargin=24,
        title="AI-Powered Road Accident Prediction & Emergency Response System",
        author="IBM Bob — Bengaluru Hackathon 2026",
    )

    # Custom canvas to draw slide frames
    class FrameCanvas(canvas.Canvas):
        def __init__(self, filename, **kw):
            super().__init__(filename, **kw)
            self._page_num = 1
        def showPage(self):
            draw_slide_frame(self, None, self._page_num)
            self._page_num += 1
            super().showPage()
        def save(self):
            draw_slide_frame(self, None, self._page_num)
            super().save()

    doc.build(story, canvasmaker=FrameCanvas)
    print(f"PDF created: {PDF_PATH}")

if __name__ == "__main__":
    build_pdf()
