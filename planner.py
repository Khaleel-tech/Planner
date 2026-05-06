import streamlit as st
import json, re, os
from groq import Groq
from tavily import TavilyClient
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer,
    Table, TableStyle, PageBreak, HRFlowable
)
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

# ── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AI Learning Roadmap Agent",
    page_icon="🗺️",
    layout="centered"
)

st.title("🗺️ AI Learning Roadmap Agent")
st.caption("Enter a topic and get a day-by-day study plan with PDF and Excel downloads.")

# ── Sidebar — API keys ───────────────────────────────────────────────────────
with st.sidebar:
    st.header("API Keys")
    groq_key   = st.text_input("Groq API Key",   type="password", placeholder="gsk_...")
    tavily_key = st.text_input("Tavily API Key (optional)", type="password", placeholder="tvly-...")
    st.caption("Keys are never stored. They live only in your browser session.")
    st.divider()
    st.markdown("**Models used**")
    st.caption("LLM: llama-3.3-70b-versatile (Groq)")
    st.caption("Search: Tavily basic search")

# ── Main inputs ──────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    topic = st.text_input("What do you want to learn?",
                          placeholder="e.g. Machine Learning, React, DSA...")
    level = st.selectbox("Your current level",
                         ["complete beginner","some basics",
                          "intermediate","advanced"])
with col2:
    days  = st.slider("Number of days",  min_value=3,  max_value=90, value=14)
    hours = st.slider("Study hours/day", min_value=1,  max_value=8,  value=2)

generate_btn = st.button("Generate Roadmap", type="primary", use_container_width=True)

# ── Helper: phase colors ─────────────────────────────────────────────────────
PHASE_FILLS = ["B5D4F4","C0DD97","FAC775","F4C0D1","CECBF6"]
PHASE_TEXT  = ["042C53","173404","412402","4B1528","26215C"]

PDF_FILLS = [
    colors.HexColor("#B5D4F4"), colors.HexColor("#C0DD97"),
    colors.HexColor("#FAC775"), colors.HexColor("#F4C0D1"),
    colors.HexColor("#CECBF6"),
]
PDF_TEXT = [
    colors.HexColor("#042C53"), colors.HexColor("#173404"),
    colors.HexColor("#412402"), colors.HexColor("#4B1528"),
    colors.HexColor("#26215C"),
]

# ── fetch_resources() ────────────────────────────────────────────────────────
def fetch_resources(topic, level, api_key):
    if not api_key:
        return []
    try:
        client  = TavilyClient(api_key=api_key)
        results = client.search(
            query=f"best resources to learn {topic} for {level}",
            search_depth="basic", max_results=5
        )
        return [{"title": r["title"], "url": r["url"], "type": "web"}
                for r in results.get("results", [])]
    except:
        return []

# ── generate_roadmap() ───────────────────────────────────────────────────────
def generate_roadmap(topic, days, hours, level, resources, api_key):
    client = Groq(api_key=api_key)

    resource_context = ""
    if resources:
        lines = [f"{i+1}. {r['title']}: {r['url']}" for i, r in enumerate(resources)]
        resource_context = "\n\nReal web resources found — use these URLs:\n" + "\n".join(lines)

    prompt = f"""You are an expert learning roadmap designer.
Create a detailed {days}-day learning roadmap for someone who wants to learn "{topic}".
The student is a {level} and can study {hours} hours per day.{resource_context}

Return ONLY valid JSON — no markdown, no explanation, no code fences.
{{
  "title": "Learning Roadmap: {topic}",
  "overview": "2-3 sentence summary",
  "totalDays": {days},
  "hoursPerDay": {hours},
  "level": "{level}",
  "phases": [{{"name":"phase name","days":"Day 1-N","goal":"what student achieves"}}],
  "days": [{{
    "day": 1, "phase": 0,
    "title": "short title",
    "description": "specific and actionable, 2-3 sentences",
    "tasks": ["task 1","task 2","task 3"],
    "milestone": false
  }}],
  "resources": [{{"title":"name","url":"https://...","type":"free|paid|book|video"}}]
}}
Rules:
- Exactly {days} day entries
- 3-5 phases: foundation → core → practice → projects → mastery
- milestone: true on every 7th day
- phase is 0-based index into phases array
- include provided URLs in resources
- every day must have specific topic names, not vague instructions
"""

    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.7, max_tokens=6000
    )
    raw = response.choices[0].message.content.strip()
    raw = re.sub(r'^```json\s*','',raw)
    raw = re.sub(r'^```\s*','',raw)
    raw = re.sub(r'\s*```$','',raw)
    return json.loads(raw)

# ── generate_pdf() ───────────────────────────────────────────────────────────
def generate_pdf(roadmap):
    buf    = io.BytesIO()
    PAGE_W = A4[0] - 4*cm
    styles = getSampleStyleSheet()

    doc = SimpleDocTemplate(buf, pagesize=A4,
          rightMargin=2*cm, leftMargin=2*cm,
          topMargin=2*cm,   bottomMargin=2*cm)

    S_TITLE   = ParagraphStyle("t",  parent=styles["Title"],   fontSize=22, spaceAfter=6,   textColor=colors.HexColor("#111111"))
    S_META    = ParagraphStyle("m",  parent=styles["Normal"],  fontSize=10, spaceAfter=12,  textColor=colors.HexColor("#666666"))
    S_SEC     = ParagraphStyle("s",  parent=styles["Heading2"],fontSize=13, spaceBefore=18, spaceAfter=8, textColor=colors.HexColor("#222222"))
    S_DTITLE  = ParagraphStyle("dt", parent=styles["Normal"],  fontSize=12, fontName="Helvetica-Bold", textColor=colors.HexColor("#111111"))
    S_DDESC   = ParagraphStyle("dd", parent=styles["Normal"],  fontSize=10, leading=14, textColor=colors.HexColor("#444444"))
    S_TASK    = ParagraphStyle("tk", parent=styles["Normal"],  fontSize=9,  leading=13, leftIndent=8, textColor=colors.HexColor("#555555"))
    S_RES     = ParagraphStyle("r",  parent=styles["Normal"],  fontSize=10, leading=14, textColor=colors.HexColor("#185FA5"))
    S_FOOTER  = ParagraphStyle("f",  parent=styles["Normal"],  fontSize=8,  alignment=TA_CENTER, textColor=colors.HexColor("#AAAAAA"))

    story = []
    story.append(Paragraph(roadmap["title"], S_TITLE))
    story.append(Paragraph(
        f"{roadmap['totalDays']} days · {roadmap['hoursPerDay']} hrs/day · "
        f"{roadmap['level']} · {roadmap['hoursPerDay']*roadmap['totalDays']}h total", S_META))

    ov = Table([[Paragraph(roadmap["overview"],
                ParagraphStyle("ov", parent=styles["Normal"], fontSize=11, leading=16,
                textColor=colors.HexColor("#333333")))]], colWidths=[PAGE_W])
    ov.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#F8F8F8")),
        ("BOX",(0,0),(-1,-1),0.5,colors.HexColor("#E0E0E0")),
        ("TOPPADDING",(0,0),(-1,-1),10),("BOTTOMPADDING",(0,0),(-1,-1),10),
        ("LEFTPADDING",(0,0),(-1,-1),12),("RIGHTPADDING",(0,0),(-1,-1),12),
    ]))
    story.append(ov)
    story.append(Spacer(1,14))

    story.append(Paragraph("Learning phases", S_SEC))
    pdata = [["Phase","Days","Goal"]]
    for ph in roadmap["phases"]:
        pdata.append([ph["name"], ph["days"], ph.get("goal","")])
    pt = Table(pdata, colWidths=[PAGE_W*0.25, PAGE_W*0.15, PAGE_W*0.60])
    pts = TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#222222")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),10),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F7F7F7")]),
        ("GRID",(0,0),(-1,-1),0.5,colors.HexColor("#E0E0E0")),
        ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
        ("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8),
    ])
    for i in range(1, len(pdata)):
        idx = (i-1) % len(PDF_FILLS)
        pts.add("BACKGROUND",(0,i),(0,i),PDF_FILLS[idx])
        pts.add("TEXTCOLOR",(0,i),(0,i),PDF_TEXT[idx])
        pts.add("FONTNAME",(0,i),(0,i),"Helvetica-Bold")
    pt.setStyle(pts)
    story.append(pt)
    story.append(Spacer(1,16))

    story.append(Paragraph("Day-by-day plan", S_SEC))
    for day in roadmap["days"]:
        idx      = (day.get("phase") or 0) % len(PDF_FILLS)
        label    = f"Day {day['day']}" + (" ★" if day.get("milestone") else "")
        tasks_str= "     ".join(f"• {t}" for t in day.get("tasks",[]))
        lp = Paragraph(label, ParagraphStyle("dl",fontSize=9,
             fontName="Helvetica-Bold", textColor=PDF_TEXT[idx]))
        lc = Table([[lp]], style=TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),PDF_FILLS[idx]),
            ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
            ("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8),
        ]))
        inner = Table(
            [[lc, Paragraph(day["title"], S_DTITLE)],
             ["",  Paragraph(day["description"], S_DDESC)],
             ["",  Paragraph(tasks_str, S_TASK)]],
            colWidths=[PAGE_W*0.14, PAGE_W*0.86]
        )
        inner.setStyle(TableStyle([
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
            ("LEFTPADDING",(1,0),(1,-1),10),("LEFTPADDING",(0,0),(0,-1),0),
        ]))
        outer = Table([[inner]], colWidths=[PAGE_W])
        outer.setStyle(TableStyle([
            ("BOX",(0,0),(-1,-1),0.5,colors.HexColor("#E0E0E0")),
            ("BACKGROUND",(0,0),(-1,-1),colors.white),
            ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
            ("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8),
        ]))
        story.append(outer)
        story.append(Spacer(1,5))

    story.append(PageBreak())
    story.append(Paragraph("Recommended resources", S_SEC))
    for r in roadmap.get("resources",[]):
        story.append(Paragraph(
            f"→  <a href='{r['url']}'><u>{r['title']}</u></a>"
            f"  <font color='#888888' size='9'>({r.get('type','resource')})</font>", S_RES))
        story.append(Spacer(1,4))

    story.append(Spacer(1,30))
    story.append(HRFlowable(width=PAGE_W, thickness=0.5, color=colors.HexColor("#E0E0E0")))
    story.append(Spacer(1,8))
    story.append(Paragraph("Generated by AI Learning Roadmap Agent", S_FOOTER))

    doc.build(story)
    buf.seek(0)
    return buf

# ── generate_excel() ─────────────────────────────────────────────────────────
def generate_excel(roadmap):
    buf    = io.BytesIO()
    wb     = Workbook()
    phases = roadmap.get("phases", [])
    total  = roadmap["totalDays"]

    THIN   = Side(style="thin", color="E0E0E0")
    BDR    = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    AC     = Alignment(horizontal="center", vertical="center", wrap_text=True)
    AL     = Alignment(horizontal="left",   vertical="top",    wrap_text=True)
    HF     = Font(bold=True, color="FFFFFF", size=10)
    HFill  = PatternFill("solid", fgColor="222222")

    # Sheet 1
    ws1 = wb.active
    ws1.title = "Timetable"
    ws1.merge_cells("A1:I1")
    ws1["A1"] = roadmap["title"]
    ws1["A1"].font      = Font(bold=True, size=14, color="FFFFFF")
    ws1["A1"].fill      = PatternFill("solid", fgColor="111111")
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 32

    ws1.merge_cells("A2:I2")
    ws1["A2"] = f"{total} days | {roadmap['hoursPerDay']} hrs/day | {roadmap['level']}"
    ws1["A2"].font      = Font(size=9, color="666666", italic=True)
    ws1["A2"].alignment = Alignment(horizontal="left", vertical="center")
    ws1.row_dimensions[2].height = 22

    headers    = ["Day","Phase","Title","Description","Tasks",
                  "Milestone","Est. Hours","Completed ✓","Notes"]
    col_widths = [6, 18, 28, 48, 42, 11, 12, 14, 28]
    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        c = ws1.cell(row=3, column=col, value=h)
        c.font = HF; c.fill = HFill; c.alignment = AC; c.border = BDR
        ws1.column_dimensions[get_column_letter(col)].width = w
    ws1.row_dimensions[3].height = 20

    for day in roadmap["days"]:
        row    = day["day"] + 3
        ph_idx = (day.get("phase") or 0) % len(PHASE_FILLS)
        pFill  = PatternFill("solid", fgColor=PHASE_FILLS[ph_idx])
        pFont  = Font(size=9, bold=True, color=PHASE_TEXT[ph_idx])
        pname  = phases[day.get("phase",0)]["name"] if phases else ""
        tasks  = "\n".join(f"• {t}" for t in day.get("tasks",[]))
        mstone = "★ Milestone" if day.get("milestone") else ""
        vals   = [day["day"], pname, day["title"], day["description"],
                  tasks, mstone, roadmap["hoursPerDay"], "", ""]
        for col, val in enumerate(vals, 1):
            c = ws1.cell(row=row, column=col, value=val)
            c.border = BDR; c.alignment = AL; c.font = Font(size=9)
            if col in (1,2): c.fill = pFill; c.font = pFont
            if col in (1,6,7,8): c.alignment = AC
            if mstone and col not in (1,2):
                c.fill = PatternFill("solid", fgColor="FFFBF0")
        ws1.row_dimensions[row].height = max(30, 14*len(day.get("tasks",[])))

    ws1.freeze_panes = "A4"
    ws1.auto_filter.ref = f"A3:I{total+3}"

    # Sheet 2
    ws2 = wb.create_sheet("Progress Dashboard")
    ws2["A1"] = "Progress Dashboard"
    ws2["A1"].font = Font(bold=True, size=14)
    ws2.column_dimensions["A"].width = 22
    ws2.column_dimensions["B"].width = 18
    ws2.column_dimensions["C"].width = 40
    for col, lbl in enumerate(["Metric","Value"], 1):
        c = ws2.cell(row=3, column=col, value=lbl)
        c.font = HF; c.fill = HFill; c.alignment = AC

    metrics = [
        ("Total days", total),
        ("Hours per day", roadmap["hoursPerDay"]),
        ("Total study hours", roadmap["hoursPerDay"]*total),
        ("Phases", len(phases)),
        ("Days completed",   f'=COUNTIF(Timetable!H4:H{total+3},"✓")'),
        ("Days remaining",   f"={total}-B8"),
        ("% Complete",       "=ROUND(B8/B4*100,1)"),
    ]
    for i, (lbl, val) in enumerate(metrics, 4):
        ws2.cell(row=i,column=1,value=lbl).font  = Font(size=10)
        ws2.cell(row=i,column=2,value=val).font  = Font(size=10,bold=True)
        ws2.cell(row=i,column=1).alignment = Alignment(horizontal="left")
        ws2.cell(row=i,column=2).alignment = Alignment(horizontal="center")
        af = PatternFill("solid", fgColor="F7F7F7" if i%2==0 else "FFFFFF")
        ws2.cell(row=i,column=1).fill = af
        ws2.cell(row=i,column=2).fill = af
        for col in (1,2): ws2.cell(row=i,column=col).border = BDR

    ws2["A12"] = "Phase breakdown"
    ws2["A12"].font = Font(bold=True, size=11)
    for col, lbl in enumerate(["Phase","Days","Goal"],1):
        c = ws2.cell(row=13,column=col,value=lbl)
        c.font=HF; c.fill=HFill; c.alignment=AC
    for i, ph in enumerate(phases):
        row = 14+i; idx = i%len(PHASE_FILLS)
        for col, val in enumerate([ph["name"],ph["days"],ph.get("goal","")],1):
            c = ws2.cell(row=row,column=col,value=val)
            c.font      = Font(size=9,bold=(col<=2),color=PHASE_TEXT[idx])
            c.fill      = PatternFill("solid",fgColor=PHASE_FILLS[idx])
            c.alignment = AL; c.border = BDR

    # Sheet 3
    ws3 = wb.create_sheet("Resources")
    ws3["A1"] = "Learning Resources"
    ws3["A1"].font = Font(bold=True,size=14)
    ws3.column_dimensions["A"].width=35
    ws3.column_dimensions["B"].width=55
    ws3.column_dimensions["C"].width=12
    for col, lbl in enumerate(["Title","URL","Type"],1):
        c = ws3.cell(row=3,column=col,value=lbl)
        c.font=HF; c.fill=HFill; c.alignment=AC
    for i, r in enumerate(roadmap.get("resources",[]),4):
        af = PatternFill("solid",fgColor="F7F7F7" if i%2==0 else "FFFFFF")
        for col, val in enumerate([r["title"],r["url"],r.get("type","")],1):
            c = ws3.cell(row=i,column=col,value=val)
            c.font      = Font(size=10,color="185FA5" if col==2 else "111111")
            c.fill      = af; c.alignment=AL; c.border=BDR

    wb.save(buf)
    buf.seek(0)
    return buf

# ── Main logic ───────────────────────────────────────────────────────────────
if generate_btn:
    if not topic:
        st.error("Please enter a topic.")
    elif not groq_key:
        st.error("Please enter your Groq API key in the sidebar.")
    else:
        with st.spinner("Searching resources..."):
            resources = fetch_resources(topic, level, tavily_key)

        with st.spinner("Generating roadmap with AI..."):
            try:
                roadmap = generate_roadmap(topic, days, hours, level,
                                           resources, groq_key)
            except Exception as e:
                st.error(f"Groq error: {e}")
                st.stop()

        # ── Display roadmap ──────────────────────────────────────────────────
        st.success(f"Roadmap generated — {len(roadmap['days'])} days across {len(roadmap['phases'])} phases")

        st.subheader(roadmap["title"])
        st.caption(roadmap["overview"])

        # Stats
        c1, c2, c3 = st.columns(3)
        c1.metric("Total days",        roadmap["totalDays"])
        c2.metric("Total study hours", roadmap["hoursPerDay"] * roadmap["totalDays"])
        c3.metric("Phases",            len(roadmap["phases"]))

        # Phases
        st.markdown("### Phases")
        for i, ph in enumerate(roadmap["phases"]):
            st.markdown(f"**{ph['name']}** · {ph['days']} — {ph['goal']}")

        # Days
        st.markdown("### Day-by-day plan")
        for day in roadmap["days"]:
            label = f"Day {day['day']} — {day['title']}" + (" ★" if day.get("milestone") else "")
            with st.expander(label):
                st.write(day["description"])
                for t in day.get("tasks", []):
                    st.markdown(f"- {t}")

        # Resources
        if roadmap.get("resources"):
            st.markdown("### Resources")
            for r in roadmap["resources"]:
                st.markdown(f"- [{r['title']}]({r['url']}) `{r.get('type','')}`")

        # ── Download buttons ─────────────────────────────────────────────────
        st.markdown("### Download")
        d1, d2 = st.columns(2)

        with st.spinner("Building PDF..."):
            pdf_buf = generate_pdf(roadmap)
        with d1:
            st.download_button(
                label="Download PDF Roadmap",
                data=pdf_buf,
                file_name=f"{topic.replace(' ','_')}_roadmap.pdf",
                mime="application/pdf",
                use_container_width=True
            )

        with st.spinner("Building Excel..."):
            xl_buf = generate_excel(roadmap)
        with d2:
            st.download_button(
                label="Download Excel Tracker",
                data=xl_buf,
                file_name=f"{topic.replace(' ','_')}_tracker.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
