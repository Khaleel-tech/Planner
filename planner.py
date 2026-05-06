import streamlit as st
import json, re, io

from groq import Groq
from tavily import TavilyClient

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

from openpyxl import Workbook


# ─────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────
st.set_page_config(page_title="AI Roadmap Generator", page_icon="🗺️")
st.title("🗺️ AI Learning Roadmap Generator")

# ─────────────────────────────────────────
# LOAD SECRETS
# ─────────────────────────────────────────
try:
    GROQ_API_KEY = st.secrets["GROQ_API_KEY"]
except:
    st.error("Missing GROQ_API_KEY in Streamlit Secrets")
    st.stop()

TAVILY_API_KEY = st.secrets.get("TAVILY_API_KEY", None)


# ─────────────────────────────────────────
# INPUTS
# ─────────────────────────────────────────
topic = st.text_input("What do you want to learn?")
level = st.selectbox("Level", ["beginner", "intermediate", "advanced"])
days = st.slider("Days", 3, 30, 14)
hours = st.slider("Hours per day", 1, 8, 2)


# ─────────────────────────────────────────
# FUNCTIONS
# ─────────────────────────────────────────

def fetch_resources(topic, level):
    if not TAVILY_API_KEY:
        return []

    try:
        client = TavilyClient(api_key=TAVILY_API_KEY)
        results = client.search(
            query=f"best resources to learn {topic} for {level}",
            max_results=5
        )
        return [{"title": r["title"], "url": r["url"]} for r in results.get("results", [])]
    except:
        return []


def safe_json_parse(raw):
    raw = re.sub(r'```json|```', '', raw)
    try:
        return json.loads(raw)
    except:
        return None


def generate_roadmap(topic, days, hours, level, resources):
    client = Groq(api_key=GROQ_API_KEY)

    prompt = f"""
Return ONLY valid JSON.

{{
  "title": "Learning Roadmap: {topic}",
  "overview": "Short summary",
  "totalDays": {days},
  "hoursPerDay": {hours},
  "level": "{level}",
  "days": [
    {{
      "day": 1,
      "title": "Topic title",
      "description": "What to study"
    }}
  ]
}}
"""

    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.7
    )

    raw = response.choices[0].message.content.strip()
    data = safe_json_parse(raw)

    if not data:
        st.error("LLM returned invalid JSON")
        st.text(raw)
        st.stop()

    # fallback safety
    data.setdefault("title", f"Learning Roadmap: {topic}")
    data.setdefault("overview", "No overview generated")
    data.setdefault("days", [])

    return data


def generate_pdf(roadmap):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)

    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(roadmap.get("title", ""), styles["Title"]))
    story.append(Spacer(1, 10))

    for day in roadmap.get("days", []):
        story.append(Paragraph(
            f"Day {day.get('day','?')}: {day.get('title','No title')}",
            styles["Heading2"]
        ))
        story.append(Paragraph(day.get("description", ""), styles["Normal"]))
        story.append(Spacer(1, 8))

    doc.build(story)
    buffer.seek(0)
    return buffer


def generate_excel(roadmap):
    buffer = io.BytesIO()
    wb = Workbook()
    ws = wb.active

    ws.append(["Day", "Title", "Description"])

    for d in roadmap.get("days", []):
        ws.append([
            d.get("day", ""),
            d.get("title", ""),
            d.get("description", "")
        ])

    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ─────────────────────────────────────────
# MAIN LOGIC
# ─────────────────────────────────────────
if st.button("Generate Roadmap"):

    if not topic:
        st.error("Please enter a topic")
        st.stop()

    with st.spinner("Fetching resources..."):
        resources = fetch_resources(topic, level)

    with st.spinner("Generating roadmap..."):
        roadmap = generate_roadmap(topic, days, hours, level, resources)

    st.success("Roadmap generated successfully!")

    # DISPLAY
    st.subheader(roadmap.get("title"))
    st.write(roadmap.get("overview"))

    for day in roadmap.get("days", []):
        with st.expander(f"Day {day.get('day','?')} - {day.get('title','No title')}"):
            st.write(day.get("description", ""))

    # DOWNLOAD
    pdf = generate_pdf(roadmap)
    excel = generate_excel(roadmap)

    st.download_button("📄 Download PDF", pdf, "roadmap.pdf")
    st.download_button("📊 Download Excel", excel, "roadmap.xlsx")
