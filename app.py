import streamlit as st
import google.generativeai as genai
from google.generativeai.types import GenerationConfig
import os
from dotenv import load_dotenv
import pypdf
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import time

# ================= 1. SYSTEM CONFIGURATION =================
st.set_page_config(
    page_title="Nova Assistant Pro", 
    page_icon="🌠", 
    layout="wide",
    initial_sidebar_state="expanded"
)
load_dotenv()

# --- CSS Styling for "Elite" Feel ---
st.markdown("""
<style>
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        height: 3em;
        font-weight: bold;
    }
    .reportview-container {
        background: #f0f2f6;
    }
    h1 { color: #1E3A8A; }
    h2 { color: #1E3A8A; font-size: 1.5rem; }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d1fae5;
        color: #065f46;
        border: 1px solid #34d399;
    }
</style>
""", unsafe_allow_html=True)

# --- API Key Setup ---
try:
    if "GOOGLE_API_KEY" in st.secrets:
        api_key = st.secrets["GOOGLE_API_KEY"]
    else:
        api_key = os.getenv("GOOGLE_API_KEY")
    
    if not api_key:
        st.error("🚨 API Key missing. Please set GOOGLE_API_KEY in .env or Secrets.")
        st.stop()
    
    genai.configure(api_key=api_key)
except Exception as e:
    st.error(f"Setup Error: {e}")
    st.stop()

# ================= 2. THE "ELITE" LOGIC PROMPTS (FULL SPEC) =================

# I. PERSONAL STATEMENT LOGIC (The A-N Spec)
FULL_ESSAY_SPEC = """
LOGIC SPEC: Common App Personal Statement (Strict 550–600 Words)

A) NON-NEGOTIABLES
- Target word count: 575 words (Range: 550-600).
- Core requirement: Reveal how the student thinks/chooses, not just what they did.
- Max 2 full scenes total.
- NO Moralizing: Do not use "I learned," "This taught me," "I realized." Show change through behavior.
- NO Motivational Ending: No "change the world," "dream come true."

B) INPUTS & LENS SELECTION
- Define a "Lens": A repeatable pattern of operation.
- Format: "When [trigger] happens, I tend to [instinct], so I [action]."
- Valid only if proven in 2 different contexts.

C) STRUCTURE: THEMATIC NARRATIVE
1. Hook (55-80 words): Lens in action. Immediate motion. No quotes.
2. Scene 1 (170-210 words): Origin moment. Must include TENSION and an A/B DECISION POINT.
   - Format: "I could do A... or B... So I chose B."
3. Scene 2 (170-210 words): Later moment where lens becomes intentional. Must include Internal Processing (cognition, not emotion).
4. Closing (55-80 words): Grounded. How they operate now. "Operating Manual" style.

D) VOICE & STYLE
- Tone: Sparky, simple, precise.
- Uniqueness Device: Choose ONE (System Thinking, Translation, Signal Detection, or Controlled Stubbornness).
- NO generic transitions (Moreover, Furthermore).
- NO Trauma Dumping.

E) SCENE WRITING SPEC
- Setup -> Tension -> Decision Point (A/B) -> Internal Processing -> Action -> Consequence.
- Connection line: Links scene to lens without moralizing.

IMPROVISATION INSTRUCTIONS:
- Context: Uzbekistan High School Student.
- If details are missing, improvise using local context:
  - Zakovat (Intellectual games)
  - Math Olympiads (National obsession)
  - Mahalla (Community duties)
  - Lyceum vs. Public School dynamics
  - Shadow Education (Training centers/Repetitors)
"""

# II. RECOMMENDATION LETTER LOGIC (The A-M Spec)
FULL_REC_SPEC = """
LOGIC SPEC: Recommendation Letter (Counselor + Teachers + Director)

A) NON-NEGOTIABLES
- Authenticity over poetry. Sounds like an adult educator.
- EVIDENCE BASED: Every claim backed by a specific story/measurable detail.
- NO Resume repeating. Select 2-4 main proof points.
- Include 1 "Growth Edge" (small weakness that becomes strength).

B) ROLES
1. SCHOOL DIRECTOR (Counselor Role): Focus on community impact, leadership, maturity, school context.
2. TEACHER (Subject Specific): Focus on intellectual habits, classroom behavior, "how they learn."

C) UZBEKISTAN CONTEXT
- Directors often act as counselors.
- Teachers emphasize discipline, respect, and national curriculum (Prezident maktabi, specialized subjects).
- Mentions of class rank are common and acceptable.

D) STRUCTURE
1. Opening: Credibility + Relationship + One-line summary.
2. Academic Strength: A vivid classroom moment/story.
3. Character/Community: How they elevate peers.
4. Growth Edge: Safe weakness + response.
5. Closing: Strong endorsement.

E) IMPROVISATION ENGINE
- If details missing, generate plausible:
  - Classroom debates (e.g., Alisher Navoi literature analysis, Physics lab on optics).
  - Leadership moments (Class monitor, cleaning day organizer - Hashar).
  - Peer tutoring.
- DO NOT invent specific awards/test scores unless inferred.
"""

# III. ACTIVITIES LOGIC
FULL_ACTIVITIES_SPEC = """
LOGIC SPEC: Common App Activities List
1. Analyze record. If activities < 10, IMPROVISE up to 6 high-quality entries fitting an ambitious Uzbek student.
2. Potential Improvised Activities:
   - English Tutor (Private lessons)
   - Family Business Helper (Bazaar/Shop logistics)
   - Mahalla Volunteer (Community aid)
   - Sports (Football, Kurash, Chess)
   - Telegram Channel Admin (Educational blog)
3. FORMAT:
   - Type (from Common App list)
   - Position/Leadership (Max 50 chars)
   - Organization (Max 100 chars)
   - Description (Max 150 chars, action verbs, impact)
   - Participation (Grades 9-11, Hours/Week estimate)
"""

# ================= 3. HELPER FUNCTIONS =================

def extract_text_from_pdf(file):
    """Robust PDF text extraction."""
    try:
        pdf_reader = pypdf.PdfReader(file)
        text = ""
        for page in pdf_reader.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
        return text if text.strip() else None
    except Exception as e:
        st.error(f"Error reading PDF: {e}")
        return None

def call_gemini(system_prompt, user_content, model_name="gemini-2.5-pro"):
    """Wrapper for Gemini with fallback logic."""
    full_prompt = f"{system_prompt}\n\nSTUDENT RECORD:\n{user_content}"
    
    try:
        # Primary Attempt: Gemini 2.5 Pro (Best Logic)
        model = genai.GenerativeModel("gemini-2.5-pro")
        response = model.generate_content(
            contents=[{"role": "user", "parts": [full_prompt]}],
            generation_config=GenerationConfig(temperature=0.75, max_output_tokens=3000)
        )
        return response.text
    except Exception:
        try:
            # Fallback 1: Gemini 1.5 Pro (Good Logic)
            model = genai.GenerativeModel("gemini-1.5-pro")
            response = model.generate_content(
                contents=[{"role": "user", "parts": [full_prompt]}]
            )
            return response.text
        except Exception as e:
            # Fallback 2: Flash (Speed/Stability)
            try:
                model = genai.GenerativeModel("gemini-1.5-flash")
                response = model.generate_content(
                    contents=[{"role": "user", "parts": [full_prompt]}]
                )
                return response.text
            except:
                return f"Error generation content. Details: {e}"

def create_docx_report(data_dict):
    """Generates a professional DOCX file with all content."""
    doc = Document()
    
    # Style Setup
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(11)
    
    # Title
    head = doc.add_heading(f"Nova Admissions Report: {data_dict.get('name', 'Student')}", 0)
    head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph(f"Generated on: {time.strftime('%Y-%m-%d')}")
    doc.add_paragraph("Confidential Application Materials").italic = True
    doc.add_page_break()

    # SECTION 1: STUDENT ANALYSIS
    doc.add_heading('1. Student Analysis & Strategy', level=1)
    doc.add_paragraph(data_dict.get('analysis', 'N/A'))
    doc.add_page_break()

    # SECTION 2: ACTIVITIES
    doc.add_heading('2. Extracurricular Activities (Common App)', level=1)
    doc.add_paragraph(data_dict.get('activities', 'N/A'))
    doc.add_page_break()

    # SECTION 3: PERSONAL STATEMENT
    doc.add_heading('3. Personal Statement (550-600 Words)', level=1)
    doc.add_paragraph("Note: Ensure formatting is clean (no bolding) before pasting into Common App.").italic = True
    doc.add_paragraph(data_dict.get('essay', 'N/A'))
    doc.add_page_break()

    # SECTION 4: DIRECTOR REC
    doc.add_heading('4. School Director / Counselor Recommendation', level=1)
    doc.add_paragraph(data_dict.get('director_rec', 'N/A'))
    doc.add_page_break()

    # SECTION 5: TEACHER 1 REC
    doc.add_heading('5. Teacher Recommendation 1', level=1)
    doc.add_paragraph(data_dict.get('teacher1_rec', 'N/A'))
    doc.add_page_break()

    # SECTION 6: TEACHER 2 REC
    doc.add_heading('6. Teacher Recommendation 2', level=1)
    doc.add_paragraph(data_dict.get('teacher2_rec', 'N/A'))

    # Save to buffer
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ================= 4. MAIN APPLICATION FLOW =================

def main():
    st.title("🌠 Nova Assistant Pro")
    st.markdown("### The Elite Admissions Generator (Uzbekistan Specialized)")
    
    # --- SESSION STATE INITIALIZATION ---
    if "step" not in st.session_state:
        st.session_state.step = 1
    if "student_text" not in st.session_state:
        st.session_state.student_text = ""
    if "analysis_result" not in st.session_state:
        st.session_state.analysis_result = ""
    if "student_name" not in st.session_state:
        st.session_state.student_name = ""
    if "generated_content" not in st.session_state:
        st.session_state.generated_content = {}

    # --- SIDEBAR ---
    with st.sidebar:
        st.header("Upload Record")
        uploaded_file = st.file_uploader("Select PDF", type=["pdf"])
        if uploaded_file and st.session_state.step == 1:
            if st.button("Start Analysis"):
                with st.spinner("Extracting & analyzing..."):
                    text = extract_text_from_pdf(uploaded_file)
                    if text:
                        st.session_state.student_text = text
                        # Quick Analysis Call
                        analysis_prompt = """
                        Analyze this student record. 
                        1. Identify Name.
                        2. Identify Intended Major (or infer best fit).
                        3. List 3 key strengths.
                        4. Identify gaps that need improvisation.
                        Output format: Plain text.
                        """
                        analysis = call_gemini(analysis_prompt, text)
                        st.session_state.analysis_result = analysis
                        
                        # Try to extract name simply
                        try:
                            lines = analysis.split('\n')
                            st.session_state.student_name = lines[0].replace("Name:", "").strip()
                        except:
                            st.session_state.student_name = "Student"
                            
                        st.session_state.step = 2
                        st.rerun()
                    else:
                        st.error("Could not extract text from PDF. It might be an image scan.")

    # --- MAIN PAGE LOGIC ---

    # STEP 1: INITIAL STATE
    if st.session_state.step == 1:
        st.info("👋 Welcome. Please upload a student PDF in the sidebar to begin.")
        st.markdown("""
        **Capabilities:**
        * **Deep Analysis:** Finds the narrative "Lens".
        * **Essay Generation:** Follows the 14-point Elite Logic Spec.
        * **Improvisation:** Fills gaps with authentic Uzbek context (*Zakovat*, *Mahalla*, etc.).
        * **Full Suite:** Generates Activities List + 3 Distinct Recommendation Letters.
        """)

    # STEP 2: REVIEW & CONFIRM
    elif st.session_state.step == 2:
        st.success("✅ Analysis Complete")
        
        col1, col2 = st.columns([1, 1])
        with col1:
            st.subheader("Student Profile Analysis")
            st.write(st.session_state.analysis_result)
        
        with col2:
            st.subheader("Configuration")
            st.info("Review the analysis on the left. If accurate, proceed to generate the full application suite.")
            
            # User Inputs to guide generation
            intended_major = st.text_input("Confirm Intended Major", placeholder="e.g. Computer Science")
            improvisation_level = st.select_slider("Improvisation Level", options=["Strict (Facts Only)", "Balanced", "Creative (Fill Gaps)"], value="Balanced")
            
            if st.button("🚀 Generate Full Application Suite"):
                st.session_state.config_major = intended_major
                st.session_state.config_improv = improvisation_level
                st.session_state.step = 3
                st.rerun()
            
            if st.button("⬅️ Start Over"):
                st.session_state.clear()
                st.rerun()

    # STEP 3: GENERATION & OUTPUT
    elif st.session_state.step == 3:
        st.subheader(f"generating materials for {st.session_state.student_name}...")
        
        # We run these sequentially to avoid API rate limits and ensure logical flow
        
        # 1. Activities
        if "activities" not in st.session_state.generated_content:
            with st.spinner("1/5 Generating Activities List..."):
                prompt = FULL_ACTIVITIES_SPEC
                if st.session_state.config_major:
                    prompt += f"\nNote: Highlight activities relevant to {st.session_state.config_major}."
                res = call_gemini(prompt, st.session_state.student_text)
                st.session_state.generated_content["activities"] = res
        
        # 2. Personal Statement
        if "essay" not in st.session_state.generated_content:
            with st.spinner("2/5 Writing Personal Statement (Elite Logic)..."):
                prompt = FULL_ESSAY_SPEC
                if st.session_state.config_major:
                    prompt += f"\nNote: The 'Lens' should align with a student interested in {st.session_state.config_major}."
                res = call_gemini(prompt, st.session_state.student_text)
                st.session_state.generated_content["essay"] = res

        # 3. Director Rec
        if "director_rec" not in st.session_state.generated_content:
            with st.spinner("3/5 Drafting Director Recommendation..."):
                prompt = FULL_REC_SPEC + "\nROLE: SCHOOL DIRECTOR / COUNSELOR"
                res = call_gemini(prompt, st.session_state.student_text)
                st.session_state.generated_content["director_rec"] = res

        # 4. Teacher 1 Rec
        if "teacher1_rec" not in st.session_state.generated_content:
            with st.spinner("4/5 Drafting Teacher Rec 1..."):
                prompt = FULL_REC_SPEC + f"\nROLE: TEACHER 1 (Subject aligned with {st.session_state.config_major})"
                res = call_gemini(prompt, st.session_state.student_text)
                st.session_state.generated_content["teacher1_rec"] = res

        # 5. Teacher 2 Rec
        if "teacher2_rec" not in st.session_state.generated_content:
            with st.spinner("5/5 Drafting Teacher Rec 2..."):
                prompt = FULL_REC_SPEC + "\nROLE: TEACHER 2 (Core subject like English/Math/History)"
                res = call_gemini(prompt, st.session_state.student_text)
                st.session_state.generated_content["teacher2_rec"] = res

        # Display Results
        st.success("🎉 All documents generated successfully!")
        
        # Prepare Data for Docx
        final_data = {
            "name": st.session_state.student_name,
            "analysis": st.session_state.analysis_result,
            "activities": st.session_state.generated_content["activities"],
            "essay": st.session_state.generated_content["essay"],
            "director_rec": st.session_state.generated_content["director_rec"],
            "teacher1_rec": st.session_state.generated_content["teacher1_rec"],
            "teacher2_rec": st.session_state.generated_content["teacher2_rec"]
        }
        
        # Tabs for preview
        t1, t2, t3, t4, t5 = st.tabs(["Essay", "Activities", "Director Rec", "Teacher Recs", "Analysis"])
        
        with t1:
            st.markdown("### Personal Statement")
            st.markdown(final_data["essay"])
        with t2:
            st.markdown("### Activities List")
            st.markdown(final_data["activities"])
        with t3:
            st.markdown("### Director Letter")
            st.markdown(final_data["director_rec"])
        with t4:
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("### Teacher 1")
                st.markdown(final_data["teacher1_rec"])
            with c2:
                st.markdown("### Teacher 2")
                st.markdown(final_data["teacher2_rec"])
        with t5:
            st.text(final_data["analysis"])

        # Download Button
        docx_file = create_docx_report(final_data)
        st.download_button(
            label="📄 Download Complete Application Report (.docx)",
            data=docx_file,
            file_name=f"Nova_App_{st.session_state.student_name.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )
        
        if st.button("Start New Student"):
            st.session_state.clear()
            st.rerun()

if __name__ == "__main__":
    main()
