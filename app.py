import streamlit as st
import google.generativeai as genai
from google.generativeai.types import GenerationConfig
import os
from dotenv import load_dotenv
import pypdf
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import time

# ================= 1. SYSTEM CONFIGURATION =================
st.set_page_config(
    page_title="Nova Assistant: Elite Admissions", 
    page_icon="🌠", 
    layout="wide",
    initial_sidebar_state="expanded"
)
load_dotenv()

# --- CSS Styling ---
st.markdown("""
<style>
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        height: 3em;
        font-weight: bold;
        background-color: #1E3A8A;
        color: white;
    }
    .stButton>button:hover {
        background-color: #2563EB;
        color: white;
    }
    h1 { color: #1E3A8A; }
    h2 { color: #1E3A8A; font-size: 1.5rem; }
    .reportview-container { background: #f0f2f6; }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d1fae5;
        color: #065f46;
        border: 1px solid #34d399;
    }
    .supp-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #eff6ff;
        border: 1px solid #bfdbfe;
        color: #1e3a8a;
    }
</style>
""", unsafe_allow_html=True)

# --- API Key Setup ---
try:
    api_key = st.secrets.get("GOOGLE_API_KEY") or os.getenv("GOOGLE_API_KEY")
    if not api_key:
        st.warning("⚠️ API Key not found. Please set GOOGLE_API_KEY in .env or Streamlit Secrets.")
    else:
        genai.configure(api_key=api_key)
except Exception as e:
    st.error(f"Setup Error: {e}")

# ================= 2. ELITE LOGIC SPECIFICATIONS =================

# I. PERSONAL STATEMENT LOGIC
FULL_ESSAY_SPEC = """
LOGIC SPEC: Common App Personal Statement (Strict 550–600 Words)
A) NON-NEGOTIABLES
- Target word count: 575 words (Range: 550-600).
- Reveal how the student thinks/chooses.
- Max 2 full scenes. NO Moralizing. NO Motivational Endings.
B) LENS: "When [trigger] happens, I tend to [instinct], so I [action]."
C) STRUCTURE: Hook -> Scene 1 (Tension/Decision) -> Scene 2 (Intentionality) -> Closing.
D) VOICE: Sparky, simple, precise. Uniqueness Device: System Thinking, Translation, Signal Detection, or Controlled Stubbornness.
F) CONTEXT: Use Uzbekistan context (Zakovat, Math Olympiads, Mahalla) if details missing.
"""

# II. RECOMMENDATION LOGIC
FULL_REC_SPEC = """
LOGIC SPEC: Recommendations (Director + Teachers)
A) NON-NEGOTIABLES: Evidence-based, No resume repeating, 1 Growth Edge.
B) ROLES: 
   - Director: Community impact, maturity, context.
   - Teacher: Intellectual habits, classroom behavior.
C) CONTEXT: Uzbekistan specifics (Director as counselor, rigorous curriculum).
D) STRUCTURE: Opening -> Academic Strength (Scene) -> Character (Peers) -> Growth Edge -> Closing.
"""

# III. ACTIVITIES LOGIC
FULL_ACTIVITIES_SPEC = """
LOGIC SPEC: Activities List
1. If <10 activities, IMPROVISE up to 6 high-quality Uzbek-specific entries (Tutor, Family Business, Mahalla).
2. FORMAT:
   Activity type: [Type]
   Position/Leadership: [Max 50 chars]
   Organization Name: [Max 100 chars]
   Description: [Max 150 chars - action verbs]
   Participation: [Grades/Timing]
   Hours/Weeks: [Realistic estimates]
"""

# IV. SUPPLEMENTAL ESSAY LOGIC (NEW)
SUPPLEMENTAL_LOGIC_SPEC = """
LOGIC SPEC: SUPPLEMENTAL ESSAYS (Analyze & Execute)

TASK:
1. Analyze the user's prompt to classify it into one of the 8 types below.
2. Use the 'Perfect answer logic' for that specific type.
3. SEARCH THE INTERNET for the specific University to find real labs, professors, values, or traditions.
4. Integrate the Student Record + Generated Personal Statement to ensure consistency.

TYPES & LOGIC:

1) "Why Us?"
- Secretly asking: Do you understand us? Will you use our resources?
- Logic: You (your direction) -> 3 Bridges (School specific detail + Your past + Future use) -> Micro-future.
- Constraint: No generic praise ("great faculty"). Must be specific (e.g., "Professor X's lab").

2) "Why Major?"
- Secretly asking: Is this real? Will you persist?
- Logic: Origin (moment you noticed problem) -> Development (2-3 steps of deeper exposure) -> Current Question -> Tools needed.
- Key: Frame as a question you are chasing.

3) Community
- Secretly asking: Can you belong without copying? What kind of roommate are you?
- Logic: Define community non-basically -> What you did (verbs) -> Impact -> Transfer (how you'll recreate it).

4) Diversity / Background
- Secretly asking: What perspective changes the room? Can you handle complexity?
- Logic: Specific tension you live with -> One scene -> How it shaped habits -> Contribution to peers.
- Constraint: No trauma dumping.

5) Challenge / Failure
- Secretly asking: Do you take responsibility? Do you improve systems?
- Logic: Failure -> Your part (own it) -> Your fix (concrete) -> Proof it worked -> New operating habit.

6) Bridge Builder / Disagreement
- Secretly asking: Can you disagree safely?
- Logic: Opposing view + why it made sense -> Initial reaction -> Pivot (questions/shared goal) -> Outcome -> Future framework.

7) Extracurricular Elaboration
- Secretly asking: What makes your role different?
- Logic: Specific role + problem -> Specific action -> Specific result -> Skill/Value trained.

8) Short Answers
- Secretly asking: Personality and specificity.
- Logic: Quirky interest / Intellectual curiosity / Human detail. Each must be a fingerprint.
"""

# ================= 3. HELPER FUNCTIONS =================

def extract_text_from_pdf(file):
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

def call_gemini(system_prompt, user_content, temperature=0.7, tools=None):
    """Wrapper for Gemini with optional Search Tools."""
    full_prompt = f"{system_prompt}\n\nINPUT DATA:\n{user_content}"
    
    try:
        # Use a model that supports search
        model = genai.GenerativeModel("gemini-2.5-flash-preview-09-2025")
        
        # Configure tools if requested (e.g., Google Search)
        tool_config = tools if tools else None
        
        response = model.generate_content(
            contents=[{"role": "user", "parts": [full_prompt]}],
            generation_config=GenerationConfig(temperature=temperature, max_output_tokens=4000),
            tools=tool_config
        )
        # Handle potential search grounding response structure
        return response.text
    except Exception as e:
        return f"Error generating content: {e}"

def create_docx_report(data_dict, supplemental_list=None):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(11)
    
    # Title
    head = doc.add_heading(f"Nova Admissions Report: {data_dict.get('name', 'Student')}", 0)
    head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Generated on: {time.strftime('%Y-%m-%d')}")
    doc.add_page_break()

    # Main Sections
    sections = [
        ('1. Analysis', 'analysis'),
        ('2. Activities', 'activities'),
        ('3. Personal Statement', 'essay'),
        ('4. Director Rec', 'director_rec'),
        ('5. Teacher 1 Rec', 'teacher1_rec'),
        ('6. Teacher 2 Rec', 'teacher2_rec')
    ]
    
    for title, key in sections:
        doc.add_heading(title, level=1)
        doc.add_paragraph(data_dict.get(key, 'N/A'))
        doc.add_page_break()

    # Supplemental Section
    if supplemental_list:
        doc.add_heading('7. Supplemental Essays', level=1)
        for idx, supp in enumerate(supplemental_list, 1):
            doc.add_heading(f"Supplement {idx}: {supp['university']}", level=2)
            doc.add_paragraph(f"Prompt: {supp['prompt']}")
            doc.add_paragraph(supp['content'])
            doc.add_paragraph("-" * 40)
            doc.add_page_break()

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ================= 4. MAIN APPLICATION FLOW =================

def main():
    st.title("🌠 Nova Assistant")
    st.markdown("### Elite Admissions & Supplemental Strategist")
    
    # --- SESSION STATE ---
    if "step" not in st.session_state: st.session_state.step = 1
    if "student_text" not in st.session_state: st.session_state.student_text = ""
    if "analysis_result" not in st.session_state: st.session_state.analysis_result = ""
    if "student_name" not in st.session_state: st.session_state.student_name = "Student"
    if "generated_content" not in st.session_state: st.session_state.generated_content = {}
    if "supplementals" not in st.session_state: st.session_state.supplementals = []

    # --- SIDEBAR ---
    with st.sidebar:
        st.header("1. Upload Record")
        uploaded_file = st.file_uploader("Student PDF", type=["pdf"])
        if uploaded_file and st.session_state.step == 1:
            if st.button("Start Processing"):
                with st.spinner("Analyzing..."):
                    text = extract_text_from_pdf(uploaded_file)
                    if text:
                        st.session_state.student_text = text
                        analysis_prompt = "Analyze student record. Identify Name, Major, Teachers, Director."
                        analysis = call_gemini(analysis_prompt, text)
                        st.session_state.analysis_result = analysis
                        try:
                            st.session_state.student_name = analysis.split('\n')[0].split(":")[1].strip()
                        except: pass
                        st.session_state.step = 2
                        st.rerun()

    # STEP 1: WELCOME
    if st.session_state.step == 1:
        st.info("👋 Upload a student PDF record (Uzbekistan Format) to begin.")

    # STEP 2: REVIEW & CONFIG
    elif st.session_state.step == 2:
        st.success("✅ Record Analyzed")
        st.write(st.session_state.analysis_result)
        st.info("Ready to generate Core Application (Essay, Activities, Recs).")
        improv_mode = st.radio("Improvisation Mode", ["Balanced", "Creative"])
        
        if st.button("🚀 Generate Core Application"):
            st.session_state.improv_mode = improv_mode
            st.session_state.step = 3
            st.rerun()

    # STEP 3 & 4: CORE OUTPUT + SUPPLEMENTAL GENERATOR
    elif st.session_state.step >= 3:
        # --- CORE GENERATION (If not done) ---
        if not st.session_state.generated_content:
            st.subheader(f"Generating Core Docs for {st.session_state.student_name}...")
            
            # Activities
            with st.spinner("1/4 Activities..."):
                prompt = FULL_ACTIVITIES_SPEC
                if st.session_state.improv_mode == "Creative": prompt += "\nNOTE: Improvise 6 entries."
                st.session_state.generated_content["activities"] = call_gemini(prompt, st.session_state.student_text)
            
            # Essay
            with st.spinner("2/4 Personal Statement..."):
                st.session_state.generated_content["essay"] = call_gemini(FULL_ESSAY_SPEC, st.session_state.student_text, temperature=0.8)
            
            # Recs
            with st.spinner("3/4 Recommendations..."):
                st.session_state.generated_content["director_rec"] = call_gemini(FULL_REC_SPEC + "\nROLE: DIRECTOR", st.session_state.student_text)
                st.session_state.generated_content["teacher1_rec"] = call_gemini(FULL_REC_SPEC + "\nROLE: TEACHER 1", st.session_state.student_text)
                st.session_state.generated_content["teacher2_rec"] = call_gemini(FULL_REC_SPEC + "\nROLE: TEACHER 2", st.session_state.student_text)
            
            st.rerun()

        # --- DISPLAY CORE TABS ---
        st.success("🎉 Core Application Generated")
        core_tabs = st.tabs(["Essay", "Activities", "Director Rec", "Teacher 1", "Teacher 2"])
        with core_tabs[0]: st.markdown(st.session_state.generated_content["essay"])
        with core_tabs[1]: st.text(st.session_state.generated_content["activities"])
        with core_tabs[2]: st.markdown(st.session_state.generated_content["director_rec"])
        with core_tabs[3]: st.markdown(st.session_state.generated_content["teacher1_rec"])
        with core_tabs[4]: st.markdown(st.session_state.generated_content["teacher2_rec"])
        
        st.markdown("---")

        # ================= STEP 4: SUPPLEMENTAL ESSAY GENERATOR =================
        st.markdown("## 📚 Supplemental Essay Generator")
        st.markdown("""
        <div class="supp-box">
        <b>Advanced Feature:</b> This tool uses Google Search to find real details about the university 
        and combines them with your student profile + generated personal statement.
        </div>
        """, unsafe_allow_html=True)
        
        c1, c2, c3 = st.columns([2, 3, 1])
        with c1:
            univ_name = st.text_input("University Name", placeholder="e.g. Duke University")
        with c2:
            supp_prompt = st.text_area("Essay Prompt", placeholder="e.g. Why do you want to study here?")
        with c3:
            word_limit = st.text_input("Word Limit", placeholder="e.g. 250 words")

        if st.button("✨ Research & Generate Supplemental Essay"):
            if not univ_name or not supp_prompt:
                st.error("Please provide University Name and Prompt.")
            else:
                with st.spinner(f"🔍 Researching {univ_name} & analyzing profile..."):
                    # Construct Context
                    context_data = f"""
                    STUDENT RECORD: {st.session_state.student_text}
                    
                    ALREADY GENERATED PERSONAL STATEMENT (Do not contradict, but build bridges):
                    {st.session_state.generated_content["essay"]}
                    
                    ALREADY GENERATED ACTIVITIES:
                    {st.session_state.generated_content["activities"]}
                    
                    TARGET UNIVERSITY: {univ_name}
                    ESSAY PROMPT: {supp_prompt}
                    WORD LIMIT: {word_limit}
                    """
                    
                    # Generate with Search Tool
                    supp_result = call_gemini(
                        system_prompt=SUPPLEMENTAL_LOGIC_SPEC, 
                        user_content=context_data,
                        tools='google_search' # Enable Internet Access
                    )
                    
                    # Save result
                    st.session_state.supplementals.append({
                        "university": univ_name,
                        "prompt": supp_prompt,
                        "content": supp_result
                    })
                    st.rerun()

        # Display Generated Supplementals
        if st.session_state.supplementals:
            st.subheader("Generated Supplementals")
            for i, supp in enumerate(st.session_state.supplementals):
                with st.expander(f"{supp['university']} - {supp['prompt'][:50]}...", expanded=True):
                    st.markdown(supp['content'])
                    if st.button(f"Delete Essay {i+1}", key=f"del_{i}"):
                        st.session_state.supplementals.pop(i)
                        st.rerun()

        # FINAL DOWNLOAD
        final_data = {
            "name": st.session_state.student_name,
            "analysis": st.session_state.analysis_result,
            "activities": st.session_state.generated_content["activities"],
            "essay": st.session_state.generated_content["essay"],
            "director_rec": st.session_state.generated_content["director_rec"],
            "teacher1_rec": st.session_state.generated_content["teacher1_rec"],
            "teacher2_rec": st.session_state.generated_content["teacher2_rec"]
        }
        
        docx_file = create_docx_report(final_data, st.session_state.supplementals)
        st.download_button(
            label="📄 Download Complete Application Package (Core + Supplementals)",
            data=docx_file,
            file_name=f"Nova_App_Package_{st.session_state.student_name.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )
        
        if st.button("Start New Student"):
            st.session_state.clear()
            st.rerun()

if __name__ == "__main__":
    main()
