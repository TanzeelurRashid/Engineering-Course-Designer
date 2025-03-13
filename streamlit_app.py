import streamlit as st
import yaml
import os
import google.generativeai as genai
from dotenv import load_dotenv
import json
import time
import random
from google.api_core import exceptions
import pandas as pd
from io import BytesIO
import tempfile

# Import functions from main.py
from main import (
    load_yaml_prompt, 
    generate_course_design,
    exponential_backoff,
    generate_course_design_output,
    generate_excel_summary,
    save_output_to_file
)

# Page configuration
st.set_page_config(
    page_title="Engineering Course Designer",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main {
        background-color: #f9f9f9;
    }
    .stButton > button {
        background-color: #1a5276;
        color: white;
    }
    .stButton > button:hover {
        background-color: #154360;
    }
    .stProgress > div > div {
        background-color: #1a5276;
    }
    .success-message {
        background-color: #d4edda;
        color: #155724;
        padding: 10px;
        border-radius: 4px;
        margin-bottom: 15px;
    }
    .info-box {
        background-color: #e8f4f8;
        border-left: 5px solid #1a5276;
        padding: 10px;
        margin-bottom: 15px;
    }
    h1, h2, h3 {
        color: #1a5276;
    }
</style>
""", unsafe_allow_html=True)

# Load environment variables from .env file (for local development)
load_dotenv()

# Configure the Google API key - Try to get from Streamlit secrets first, then from environment variables
api_key = None

# Check for API key in Streamlit secrets
try:
    if 'GEMINI_API_KEY' in st.secrets:
        api_key = st.secrets['GEMINI_API_KEY']
except:
    # Fall back to environment variables if not in secrets
    api_key = os.environ.get("GEMINI_API_KEY")

if api_key:
    genai.configure(api_key=api_key)
else:
    st.error("""
    No Gemini API key found. Please set your GEMINI_API_KEY in either:
    - Streamlit Cloud secrets
    - Local .env file
    
    You can get an API key from https://ai.google.dev/
    """)
    st.stop()

# Load prompts
@st.cache_data
def load_prompts():
    course_prompt = load_yaml_prompt("course_prompt.yaml")
    curriculum_prompt = load_yaml_prompt("curriculum_prompt.yaml")
    return course_prompt, curriculum_prompt

# Modified save_output function that works with Streamlit Cloud
def save_temp_output(output_string, suffix=".md"):
    """Saves output to a temporary file that works with Streamlit Cloud's read-only filesystem."""
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    temp_path = temp_file.name
    try:
        if suffix == ".md":
            with open(temp_path, 'w', encoding='utf-8') as f:
                f.write(output_string)
        else:
            with open(temp_path, 'wb') as f:
                f.write(output_string)
        return temp_path
    except Exception as e:
        st.error(f"Error saving temporary file: {e}")
        return None

# Modified Excel generation function for Streamlit Cloud
def generate_temp_excel(course_design):
    """Generates an Excel summary to a temporary file compatible with Streamlit Cloud."""
    try:
        # Create Excel data in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Create CLO Summary Sheet
            clo_data = []
            for i, clo in enumerate(course_design["CLOs"], 1):
                clo_data.append({
                    "Sr. No.": i,
                    "CLO": f"CLO-{i}",
                    "CLO Statement": clo["Statement"],
                    "Aligned GA": clo["GA"],
                    "Bloom's Taxonomy": clo["Taxonomy"],
                    "KPIs": ", ".join(clo["KPIs"])
                })
            
            clo_df = pd.DataFrame(clo_data)
            clo_df.to_excel(writer, sheet_name="CLO Summary", index=False)
            
            # Create Course Details Sheet
            details_data = []
            for key, value in course_design["Course Details"].items():
                if isinstance(value, dict):
                    for sub_key, sub_value in value.items():
                        details_data.append({
                            "Category": key,
                            "Item": sub_key,
                            "Value": sub_value
                        })
                else:
                    details_data.append({
                        "Category": "Course Details",
                        "Item": key,
                        "Value": value
                    })
            
            details_df = pd.DataFrame(details_data)
            details_df.to_excel(writer, sheet_name="Course Details", index=False)
            
        # Save to temporary file
        output.seek(0)
        temp_path = save_temp_output(output.getvalue(), suffix=".xlsx")
        return temp_path
    except Exception as e:
        st.error(f"Error generating Excel: {e}")
        return None

# App header
st.title("Engineering Course Designer")
st.markdown("""
<div class="info-box">
Generate professionally designed engineering courses and curriculums with AI assistance.
This tool uses Google's Gemini AI to create detailed course designs with learning outcomes
aligned to engineering graduate attributes.
</div>
""", unsafe_allow_html=True)

# Create tabs
tab1, tab2, tab3 = st.tabs(["Course Design", "Curriculum Design", "About"])

with tab1:
    st.header("Course Design Generator")
    
    with st.form("course_design_form"):
        col1, col2 = st.columns(2)
        
        with col1:
            engineering_domain = st.selectbox(
                "Engineering Domain:",
                ["Mechanical Engineering", "Electrical Engineering", "Computer Engineering", 
                 "Civil Engineering", "Chemical Engineering", "Industrial Engineering", "Aerospace Engineering"]
            )
            
            course_name = st.text_input("Course Name:", placeholder="e.g., Thermodynamics, Machine Design, etc.")
            
            course_type = st.selectbox(
                "Course Type:",
                ["Theory", "Practical", "Theory+Lab"]
            )
            
            semester_placement = st.selectbox(
                "Semester Placement:",
                ["1st Semester", "2nd Semester", "3rd Semester", "4th Semester", 
                 "5th Semester", "6th Semester", "7th Semester", "8th Semester"]
            )
        
        with col2:
            credit_hours = st.selectbox(
                "Credit Hours:",
                ["1", "2", "3", "4"]
            )
            
            prerequisites = st.text_input("Prerequisites:", placeholder="e.g., Calculus I, Physics I, or None")
            
            course_category = st.selectbox(
                "Course Category:",
                ["Core", "Elective"]
            )
            
            course_level = st.selectbox(
                "Course Level:",
                ["Foundation", "Breadth", "Depth"]
            )
            
            knowledge_area = st.selectbox(
                "Knowledge Area:",
                ["Engineering Domain - Engineering Foundation", 
                 "Engineering Domain - Major Based Core Depth",
                 "General Education - Humanities",
                 "General Education - Management Sciences",
                 "General Education - Natural Sciences"]
            )
        
        # Optional notes section
        include_notes = st.checkbox("Include Overall Note")
        
        if include_notes:
            st.markdown("### Overall Notes")
            industry_needs = st.text_area("Industry Needs:", placeholder="Describe relevant industry needs for this course...")
            technologies = st.text_area("Technologies:", placeholder="List relevant technologies for this course...")
            context = st.text_area("Local/National Context:", placeholder="Describe relevant local/national context for this course...")
        
        # Submit button
        submit_button = st.form_submit_button("Generate Course Design")

    # Handle form submission
    if submit_button:
        if not course_name:
            st.error("Course name is required")
        else:
            # Show a progress bar
            progress_bar = st.progress(0)
            status_text = st.empty()
            status_text.text("Initializing course design generation...")
            
            # Prepare inputs
            inputs = {
                "engineering_domain": engineering_domain.split()[0],  # Just take "Mechanical" from "Mechanical Engineering"
                "course_name": course_name,
                "course_type": course_type,
                "semester_placement": semester_placement,
                "credit_hours": credit_hours,
                "prerequisites": prerequisites,
                "course_category": course_category,
                "course_level": course_level,
                "knowledge_area": knowledge_area,
            }
            
            if include_notes:
                inputs["overall_note"] = {
                    "industry_needs": industry_needs,
                    "technologies": technologies,
                    "pakistani_context": context
                }
            
            # Update progress
            progress_bar.progress(25)
            status_text.text("Processing inputs and generating course design...")
            
            # Load course prompt
            course_prompt, _ = load_prompts()
            
            # Generate course design
            try:
                course_design = generate_course_design(inputs, course_prompt)
                progress_bar.progress(75)
                status_text.text("Finalizing course design...")
                
                if course_design is not None:
                    # Display results
                    progress_bar.progress(100)
                    status_text.markdown('<div class="success-message">Course design generated successfully!</div>', unsafe_allow_html=True)
                    
                    # Convert to Markdown for display and save to temp file
                    output_markdown = generate_course_design_output(course_design)
                    
                    # Save output to temporary files for download
                    temp_md_path = save_temp_output(output_markdown)
                    temp_xlsx_path = generate_temp_excel(course_design)
                    
                    # Display in expanders for better organization
                    with st.expander("Course Details", expanded=True):
                        if isinstance(course_design["Course Details"], dict):
                            for key, value in course_design["Course Details"].items():
                                if isinstance(value, dict):
                                    st.write(f"**{key}:**")
                                    st.json(value)
                                else:
                                    st.write(f"**{key}:** {value}")
                        else:
                            st.write(course_design["Course Details"])
                    
                    with st.expander("Course Introduction", expanded=True):
                        st.write(course_design["Course Introduction"])
                    
                    with st.expander("Course Objectives", expanded=True):
                        st.write(course_design["Course Objectives"])
                    
                    with st.expander("Course Learning Outcomes (CLOs)", expanded=True):
                        for clo in course_design["CLOs"]:
                            st.markdown(f"**{clo['CLO']}:** {clo['Statement']} (**{clo['GA']}**, {clo['Taxonomy']})")
                            st.write("**KPIs:**")
                            for kpi in clo["KPIs"]:
                                st.write(f"- {kpi}")
                            st.write("---")
                    
                    with st.expander("Course Contents", expanded=True):
                        st.write(course_design["Course Contents"])
                    
                    with st.expander("Range Classifications", expanded=False):
                        for key, value in course_design["Range Classifications"].items():
                            st.write(f"**{key}:** {value}")
                    
                    with st.expander("SDG Integration", expanded=False):
                        for key, value in course_design["SDG Integration"].items():
                            st.write(f"**{key}:** {value}")
                    
                    with st.expander("Pathway to Professional Competence", expanded=False):
                        st.write(course_design["Pathway to Professional Competence Profiles"])
                    
                    with st.expander("Implementation Guidelines", expanded=False):
                        for key, value in course_design["Implementation Guidelines"].items():
                            st.write(f"**{key}:** {value}")
                    
                    with st.expander("Continuous Improvement Framework", expanded=False):
                        st.write(course_design["Continuous Improvement Framework"])
                    
                    with st.expander("CLO Summary Table", expanded=False):
                        st.write(course_design["CLO Summary Table"])
                    
                    with st.expander("Knowledge Profile", expanded=False):
                        st.write(course_design["Knowledge Profile"])
                    
                    # Create downloadable files
                    st.subheader("Download Course Design")
                    
                    col1, col2, col3 = st.columns(3)
                    
                    # Markdown download
                    with col1:
                        if temp_md_path:
                            with open(temp_md_path, "r", encoding="utf-8") as file:
                                md_content = file.read()
                            st.download_button(
                                label="Download as Markdown",
                                data=md_content,
                                file_name=f"{course_name.replace(' ', '_')}_design.md",
                                mime="text/markdown",
                            )
                    
                    # JSON download
                    with col2:
                        st.download_button(
                            label="Download as JSON",
                            data=json.dumps(course_design, indent=2),
                            file_name=f"{course_name.replace(' ', '_')}_design.json",
                            mime="application/json",
                        )
                    
                    # Excel download
                    with col3:
                        if temp_xlsx_path:
                            try:
                                with open(temp_xlsx_path, "rb") as f:
                                    excel_data = f.read()
                                
                                st.download_button(
                                    label="Download as Excel",
                                    data=excel_data,
                                    file_name=f"{course_name.replace(' ', '_')}_design.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                            except Exception as e:
                                st.warning(f"Could not generate Excel file: {e}")
                
                else:
                    st.error("Failed to generate course design. Please try again.")
            
            except exceptions.ResourceExhausted:
                st.error("The API quota has been exhausted. Please try again later.")
            except Exception as e:
                st.error(f"An error occurred: {str(e)}")

with tab2:
    st.header("Curriculum Design Evaluation")
    
    st.markdown("""
    <div class="info-box">
    This feature will help you evaluate and design a complete curriculum for your engineering program.
    Please provide information about your program to get started.
    </div>
    """, unsafe_allow_html=True)

    with st.form("curriculum_form"):
        # Basic program information
        engineering_domain = st.selectbox(
            "Engineering Domain:",
            ["Mechanical Engineering", "Electrical Engineering", "Computer Engineering", 
             "Civil Engineering", "Chemical Engineering", "Industrial Engineering", "Aerospace Engineering"],
            key="curriculum_domain"
        )
        
        program_name = st.text_input("Program Name:", placeholder="e.g., Bachelor of Science in Mechanical Engineering")
        
        program_duration = st.selectbox(
            "Program Duration:",
            ["4 Years", "5 Years"]
        )
        
        total_credit_hours = st.number_input("Total Credit Hours:", min_value=120, max_value=180, value=136)
        
        # Additional information
        accreditation_body = st.selectbox(
            "Accreditation Body:",
            ["Pakistan Engineering Council (PEC)", "ABET", "Engineering Council UK", "Other"]
        )
        
        if accreditation_body == "Other":
            other_accreditation = st.text_input("Specify Accreditation Body:")
        
        university_context = st.text_area("University Context:", 
                                        placeholder="Describe the university context, mission, vision, etc.")
        
        industry_focus = st.text_area("Industry Focus:", 
                                     placeholder="Describe the industrial focus for this program...")
        
        submit_curriculum = st.form_submit_button("Evaluate Curriculum")
    
    if submit_curriculum:
        st.info("Curriculum evaluation feature is currently in development. Please check back later.")


with tab3:
    st.header("About Engineering Course Designer")
    
    st.markdown("""
    ### Overview
    
    The Engineering Course Designer is an AI-powered tool that helps educators and administrators
    design engineering courses with professionally structured learning outcomes, content, and assessment strategies.
    
    ### Features
    
    - Generate complete course designs for any engineering domain
    - Align learning outcomes with Washington Accord Graduate Attributes
    - Create course content outlines mapped to learning outcomes
    - Generate Key Performance Indicators (KPIs) for each learning outcome
    - Produce downloadable documentation in multiple formats (Markdown, JSON, Excel)
    
    ### How It Works
    
    The tool uses Google's Gemini AI to process your course requirements and generate a comprehensive
    course design document. The AI has been trained on engineering education frameworks, including
    the Washington Accord and program accreditation requirements.
    
    ### Getting Started
    
    1. Fill out the course details form
    2. Click "Generate Course Design"
    3. Review the generated content
    4. Download in your preferred format
    
    ### API Key Setup
    
    This application requires a Google Gemini API key to function. For local deployment, create a `.env` file in the same
    directory as this application with the following content:
    
    ```
    GEMINI_API_KEY=your_api_key_here
    ```
    
    For Streamlit Cloud deployment, add your API key in the Streamlit Cloud secrets management interface.
    
    Replace `your_api_key_here` with your actual API key from [Google AI Studio](https://ai.google.dev/).
    """)

if __name__ == "__main__":
    # This will only execute when the script is run directly
    pass