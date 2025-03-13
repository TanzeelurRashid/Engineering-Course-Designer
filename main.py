import yaml
import os
import google.generativeai as genai
from dotenv import load_dotenv
import re
import time
import random
from google.api_core import exceptions
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, Border, Side
import subprocess
import platform

# Load environment variables from .env file
load_dotenv()

# Configure the API key
genai.configure(api_key=os.environ.get("GEMINI_API_KEY"))

# Create a Gemini model instance
model = genai.GenerativeModel(
    "gemini-2.0-flash-lite",  # Updated model name
    generation_config=genai.types.GenerationConfig(temperature=0.2),
)

# --- Shell Command Integration ---
def execute_shell_command(command, capture_output=True):
    """
    Execute a shell command and return its output.
    Args:
        command: The command to execute (string or list)
        capture_output: Whether to capture and return the command output
    Returns:
        tuple: (return_code, stdout, stderr)
    """
    try:
        # Use shell=True for string commands, shell=False for list commands
        is_string_command = isinstance(command, str)
        
        # For Windows, force cmd.exe for string commands
        if platform.system() == 'Windows' and is_string_command:
            process = subprocess.run(command, 
                                  shell=True, 
                                  capture_output=capture_output,
                                  text=True,
                                  env=os.environ)
        else:
            process = subprocess.run(command if is_string_command else command.split(),
                                  shell=is_string_command,
                                  capture_output=capture_output,
                                  text=True,
                                  env=os.environ)
        
        return process.returncode, process.stdout, process.stderr
    except subprocess.SubprocessError as e:
        return -1, "", str(e)

def run_interactive_command(command):
    """
    Run an interactive shell command that requires user input.
    Args:
        command: The command to execute (string)
    """
    try:
        subprocess.run(command, shell=True, check=True)
    except subprocess.SubprocessError as e:
        print(f"Error executing command: {e}")

def get_shell_type():
    """
    Determine the type of shell being used.
    Returns:
        str: Name of the shell (cmd, powershell, bash, etc.)
    """
    if platform.system() == 'Windows':
        # Check if PowerShell is being used
        if 'powershell' in os.environ.get('PSModulePath', '').lower():
            return 'powershell'
        return 'cmd'
    return os.environ.get('SHELL', 'bash')

# --- Helper Functions ---
def load_yaml_prompt(filepath):
    """Loads a YAML prompt from a file."""
    try:
        with open(filepath, "r") as f:
            prompt = yaml.safe_load(f)
        return prompt
    except FileNotFoundError:
        print(f"Error: Prompt file not found at {filepath}")
        return None
    except yaml.YAMLError as e:
        print(f"Error parsing YAML file {filepath}: {e}")
        return None

def get_user_input_for_curriculum(curriculum_prompt):
    """Collects user input for the holistic curriculum design."""
    inputs = {}
    if curriculum_prompt is None:
        print("Error: curriculum_prompt is None. Cannot proceed.")
        return None
    for section_name, section_data in curriculum_prompt["part_1"]["questions"].items():
        if section_name == "engineering_domain":
            response = input(f"{section_data['text']}: ")
            inputs[section_name] = response
        elif section_data.get("optional", False):
            response = input(f"{section_data['text']} (Optional): ")
            inputs[section_name] = response
        else:
            response = input(f"{section_data['text']}: ")
            inputs[section_name] = response
    return inputs

def get_user_input_for_course(course_prompt):
    """Collects user input for a single course design."""
    inputs = {}
    if course_prompt is None:
        print("Error: course_prompt is None. Cannot proceed.")
        return None
    inputs["engineering_domain"] = input(
        "Enter the engineering domain (e.g., Mechanical, Electrical, Computer): "
    )
    inputs["course_type"] = input(
        f"{course_prompt['course_type']['text']} ({'/'.join(course_prompt['course_type']['options'])}): "
    )
    inputs["course_name"] = input(f"{course_prompt['course_name']['text']}: ")
    inputs["semester_placement"] = input(
        f"{course_prompt['semester_placement']['text']}: "
    )
    inputs["credit_hours"] = input(f"{course_prompt['credit_hours']['text']}: ")
    inputs["prerequisites"] = input(f"{course_prompt['prerequisites']['text']}: ")

    inputs["course_category"] = input("Is this a Core or Elective course? ")
    inputs["course_level"] = input(
        "Is this a Foundation, Breadth, or Depth course? "
    )
    inputs["knowledge_area"] = input(
        "To which Knowledge Area does this course belong (e.g., General Education - Humanities, Engineering Domain - Major Based Core Depth)? "
    )

    # Handle Overall Note (Optional Questions)
    if input("Do you want to provide an Overall Note (y/n)? ").lower() == "y":
        inputs["overall_note"] = {}
        for question in course_prompt["overall_note"]["questions"]:
            question_name = question["name"]
            if question.get("optional", False):
                response = input(f"{question['text']} (Optional): ")
            else:
                response = input(f"{question['text']}: ")
            inputs["overall_note"][question_name] = response

    return inputs

# --- Course Design Logic ---

def generate_course_design(inputs, course_prompt):
    """Generates a course design based on user inputs and the course prompt."""
    engineering_domain = inputs["engineering_domain"]

    # 1. Contextual Analysis and Prioritization
    prioritized_gas = exponential_backoff(
        prioritize_plos, inputs["course_name"], inputs, engineering_domain
    )

    # Ensure prioritized_gas is a list before passing it to other functions
    if prioritized_gas is None:
        prioritized_gas = []
        print("Warning: prioritize_plos returned None. Using an empty list for prioritized GAs.")
    elif isinstance(prioritized_gas, str):
        prioritized_gas = [ga.strip() for ga in prioritized_gas.split(",")]
    elif not isinstance(prioritized_gas, list):
        prioritized_gas = []
        print(f"Warning: prioritize_plos did not return a valid list. Returned: {type(prioritized_gas)}. Using an empty list.")

    # 2. Comprehensive Course Design Generation
    course_design = {}

    # A. Course Details Table
    course_design["Course Details"] = {
        "Course Title": inputs["course_name"],
        "Credit Hours": inputs["credit_hours"],
        "Prerequisites": inputs["prerequisites"],
        "Weekly Contact Hours": calculate_contact_hours(inputs["credit_hours"], inputs["course_type"]),
        "Assessment Distribution": {
            "Midterm Exam": "25%",
            "Final Exam": "50%",
            "Sessionals": "25%"
        },
        "Course Category": inputs["course_category"],
        "Course Level": inputs["course_level"],
        "Primary SDG Alignment": exponential_backoff(
            select_sdgs, inputs["course_name"], inputs
        ),
        "Recommended Books": exponential_backoff(get_recommended_books, inputs["course_name"]),
        "Reference Materials": exponential_backoff(get_reference_materials, inputs["course_name"]),
        "Knowledge Area(s) and Subcategory": inputs["knowledge_area"],
    }

    # B. Course Introduction
    course_design["Course Introduction"] = exponential_backoff(
        generate_course_introduction, inputs, prioritized_gas, engineering_domain
    )

    # C. Course Objectives
    course_design["Course Objectives"] = exponential_backoff(
        generate_course_objectives, engineering_domain, inputs["course_name"]
    )

    # D. Course Learning Outcomes (CLOs) with Mappings
    clos_list = exponential_backoff(
        generate_clos, prioritized_gas, engineering_domain, inputs["course_name"]
    )

    if clos_list is not None:
        course_design["CLOs"] = clos_list
    else:
        print("Error generating CLOs. Setting CLOs to an empty list.")
        course_design["CLOs"] = []

    # E. Course Contents
    course_design["Course Contents"] = exponential_backoff(
        generate_course_content,
        inputs["course_name"],
        course_design["CLOs"],
        engineering_domain,
    )

    # F. Range Classifications
    course_design["Range Classifications"] = exponential_backoff(
        generate_range_classifications, engineering_domain
    )

    # G. SDG Integration
    course_design["SDG Integration"] = exponential_backoff(
        generate_sdg_integration,
        course_design["Course Details"]["Primary SDG Alignment"],
        engineering_domain,
        inputs["course_name"],
    )

    # H. Pathway to Professional Competence Profiles
    course_design["Pathway to Professional Competence Profiles"] = (
        exponential_backoff(
            generate_competence_profiles_pathway, prioritized_gas, engineering_domain
        )
    )

    # I. Implementation Guidelines
    course_design["Implementation Guidelines"] = exponential_backoff(
        generate_implementation_guidelines, engineering_domain
    )

    # J. Continuous Improvement Framework
    course_design["Continuous Improvement Framework"] = (
        exponential_backoff(
            generate_continuous_improvement_framework, engineering_domain
        )
    )

    # K. CLO Summary Table
    course_design["CLO Summary Table"] = generate_clo_summary_table(
        course_design["CLOs"],
        inputs["course_name"],
        inputs["credit_hours"],
        course_design["Course Contents"],
    )

    # L. Knowledge Profile
    course_design["Knowledge Profile"] = exponential_backoff(
        generate_knowledge_profile, engineering_domain, prioritized_gas, inputs
    )

    return course_design

# --- Functions for Course Design Logic ---
def exponential_backoff(func, *args, max_retries=5, base_delay=1, **kwargs):
    retries = 0
    while retries < max_retries:
        try:
            result = func(*args, **kwargs)
            if result == "Error generating content":
                raise ValueError("Function returned an error message.")
            return result  # Success, return the result
        except exceptions.ResourceExhausted as e:
            retries += 1
            delay = base_delay * (2**retries) + random.uniform(0, 1)  # Exponential backoff with jitter
            print(
                f"Resource exhausted on attempt {retries}/{max_retries}. Waiting for {delay:.2f} seconds before retrying."
            )
            time.sleep(delay)
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            return None  # Handle other exceptions as needed

    print(f"Failed after {max_retries} attempts.")
    return None  # Indicate failure after retries

def prioritize_plos(course_name, inputs, engineering_domain):
    """Prioritizes PLOs/GAs based on course name, user inputs, and engineering domain."""
    gas_list = """
    **GA-1:** Engineering Knowledge
    **GA-2:** Problem Analysis
    **GA-3:** Design/Development of Solutions
    **GA-4:** Investigation
    **GA-5:** Modern Tool Usage
    **GA-6:** The Engineer and the World
    **GA-7:** Ethics
    **GA-8:** Individual and Team Work
    **GA-9:** Communication
    **GA-10:** Project Management and Finance
    **GA-11:** Lifelong Learning
    """
    prompt = f"""
    You are an expert in {engineering_domain} curriculum design, knowledgeable about the Washington Accord's Graduate Attributes (GAs) and the Pakistan Engineering Council's (PEC) Program Learning Outcomes (PLOs), or the relevant accrediting body for the specified domain.

    A course named '{course_name}' is being designed for a Bachelor of Science in {engineering_domain} program.

    Based on the course name and the following user inputs:

    {inputs}

    Prioritize 2-3 GAs from the list below that are most relevant to this course. Consider the PEC's framework (or the relevant accrediting body) and the Washington Accord when making your selection.

    {gas_list}

    Provide only the prioritized GAs in a comma-separated list with their full names along with mapping to GAs (e.g., GA-1: Engineering Knowledge, GA-3: Design/Development of Solutions).
    """
    response = model.generate_content(prompt)
    # Extract GA codes from the response
    ga_codes = re.findall(r"GA-\d+", response.text)

    # Map GA codes to full names using the provided list
    prioritized_gas = []
    for ga_code in ga_codes:
        for line in gas_list.split("\n"):
            if ga_code in line:
                # Extract full GA name using regular expression
                match = re.search(r"\*\*([A-Z]+-\d+):\*\* (.*)", line)
                if match:
                    code = match.group(1)
                    name = match.group(2)
                    prioritized_gas.append(f"{code}: {name}")
                break  # Stop searching once the GA is found

    return prioritized_gas

def calculate_contact_hours(credit_hours, course_type):
    """Calculates weekly contact hours based on credit hours and course type"""
    if course_type == "Theory":
        return credit_hours
    else:
        return credit_hours * 3

def select_sdgs(course_name, inputs):
    """Selects relevant SDGs based on the course name and user inputs using the Gemini API."""
    prompt = f"""
    You are an expert in curriculum design with a focus on the UN Sustainable Development Goals (SDGs).

    For a course titled '{course_name}', considering the nature of the course and any specific inputs provided by the user, select 1-2 most relevant SDGs from the list below and provide a brief justification for each selection.

    **SDG-1:** No Poverty
    **SDG-2:** Zero Hunger
    **SDG-3:** Good Health and Well-being
    **SDG-4:** Quality Education
    **SDG-5:** Gender Equality
    **SDG-6:** Clean Water and Sanitation
    **SDG-7:** Affordable and Clean Energy
    **SDG-8:** Decent Work and Economic Growth
    **SDG-9:** Industry, Innovation, and Infrastructure
    **SDG-10:** Reduced Inequality
    **SDG-11:** Sustainable Cities and Communities
    **SDG-12:** Responsible Consumption and Production
    **SDG-13:** Climate Action
    **SDG-14:** Life Below Water
    **SDG-15:** Life On Land
    **SDG-16:** Peace, Justice, and Strong Institutions
    **SDG-17:** Partnerships to achieve the Goal

    Provide your response in the following format:

    SDGs: SDG-X, SDG-Y
    Justification: [Provide a brief justification for each SDG, separated by a semicolon.]
    """
    response = model.generate_content(prompt)
    return response.text.strip()

def get_recommended_books(course_name):
    """Provides a list of recommended books based on course name using the Gemini API."""
    prompt = f"""
    You are an expert in curriculum design.

    For a course titled '{course_name}', suggest 3 recommended books that are relevant to the subject matter.

    Provide the recommended books in the following format:

    1. "Book Title 1" by Author 1.
    2. "Book Title 2" by Author 2.
    3. "Book Title 3" by Author 3.
    """
    response = model.generate_content(prompt)
    return response.text.strip()

def get_reference_materials(course_name):
    """Provides a list of reference materials based on course name using the Gemini API."""
    prompt = f"""
    You are an expert in curriculum design.

    For a course titled '{course_name}', suggest 2-3 relevant reference materials (e.g., journals, websites, publications).

    Provide the reference materials in a comma-separated list (e.g., Reference 1, Reference 2, Reference 3).
    """
    response = model.generate_content(prompt)
    return response.text.strip()

def determine_course_category(course_name):
    """Determines course category based on the PEC framework"""
    # Placeholder logic
    return "Core"

def determine_knowledge_area(course_name):
    """Determines the knowledge area and subcategory based on course name"""
    # Placeholder logic
    if course_name == "Thermodynamics":
        return "Engineering Domain - Engineering Foundation"
    elif course_name == "Machine Design":
        return "Engineering Domain - Major Based Core Depth"
    elif course_name == "Technical Writing":
        return "General Education - Humanities"
    elif course_name == "Engineering Economics":
        return "General Education - Management Sciences"
    else:
        return "Engineering Domain - Core"

def generate_course_introduction(inputs, prioritized_gas, engineering_domain):
    """Generates a course introduction using the Gemini API."""
    gas_formatted = ", ".join(prioritized_gas)
    prompt = f"""
    You are an expert in {engineering_domain} curriculum design, knowledgeable about the Washington Accord and the Pakistan Engineering Council (PEC) guidelines (or the relevant accrediting body for this domain).

    Write a concise and informative introduction (around 200-250 words) for a course called '{inputs['course_name']}'.

    This is a {inputs['credit_hours']}-credit hour {inputs['course_type']} course offered in the {inputs['semester_placement']} of the B.Sc. {engineering_domain} program.

    The course aligns with the following GAs: {gas_formatted}.

    Consider the following information from the user's overall note (if provided):

    *   Industry Focus: {inputs.get('overall_note', {}).get('industry_needs', 'N/A')}
    *   Technologies: {inputs.get('overall_note', {}).get('technologies', 'N/A')}
    *   Context: {inputs.get('overall_note', {}).get('pakistani_context', 'N/A')}

    Make the introduction engaging, relevant to undergraduate {engineering_domain} students, and adaptable to different institutional contexts. Briefly mention how this course contributes to their professional development and the capstone project or experience relevant to this domain.
    """

    response = model.generate_content(prompt)
    return response.text

def generate_course_objectives(engineering_domain, course_name):
    """Generates course objectives using the Gemini API."""
    prompt = f"""
    You are an expert in {engineering_domain} curriculum design.

    Generate 3-5 course objectives for a course named '{course_name}'. The objectives should be:

    *   Clearly stated and relevant to the course content.
    *   Actionable and describe what students will be able to do upon completing the course.
    *   Written in a general way that is adaptable to different institutional contexts.

    Provide the course objectives as a numbered list.
    """
    response = model.generate_content(prompt)
    return response.text

def generate_clos(prioritized_gas, engineering_domain, course_name):
    """Generates CLOs using the Gemini API and maps them to prioritized GAs."""
    # Ensure prioritized_gas is a list before the loop
    if isinstance(prioritized_gas, str):
        prioritized_gas = [ga.strip() for ga in prioritized_gas.split(",")]

    gas_with_mapping = []
    for ga in prioritized_gas:
        parts = ga.split(":");
        if len(parts) == 2:
            ga_code = parts[0].strip()
            gas_with_mapping.append(f"**{ga_code.strip()}**: {parts[1].strip()}")
        else:
            print(f"Warning: Invalid GA format: {ga}")
            continue  # Skip this GA and move to the next one

    gas_formatted = "\n".join(gas_with_mapping)

    prompt = f"""
    You are an expert in {engineering_domain} curriculum design, knowledgeable about the Washington Accord's Graduate Attributes (GAs) and the Pakistan Engineering Council's (PEC) Program Learning Outcomes (PLOs), or the relevant accrediting body for the specified domain.

    You are provided with the name of the course and a list of prioritized GAs for that course:

    **Course Name:** {course_name}

    **Prioritized GAs with Descriptions:**
    {gas_formatted}

    Generate a maximum of 4 measurable Course Learning Outcomes (CLOs) for this course. Each CLO should:

    *   Align with one specific prioritized GA from the provided list. Ensure that each given GA is used for one CLO only.
    *   Use action verbs that describe observable and measurable behavior (refer to Bloom's Taxonomy).
    *   Be specific and clearly state what students will be able to do upon completing the course.
    *   Be written in a way that is adaptable to different institutional contexts.

    Provide the CLOs in the following format:

    **CLO-1:** [CLO Statement] (**GA-Y**, [Bloom's Taxonomy Level])
    **CLO-2:** [CLO Statement] (**GA-Y**, [Bloom's Taxonomy Level])
    **CLO-3:** [CLO Statement] (**GA-Y**, [Bloom's Taxonomy Level])
    **CLO-4:** [CLO Statement] (**GA-Y**, [Bloom's Taxonomy Level])
    """
    response = model.generate_content(prompt)
    return parse_clos_from_response(response.text)

def parse_clos_from_response(response_text):
    """Parses the response text from the Gemini API to extract CLOs."""
    clos = []
    lines = response_text.split('\n')
    for line in lines:
        if line.strip().startswith("**CLO"):
            parts = line.split(":** ")
            if len(parts) == 2:
                clo_label = parts[0].strip("* ")
                statement_part = parts[1].strip()
                # Find the start and end of the parenthetical part
                start_paren = statement_part.find("(")
                end_paren = statement_part.find(")")
                if start_paren != -1 and end_paren != -1:
                    statement = statement_part[:start_paren].strip()
                    mapping = statement_part[start_paren + 1:end_paren].strip()
                    # Split the mapping part into GA and Taxonomy
                    if ", " in mapping:
                        ga, taxonomy = mapping.split(", ", 1)
                    else:
                        ga = mapping
                        taxonomy = "N/A"  # Default value

                    # Remove any trailing asterisks
                    statement = statement.replace("*", "").strip()
                    ga = ga.replace("*", "").strip()
                    taxonomy = taxonomy.replace("*", "").strip()

                    clos.append(
                        {
                            "CLO": clo_label,
                            "Statement": statement,
                            "GA": ga,
                            "Taxonomy": taxonomy,
                            "KPIs": generate_kpis(statement, taxonomy),
                        }
                    )  # Generate KPIs
    return clos

def generate_kpis(statement, taxonomy):
    """Generates KPIs for a CLO statement using the Gemini API."""
    prompt = f"""
    Generate 2-3 Key Performance Indicators (KPIs) for the following Course Learning Outcome (CLO) statement:

    **CLO Statement:** {statement}
    **Bloom's Taxonomy Level:** {taxonomy}

    The KPIs should be based on the 'SMART' principle, which means that they should be Specific, Measurable, Achievable, Relevant, and Time-bound.

    Provide the KPIs in a comma-separated list (e.g., KPI-1, KPI-2, KPI-3).
    """
    response = model.generate_content(prompt)
    kpis = [kpi.strip() for kpi in response.text.split(",")]
    return kpis

def generate_course_content(course_name, clos, engineering_domain):
    """Generates a basic course content outline using the Gemini API."""
    clos_formatted = ""
    for i, clo in enumerate(clos):
        clos_formatted += f"  **CLO-{i + 1}:** {clo['Statement']} ({clo['GA']}, {clo['Taxonomy']})\n"

    prompt = f"""
    You are an expert in {engineering_domain} curriculum design. You are provided with the course name and a list of its Course Learning Outcomes (CLOs):

    **Course Name:** {course_name}

    **CLOs:**
    {clos_formatted}

    Generate a basic course content outline for this course. The outline should include:

    *   **Modules or Topics:** Organize the content into modules or topics.
    *   **Module/Topic Titles:** Provide clear and descriptive titles for each module or topic.
    *   **Alignment with CLOs:** Indicate which CLOs are addressed by each module or topic.

    Provide the course content outline in the following format:

    **Module 1:** [Module Title] (CLOs: CLO-1, CLO-2)
    **Module 2:** [Module Title] (CLOs: CLO-3, CLO-4)
    ...
    """
    response = model.generate_content(prompt)
    return response.text

def generate_range_classifications(engineering_domain):
    """Generates a brief description of the Problem Solving and Engineering Activities Ranges using the Gemini API."""
    prompt = f"""
    You are an expert in {engineering_domain} curriculum design, familiar with the Washington Accord's definitions of Problem Solving Ranges (WP1-WP7) and Engineering Activities Ranges (EA1-EA5).

    Considering a typical course in a B.Sc. {engineering_domain} program, provide a brief description (1-2 sentences each) for:

    1. **Problem Solving Range:**  What is the typical range of problem-solving complexity (WP1-WP7) that students are expected to handle in this course?
    2. **Engineering Activities Range:** What is the typical range of engineering activities (EA1-EA5) that students are expected to engage in during this course?

    Provide the descriptions in the following format:

    **Problem Solving Range:** [Description]
    **Engineering Activities Range:** [Description]
    """
    response = model.generate_content(prompt)
    classifications = {}
    lines = response.text.split("\n")
    for line in lines:
        if line.startswith("**Problem Solving Range:**"):
            classifications["Problem Solving Range"] = line.replace("**Problem Solving Range:**", "").strip()
        elif line.startswith("**Engineering Activities Range:**"):
            classifications["Engineering Activities Range"] = line.replace("**Engineering Activities Range:**", "").strip()

    return classifications

def generate_sdg_integration(sdgs, engineering_domain, course_name):
    """Generates a description of SDG integration using the Gemini API."""

    if isinstance(sdgs, str):
        sdgs = [s.strip() for s in sdgs.split(",")]
    sdgs = [s for s in sdgs if s]

    prompt = f"""
    You are an expert in {engineering_domain} curriculum design with a focus on sustainable development.

    For the course '{course_name}', provide a description (around 150-200 words) of how this course integrates the UN Sustainable Development Goals (SDGs), including:

    *   **Justification:** Briefly explain how the course content relates to the chosen SDGs.
    *   **Integration Points:** Provide specific examples of modules, topics, or assignments where SDG considerations are integrated into the course.
    *   **Emphasis:** Describe how the course emphasizes sustainable engineering practices.

    Provide the description in the following format:

    **Justification:** [Justification text]
    **Integration Points:** [Integration Points text]
    **Emphasis:** [Emphasis text]
    """
    response = model.generate_content(prompt)
    sdg_data = {}
    lines = response.text.split("\n")
    for line in lines:
        if line.startswith("**Justification:**"):
            sdg_data["Justification"] = line.replace("**Justification:**", "").strip()
        elif line.startswith("**Integration Points:**"):
            sdg_data["Integration Points"] = line.replace("**Integration Points:**", "").strip()
        elif line.startswith("**Emphasis:**"):
            sdg_data["Emphasis"] = line.replace("**Emphasis:**", "").strip()

    return sdg_data

def generate_competence_profiles_pathway(prioritized_gas, engineering_domain):
    """
    Generates a brief overview of how the course contributes to competence profiles using the Gemini API.
    Restricts the competence profiles to a maximum of 5.
    """
    gas_formatted = ", ".join(prioritized_gas)
    prompt = f"""
    You are an expert in {engineering_domain} curriculum design, familiar with the Washington Accord's Professional Competence Profiles for Professional Engineers.

    Provide a brief overview (around 150-200 words) of how this course contributes to the development of the 16 Professional Competence Profiles defined by the Washington Accord. Specifically, mention which competence profiles are most strongly addressed in this course (limit to a maximum of 4 profiles).

    Refer to the competence profiles using their codes (EC1-EC16).
    """
    response = model.generate_content(prompt)
    return response.text

def generate_implementation_guidelines(engineering_domain):
    """Generates general implementation guidelines using the Gemini API."""
    prompt = f"""
    You are an expert in {engineering_domain} curriculum design and pedagogy.

    Provide general implementation guidelines for a course in a B.Sc. {engineering_domain} program. Consider the following aspects:

    *   **Teaching Strategies:** Suggest effective teaching strategies for this type of course.
    *   **Assessment Methods:** Recommend appropriate assessment methods to evaluate student learning.
    *   **Success Metrics:** What metrics can be used to measure the success of the course in achieving its objectives and CLOs?

    Provide the guidelines in the following format:

    **Teaching Strategies:** [Text]
    **Assessment Methods:** [Text]
    **Success Metrics:** [Text]
    """
    response = model.generate_content(prompt)
    guidelines = {}
    lines = response.text.split("\n")
    for line in lines:
        if line.startswith("**Teaching Strategies:**"):
            guidelines["Teaching Strategies"] = line.replace("**Teaching Strategies:**", "").strip()
        elif line.startswith("**Assessment Methods:**"):
            guidelines["Assessment Methods"] = line.replace("**Assessment Methods:**", "").strip()
        elif line.startswith("**Success Metrics:**"):
            guidelines["Success Metrics"] = line.replace("**Success Metrics:**", "").strip()

    return guidelines

def generate_continuous_improvement_framework(engineering_domain):
    """Generates a basic continuous improvement framework using the Gemini API."""
    prompt = f"""
    You are an expert in {engineering_domain} curriculum design and quality assurance.

    Develop a basic framework for continuous improvement for a course in a B.Sc. {engineering_domain} program.

    The framework should mention how the course will be regularly reviewed and updated, and how student feedback will be gathered and incorporated.

    Provide the framework as a concise paragraph (around 100-150 words).
    """
    response = model.generate_content(prompt)
    return response.text

def generate_knowledge_profile(engineering_domain, prioritized_gas, inputs):
    """Generates a description of the Knowledge Profile using the Gemini API."""
    gas_formatted = ", ".join(prioritized_gas)
    prompt = f"""
    You are an expert in {engineering_domain} curriculum design, familiar with the Washington Accord's definition of Knowledge Profile (WK1-WK9) for engineers.

    For a course in the B.Sc. {engineering_domain} program, provide a description of the Knowledge Profile that students are expected to develop.

    Consider the prioritized Graduate Attributes (GAs) for this course: {gas_formatted}

    Also, take into account any specific information provided in the user inputs:

    {inputs}

    Provide a description that outlines the key areas of knowledge (WK1-WK9) that are relevant to this course and explains how the course content and activities contribute to the development of this knowledge profile.

    Refer to the knowledge profiles using their codes (WK1-WK9). Limit the selection to only 2-3 most relevant knowledge profiles.
    """
    response = model.generate_content(prompt)
    return response.text.strip()

# --- Helper function for the CLO Summary Table ---
def extract_module_number(module_line):
    """
    Extracts the module number from a line of text.

    Assumes the format is "Module [Number]: [Title]".
    Returns the module number as an integer, or None if not found.
    """
    match = re.match(r"\*\*?Module\s*(\d+):", module_line)
    if match:
        return int(match.group(1))
    return None

def extract_modules_for_clo(course_contents, clo_label):
    """
    Extracts the modules that address a specific CLO.

    Args:
        course_contents (str): The course content outline as a string.
        clo_label (str): The CLO label (e.g., "CLO-1").

    Returns:
        str: A comma-separated string of module numbers that address the CLO.
    """
    modules = []
    module_lines = course_contents.split("**Module")[1:]
    for module_line in module_lines:
        module_number = extract_module_number(module_line)
        if module_number and clo_label in module_line:
            modules.append(f"Module {module_number}")
    return ", ".join(modules)

def generate_clo_summary_table(clos, course_name, credit_hours, course_contents):
    """Generates the CLO Summary Table in Markdown format"""
    return "CLO Summary Table placeholder"

# --- Output Generation ---

def generate_course_design_output(course_design):
    """Formats the course design into a Markdown string."""
    output = ""
    for section_name, section_content in course_design.items():
        output += f"## {section_name}\n"
        if isinstance(section_content, dict):
            for key, value in section_content.items():
                output += f"**{key}:** {value}\n"
        elif isinstance(section_content, list):
            if section_name == "CLOs":
                for clo in section_content:
                    output += f"- **{clo['CLO']}:** {clo['Statement']} (**{clo['GA']}**, {clo['Taxonomy']})\n"
            elif section_name == "Course Contents":
                for module in section_content:
                    output += f"- **{module['Module']}:** {module['Title']}\n"
            else:
                for item in section_content:
                    output += f"- {item}\n"
        elif isinstance(section_content, str):
            output += f"{section_content}\n"
        output += "\n"
    return output

def save_output_to_file(output_string, filename="course_design_output.md"):
    """Saves the output string to a file."""
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(output_string)
    except Exception as e:
        print(f"Error saving output to file: {e}")

def generate_excel_summary(course_design, filename="course_design_summary.xlsx"):
    """Generates an MS Excel summary of the course design with CLOs on separate rows."""
    try:
        # Create a new workbook and select the active worksheet
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Course Design Summary"

        # Define the header row
        headers = [
            "Sr. No.", "CLO", "CLO Statement", "Aligned GA",
            "Bloom's Taxonomy", "Course Contents/Modules", "KPIs"
        ]
        worksheet.append(headers)

        # Set the header row style
        for col_num, header in enumerate(headers, 1):
            cell = worksheet.cell(row=1, column=col_num, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(
                left=Side(style="thin"), right=Side(style="thin"),
                top=Side(style="thin"), bottom=Side(style="thin")
            )

        # Add the CLO data to the worksheet
        for i, clo in enumerate(course_design["CLOs"], start=1):
            row = [
                i,
                f"CLO-{i}",
                clo["Statement"],
                clo["GA"],
                clo["Taxonomy"],
                extract_modules_for_clo(course_design["Course Contents"], f"CLO-{i}"),
                ", ".join(clo["KPIs"])
            ]
            worksheet.append(row)

        # Adjust column widths
        for col_num, _ in enumerate(headers, 1):
            column_letter = get_column_letter(col_num)
            worksheet.column_dimensions[column_letter].width = 20

        # Save the workbook to a file
        workbook.save(filename)
        print(f"Course design summary saved to {filename}")
    except PermissionError as e:
        print(f"Error saving to Excel: {e}. Please ensure that the Excel file is closed before running the program")
    except Exception as e:
        print(f"Error saving to Excel: {e}. Please ensure you have openpyxl installed (`pip install openpyxl`). Also review CLO generation logic.")

# --- Main Program Logic ---

def main():
    # Example of using shell commands
    print(f"Current shell: {get_shell_type()}")
    
    # Example of executing a non-interactive command
    returncode, stdout, stderr = execute_shell_command("dir" if platform.system() == "Windows" else "ls")
    if returncode == 0:
        print("Command output:", stdout)
    else:
        print("Error:", stderr)

    # Execute a simple command
    returncode, stdout, stderr = execute_shell_command("dir")  # or "ls" on Unix
    if returncode == 0:
        print("Success:", stdout)
    else:
        print("Error:", stderr)

    # Run an interactive command
    run_interactive_command("python -m pip install -r requirements.txt")

    # Check shell type
    shell = get_shell_type()  # Returns 'cmd', 'powershell', or 'bash'

    curriculum_prompt = load_yaml_prompt("curriculum_prompt.yaml")
    course_prompt = load_yaml_prompt("course_prompt.yaml")

    while True:
        choice = input("Do you want to design a course (c) or evaluate a curriculum (e)? (Enter 'q' to quit): ")

        if choice.lower() == 'c':
            inputs = get_user_input_for_course(course_prompt)
            if inputs is not None:
                course_design = generate_course_design(inputs, course_prompt)
                if course_design is not None:
                    output = generate_course_design_output(course_design)
                    save_output_to_file(output)
                    print("Course design saved to course_design_output.md")
                    generate_excel_summary(course_design)
                    print("Course design summary saved to course_design_summary.xlsx")

        elif choice.lower() == "e":
            inputs = get_user_input_for_curriculum(curriculum_prompt)
            if inputs is not None:
                print("Curriculum evaluation functionality is not yet implemented.")

        elif choice.lower() == "q":
            break
        else:
            print("Invalid choice. Please enter 'c', 'e', or 'q'.")

if __name__ == "__main__":
    main()