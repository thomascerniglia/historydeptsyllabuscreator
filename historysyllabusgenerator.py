import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter import scrolledtext
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from tkinter import font as tkfont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
import json
import os

# Default template text for various sections (from official AMH2020 syllabus)
course_description_default = (
    "In this course, students will trace the history of the United States from the "
    "end of the Reconstruction era to the contemporary era. Topics will include but are not limited to the rise "
    "of Industrialization, the United States' emergence as an actor on the world stage, Constitutional amendments "
    "and their impact, the Progressive era, World War I, the Great Depression and New Deal, World War II, the "
    "Civil Rights era, the Cold War, and the United States since 1989.\n\n"
    "NOTE: All topics in this course will be taught objectively as objects of analysis, without endorsement of "
    "particular viewpoints, and will be observed from multiple perspectives. No lesson is intended to espouse, promote, advance, "
    "inculcate, or compel a particular feeling, perception, or belief. Students are encouraged to employ critical thinking and to rely on data and verifiable "
    "sources to explore readings and subject matter in this course. All perspectives will be respected in class discussions."
)
prerequisites_default = "None"
gen_ed_default = (
    "General Education Designation: Social and Behavioral Sciences (S)\n"
    "Social Science courses must afford students an understanding of the basic social and behavioral science concepts "
    "and principles used in the analysis of behavior and past and present social, political, and economic issues. Social "
    "and Behavioral Sciences (S) is a sub-designation of Social Sciences at the University of Florida. These courses "
    "provide instruction in the history, key themes, principles, terminology, and underlying theory or methodologies used in "
    "the social and behavioral sciences. Students will learn to identify, describe and explain social institutions, structures or processes. "
    "These courses emphasize the effective application of accepted problem-solving techniques. Students will apply formal and informal qualitative or "
    "quantitative analysis to examine the processes and means by which individuals make personal and group decisions, as well as the evaluation of opinions, "
    "outcomes or human behavior. Students are expected to assess and analyze ethical perspectives in individual and societal decisions.\n\n"
    "Your successful completion of AMH2020 with a grade of \"C\" or higher will count towards UF's General Education State Core in Social Science (S). "
    "It will also count towards the State of Florida's Civic Literacy requirement."
)
course_objectives_default = (
    "All General Education area objectives can be found **here** (UF General Education website).\n\n"
    "The AMH2020 curriculum will also cover the following course-specific objectives:\n"
    "- Address how the Civil War and Reconstruction set the stage for the development of the modern United States.\n"
    "- Explore how U.S. involvement in the Spanish-American War, World War I, and World War II reshaped U.S. foreign policy and civil society.\n"
    "- Present the origins of the Cold War, its implications for U.S. international relations, and its influence on American political culture.\n"
    "- Enable students to analyze and evaluate the origins and influences of the civil rights movement, the Vietnam War, the women's movement, and New Right conservatism.\n"
    "- Teach students how to analyze historical documents and scholarship from a range of authors and time periods."
)

student_learning_outcomes_default = (
    "A student who successfully completes this course will be able to:\n"
    "- Describe the factual details of the substantive historical episodes under study.\n"
    "- Identify and analyze foundational developments that shaped American history since 1877 using critical thinking skills.\n"
    "- Demonstrate an understanding of the primary ideas, values, and perceptions that have shaped American history.\n"
    "- Demonstrate competency in civic literacy."
)

# Default text for university policies and resources sections:
honesty_plagiarism_default = (
    "UF students are bound by The Honor Pledge which states: \"We, the members of the University of Florida community, "
    "pledge to hold ourselves and our peers to the highest standards of honor and integrity by abiding by the Honor Code. On all work submitted for credit "
    "by students at the University of Florida, the following pledge is either required or implied: 'On my honor, I have neither given nor received unauthorized aid "
    "in doing this assignment.' The Conduct Code specifies a number of behaviors that are in violation of this code and the possible sanctions. See the UF Conduct Code website for more information. "
    "If you have any questions or concerns, please consult with the instructor or TAs in this class.\n\n"
    "Ethical violations such as plagiarism, cheating, and other academic misconduct will not be tolerated and will result in a failing grade in this course. "
    "Students must be especially wary of plagiarism. The UF Student Honor Code defines plagiarism as: *A student shall not represent as the student's own work all or any portion of the work of another.* "
    "Plagiarism includes (but is not limited to): **a.** Quoting oral or written materials, whether published or unpublished, without proper attribution; **b.** Submitting a document or assignment which in whole or in part is identical or substantially identical to a document or assignment not authored by the student. "
    "We will discuss these issues in greater detail prior to the first written assignment. *Note:* plagiarism also includes the use of any artificial intelligence program (e.g., ChatGPT) to produce work for this course."
)
recording_policy_default = (
    "Students are allowed to record video or audio of class lectures. However, the purposes for which these recordings may be used are strictly controlled. "
    "The only allowable purposes are **(1)** for personal educational use, **(2)** in connection with a complaint to the university, or **(3)** as evidence in, or in preparation for, a criminal or civil proceeding. "
    "All other purposes are prohibited. Specifically, students may **not publish** recorded lectures without the written consent of the instructor.\n\n"
    'A "class lecture" is an educational presentation intended to inform or teach enrolled students about a particular subject, including any instructor-led discussions, delivered by the instructor or an invited guest lecturer. '
    "A class lecture **does not** include lab sessions, student presentations, clinical presentations, assessments (quizzes, tests, exams), field trips, or private conversations during a class session.\n\n"
    'To "publish" a recording means to share or transmit it to another person, or to upload it to any media platform (including social media or note-sharing services). A student who publishes a recording without permission may be subject to a civil lawsuit and/or disciplinary action by the university.'
)
accommodations_default = (
    "Students with disabilities who experience learning barriers and would like to request academic accommodations should connect with the Disability Resource Center by visiting **https://disability.ufl.edu/students/get-started/**. "
    "It is important for students to share their accommodation letter with the instructor and discuss their access needs as early as possible in the semester."
)
canvas_policy_default = (
    "Class announcements will be made through **Canvas**, and all assignments must be submitted via Canvas. Course materials (handouts, lecture slides, assignment rubrics, readings, study guides, writing samples, and this syllabus) will be available on Canvas. "
    "Please check your Canvas inbox regularly and read all course announcements."
)
technology_policy_default = (
    "To accommodate various learning styles, laptops and tablets are permitted in lecture for note-taking or class-related activities **as long as they do not become a distraction**. "
    "Abuse of this policy (e.g., unrelated web browsing or messaging during class) may result in loss of technology privileges or being marked absent for that session. "
    "No computers or electronic devices are allowed during exams."
)
evaluations_default = (
    "Students are expected to provide professional and respectful feedback on the quality of instruction by completing **course evaluations** online via GatorEvals. "
    "Evaluations can be done via the email link from GatorEvals, the link in Canvas, or by logging in to the GatorEvals portal. Students will be notified when the evaluation period opens, and can view summary results of past evaluations on the GatorEvals website."
)
conflict_resolution_default = (
    "Any classroom issues, disagreements, or grade disputes should first be discussed between the student and instructor. "
    "If the problem cannot be resolved, please contact the Associate Chair of the History Department. Be prepared to provide documentation of the issue and all graded materials. "
    "Unresolved issues may be referred to the University Ombuds Office or the Dean of Students Office for further resolution."
)
campus_resources_default = (
    "**Campus Resources:**\n"
    "- *U Matter, We Care*: If you or someone you know is in distress, please email umatter@ufl.edu or call 352-392-1575.\n"
    "- *Counseling and Wellness Center*: Visit counseling.ufl.edu or call 352-392-1575 for information on crisis services and counseling.\n"
    "- *Student Health Care Center*: Call 352-392-1161 for 24/7 assistance or visit the SHCC website.\n"
    "- *University Police Department*: 352-392-1111 (or 911 for emergencies).\n"
    "- *UF Health Shands ER*: For immediate medical care, call 352-733-0111 or go to the ER at 1515 SW Archer Road.\n"
    "- *GatorWell Health Promotion Services*: For wellness coaching and services, call 352-273-4450 or visit gatorwell.ufsa.ufl.edu.\n"
    "- *Field and Fork Food Pantry*: For students experiencing food insecurity.\n\n"
    "**Academic Resources:**\n"
    "- *E-learning Technical Support*: 352-392-4357 (helpdesk@ufl.edu) – assistance with Canvas and tech issues.\n"
    "- *Career Connections Center*: 352-392-1601 – career planning and placement services.\n"
    "- *Library Support*: assistance with library resources – library.ufl.edu/help.\n"
    "- *Teaching Center*: 352-392-2010 – tutoring and study skills service (Broward Hall).\n"
    "- *Writing Studio*: 352-846-1138 – help with brainstorming and writing (2215 Turlington Hall).\n"
    "- *Student Complaints*: Visit the Student Honor Code and Student Conduct Code webpage for information on filing complaints; Online students can use the Distance Learning Complaint Process."
)

class SyllabusTemplate:
    def __init__(self, course_code, title, description="", objectives=None, outcomes=None):
        self.course_code = course_code
        self.title = title
        self.description = description
        self.objectives = objectives or []
        self.outcomes = outcomes or []
        self.fixed_policies = True
        self.required_sections = {
            "course_info": True,
            "instructor_info": True,
            "course_description": True,
            "objectives": True,
            "outcomes": True,
            "grading": True,
            "schedule": True,
            "policies": True
        }

class HistorySyllabusGenerator:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("History Department Syllabus Generator")
        self.root.geometry("1200x800")
        
        # Load available templates
        self.templates = self.load_default_templates()
        
        self.setup_styles()
        self.create_main_interface()
        
    def load_default_templates(self):
        """Load the default course templates"""
        templates = {}
        
        # Test Template with full sample data
        test_template = SyllabusTemplate(
            "TEST101",
            "Introduction to Testing and Quality Assurance",
            description=(
                "This comprehensive course introduces students to the fundamental principles and practices of "
                "software testing and quality assurance. Students will learn various testing methodologies, "
                "test design techniques, and quality control processes essential in modern software development. "
                "The course covers both theoretical foundations and practical applications, including unit testing, "
                "integration testing, system testing, and acceptance testing. Special emphasis is placed on "
                "automated testing tools, continuous integration, and quality metrics.\n\n"
                "Through hands-on projects and real-world case studies, students will develop critical thinking "
                "skills necessary for identifying and resolving software defects, improving code quality, and "
                "ensuring robust software delivery. The course also explores the role of quality assurance in "
                "the software development lifecycle and its impact on project success."
            ),
            objectives=[
                "Test Objective 1: Verify proper formatting of course objectives",
                "Test Objective 2: Ensure proper export of lists and bullet points",
                "Test Objective 3: Check alignment and spacing in the final document",
                "Test Objective 4: Validate proper handling of special characters & symbols"
            ],
            outcomes=[
                "Test Outcome 1: Successfully generate properly formatted syllabi",
                "Test Outcome 2: Correctly display all template elements",
                "Test Outcome 3: Maintain consistent styling throughout the document",
                "Test Outcome 4: Handle various content types appropriately"
            ]
        )
        test_template.sample_data = {
            # Course Information
            "course_info": {
                "term": "Fall 2024",
                "credits": "3",
                "meeting_times": "MWF 10:40 AM - 11:30 AM",
                "location": "Keene-Flint Hall 050",
                "materials_fee": "25",
                "required_materials": "1. Test Textbook (ISBN: 978-0123456789)\n2. Course Pack from Target Copy\n3. Laptop or tablet for in-class activities"
            },
            # Instructor Information
            "instructor_info": {
                "name": "Dr. Test Professor",
                "office": "Keene-Flint Hall 234",
                "phone": "(352) 123-4567",
                "email": "test.professor@ufl.edu",
                "office_hours": "Monday 2:00-4:00 PM\nWednesday 1:00-3:00 PM\nOr by appointment"
            },
            # Teaching Assistants
            "tas": [
                {
                    "name": "Alice Anderson",
                    "email": "a.anderson@ufl.edu",
                    "office_hours": "Tuesday 10:00-11:30 AM"
                },
                {
                    "name": "Bob Brown",
                    "email": "b.brown@ufl.edu",
                    "office_hours": "Thursday 2:00-3:30 PM"
                },
                {
                    "name": "Charlie Chen",
                    "email": "c.chen@ufl.edu",
                    "office_hours": "Friday 1:00-2:30 PM"
                }
            ],
            # Grading Categories
            "grading_categories": [
                {
                    "name": "Participation",
                    "weight": "10",
                    "description": "Active participation in class discussions and activities",
                    "assignments": [
                        {"title": "Discussion Posts", "due": "Weekly", "points": "5"},
                        {"title": "In-class Activities", "due": "Various", "points": "5"}
                    ]
                },
                {
                    "name": "Essays",
                    "weight": "30",
                    "description": "Three analytical essays on course topics",
                    "assignments": [
                        {"title": "Essay 1", "due": "September 15", "points": "100"},
                        {"title": "Essay 2", "due": "October 20", "points": "100"},
                        {"title": "Essay 3", "due": "November 25", "points": "100"}
                    ]
                },
                {
                    "name": "Exams",
                    "weight": "40",
                    "description": "Midterm and final examinations",
                    "assignments": [
                        {"title": "Midterm Exam", "due": "October 10", "points": "100"},
                        {"title": "Final Exam", "due": "December 15", "points": "100"}
                    ]
                },
                {
                    "name": "Project",
                    "weight": "20",
                    "description": "Group research project and presentation",
                    "assignments": [
                        {"title": "Project Proposal", "due": "September 30", "points": "20"},
                        {"title": "Progress Report", "due": "November 1", "points": "30"},
                        {"title": "Final Presentation", "due": "December 1", "points": "50"}
                    ]
                }
            ],
            # Course Schedule
            "schedule": [
                {
                    "date": "Aug 23",
                    "topic": "Course Introduction",
                    "readings": "Syllabus Review",
                    "work_due": ""
                },
                {
                    "date": "Aug 30",
                    "topic": "Topic 1: Sample Lecture",
                    "readings": "[P] Primary Source Reading 1 (pp. 1-15)\nSecondary Source Article (pp. 16-30)",
                    "work_due": "Discussion Post 1"
                },
                {
                    "date": "Sep 6",
                    "topic": "Topic 2: Example Discussion",
                    "readings": "Textbook Chapter 1 (pp. 31-45)\n[P] Document Analysis Exercise",
                    "work_due": "Essay 1"
                },
                {
                    "date": "Sep 13",
                    "topic": "Topic 3: Test Lecture",
                    "readings": "[P] Primary Source Collection (pp. 46-60)\nScholarly Article Review",
                    "work_due": "Group Activity"
                },
                {
                    "date": "Sep 20",
                    "topic": "Topic 4: Sample Seminar",
                    "readings": "Textbook Chapter 2 (pp. 61-75)\n[P] Historical Documents",
                    "work_due": "Project Proposal"
                }
            ],
            # Policies
            "policies": {
                "late_work": True,
                "attendance": True,
                "technology": True,
                "extra_credit": True,
                "late_policy": "Standard (10% per day)",
                "extra_credit_policy": "Optional assignments"
            }
        }
        templates["TEST101"] = test_template
        
        return templates
    
    def setup_styles(self):
        """Configure the visual styles for the application"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure modern styles
        style.configure("TButton", padding=6, relief="flat", background="#2196F3")
        style.configure("TEntry", padding=5)
        style.configure("TLabel", padding=5)
        style.configure("Heading.TLabel", font=('Helvetica', 12, 'bold'))
        style.configure("Template.TFrame", padding=10, relief="raised")
        style.configure("Italic.TLabel", font=('Helvetica', 10, 'italic'))
        
        # New styles for schedule tab
        style.configure("Schedule.TButton", padding=5, relief="flat", background="#2196F3")
        style.configure("Action.TButton", 
                       padding=(10, 5), 
                       relief="flat", 
                       background="#2196F3",
                       font=('Helvetica', 9))
        style.configure("Delete.TButton", 
                       padding=2, 
                       relief="flat", 
                       background="#ff5252",
                       width=2)
        style.configure("Small.TButton", 
                       padding=2, 
                       relief="flat", 
                       background="#2196F3",
                       width=3)
        
    def create_main_interface(self):
        """Create the main interface with template selection"""
        # Create main container with scrolling
        self.main_container = ttk.Frame(self.root)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Template selection at the top
        template_frame = ttk.Frame(self.main_container, style="Template.TFrame")
        template_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(template_frame, text="Select Course Template:", style="Heading.TLabel").pack(side=tk.LEFT, padx=5)
        self.template_var = tk.StringVar()
        template_combo = ttk.Combobox(template_frame, textvariable=self.template_var, 
                                    values=list(self.templates.keys()),
                                    width=40)  # Made wider
        template_combo.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        template_combo.bind('<<ComboboxSelected>>', self.on_template_selected)
        
        # Create notebook for different sections
        self.notebook = ttk.Notebook(self.main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create tabs
        self.create_course_info_tab()
        self.create_instructor_info_tab()
        self.create_schedule_tab()
        self.create_assignments_tab()
        self.create_policies_tab()
        
        # Bottom frame for actions
        self.create_action_buttons()
        
    def create_scrollable_frame(self, parent):
        """Create a scrollable frame within the given parent"""
        canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        return scrollable_frame
    
    def create_course_info_tab(self):
        """Create the course information tab"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Course Information")
        
        frame = self.create_scrollable_frame(tab)
        
        # Course info fields
        fields = [
            ("Course Number:", "course_num", 15),
            ("Course Title:", "course_title", 40),
            ("Term:", "term", 20),
            ("Credits:", "credits", 5),
            ("Meeting Days/Times:", "meeting_times", 40),
            ("Location:", "location", 30)
        ]
        
        for i, (label, field_name, width) in enumerate(fields):
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky="e", padx=5, pady=5)
            entry = ttk.Entry(frame, width=width)
            entry.grid(row=i, column=1, sticky="w", padx=5, pady=5)
            setattr(self, f"entry_{field_name}", entry)
        
        # Course description
        ttk.Label(frame, text="Course Description:").grid(row=len(fields), column=0, sticky="ne", padx=5, pady=5)
        self.txt_description = scrolledtext.ScrolledText(frame, width=60, height=6, wrap=tk.WORD)
        self.txt_description.grid(row=len(fields), column=1, sticky="w", padx=5, pady=5)
        
    def create_instructor_info_tab(self):
        """Create the instructor information tab with proper layout"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Instructor Information")
        
        # Create a canvas with scrollbar
        canvas = tk.Canvas(tab)
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        
        # Create main content frame
        content_frame = ttk.Frame(canvas)
        content_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        # Configure canvas
        canvas.create_window((0, 0), window=content_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack canvas and scrollbar
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Instructor Information
        instructor_frame = ttk.LabelFrame(content_frame, text="Instructor Information")
        instructor_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Grid for instructor info with consistent field names
        fields = [
            ("Name:", "name", 40),
            ("Office:", "office", 40),
            ("Phone:", "phone", 20),
            ("Email:", "email", 40),
            ("Office Hours:", "office_hours", 60)
        ]
        
        for i, (label, field, width) in enumerate(fields):
            ttk.Label(instructor_frame, text=label).grid(row=i, column=0, sticky="e", padx=5, pady=5)
            entry = ttk.Entry(instructor_frame, width=width)
            entry.grid(row=i, column=1, sticky="w", padx=5, pady=5)
            setattr(self, f"entry_instr_{field}", entry)
        
        # Teaching Assistants
        self.ta_frame = ttk.LabelFrame(content_frame, text="Teaching Assistants")
        self.ta_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.ta_container = ttk.Frame(self.ta_frame)
        self.ta_container.pack(fill=tk.X, padx=5, pady=5)
        
        self.ta_entries = []
        
        ttk.Button(self.ta_frame, text="Add Teaching Assistant", 
                  command=self.add_ta,
                  style="Action.TButton").pack(pady=5)
        
    def create_schedule_tab(self):
        """Create the course schedule tab with proper layout"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Course Schedule")
        
        # Main container
        main_frame = ttk.Frame(tab, padding="5")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Top controls frame with modern styling
        controls_frame = ttk.Frame(main_frame)
        controls_frame.pack(fill=tk.X, pady=(0, 5))
        
        # Left side buttons with enhanced styling
        buttons_frame = ttk.Frame(controls_frame)
        buttons_frame.pack(side=tk.LEFT)
        
        ttk.Button(buttons_frame, text="Add Entry", 
                  command=self.add_schedule_entry,
                  style="Action.TButton").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(buttons_frame, text="Import Schedule",
                  command=self.import_schedule,
                  style="Action.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Export Schedule",
                  command=self.export_schedule,
                  style="Action.TButton").pack(side=tk.LEFT, padx=5)
        
        # Right side helper text
        ttk.Label(controls_frame, text="Use [P] to mark primary sources", 
                  style="Italic.TLabel").pack(side=tk.RIGHT)
        
        # Create a frame for the schedule content with proper styling
        self.schedule_frame = ttk.Frame(main_frame)
        self.schedule_frame.pack(fill=tk.BOTH, expand=True)
        self.schedule_frame.grid_columnconfigure(1, weight=1)
        self.schedule_frame.grid_columnconfigure(2, weight=2)
        
        # Headers with enhanced styling
        ttk.Label(self.schedule_frame, text="Date", 
                 style="Heading.TLabel").grid(row=0, column=0, sticky="w", padx=(5, 10))
        ttk.Label(self.schedule_frame, text="Topic",
                 style="Heading.TLabel").grid(row=0, column=1, sticky="w", padx=5)
        ttk.Label(self.schedule_frame, text="Readings/Preparation",
                 style="Heading.TLabel").grid(row=0, column=2, sticky="w", padx=5)
        ttk.Label(self.schedule_frame, text="Work Due",
                 style="Heading.TLabel").grid(row=0, column=3, sticky="w", padx=5)
        
        # Create scrollable frame with styled background
        canvas = tk.Canvas(self.schedule_frame, bg='#f5f5f5')  # Light gray background
        scrollbar = ttk.Scrollbar(self.schedule_frame, orient="vertical", command=canvas.yview)
        
        self.entries_frame = ttk.Frame(canvas, style="Template.TFrame")
        self.entries_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        # Configure the canvas
        canvas.create_window((0, 0), window=self.entries_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Grid the canvas and scrollbar
        canvas.grid(row=1, column=0, columnspan=4, sticky="nsew", pady=5)
        scrollbar.grid(row=1, column=4, sticky="ns", pady=5)
        
        # Configure grid weights
        self.schedule_frame.grid_rowconfigure(1, weight=1)
        self.entries_frame.grid_columnconfigure(1, weight=1)
        self.entries_frame.grid_columnconfigure(2, weight=2)
        
        # Enable mousewheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        self.schedule_entries = []
        
    def create_assignments_tab(self):
        """Create the assignments and grading tab with proper layout"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Assignments & Grading")
        
        # Create a canvas with scrollbar for the main content
        canvas = tk.Canvas(tab)
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        
        # Create main content frame
        content_frame = ttk.Frame(canvas)
        content_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        # Configure canvas
        canvas.create_window((0, 0), window=content_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack canvas and scrollbar
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Required Materials Section
        materials_frame = ttk.LabelFrame(content_frame, text="Required Materials")
        materials_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.materials_text = scrolledtext.ScrolledText(materials_frame, width=60, height=4, wrap=tk.WORD)
        self.materials_text.pack(fill=tk.X, padx=5, pady=5)
        
        fee_frame = ttk.Frame(materials_frame)
        fee_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(fee_frame, text="Materials Fee: $").pack(side=tk.LEFT)
        self.fee_entry = ttk.Entry(fee_frame, width=10)
        self.fee_entry.pack(side=tk.LEFT)
        
        # Grading Components Section
        components_frame = ttk.LabelFrame(content_frame, text="Graded Components")
        components_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.categories_frame = ttk.Frame(components_frame)
        self.categories_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.category_frames = []
        
        ttk.Button(components_frame, text="Add Category", 
                  command=self.add_category,
                  style="Action.TButton").pack(padx=5, pady=5)
        
        # Course Policies Section
        policies_frame = ttk.LabelFrame(content_frame, text="Course Policies")
        policies_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Late Submissions Policy
        late_frame = ttk.Frame(policies_frame)
        late_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(late_frame, text="Late Submissions Policy:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.late_policies = {
            "Standard (10% per day)": "Unless an extension is granted, assignments will incur a 10-point penalty for every day they are late.",
            "No late work": "No late work will be accepted without prior approval.",
            "48-hour grace": "Students have a 48-hour grace period for submissions, after which no late work will be accepted.",
            "Custom": ""
        }
        
        self.late_policy_var = tk.StringVar()
        self.late_combo = ttk.Combobox(late_frame, textvariable=self.late_policy_var, 
                                     values=list(self.late_policies.keys()), width=30)
        self.late_combo.pack(side=tk.LEFT)
        self.late_combo.set("Standard (10% per day)")
        
        self.late_policy_text = scrolledtext.ScrolledText(policies_frame, width=60, height=3, wrap=tk.WORD)
        self.late_policy_text.pack(fill=tk.X, padx=5, pady=(0, 5))
        
        def update_late_policy(*args):
            selected = self.late_policy_var.get()
            if selected in self.late_policies:
                self.late_policy_text.config(state='normal')
                self.late_policy_text.delete('1.0', tk.END)
                self.late_policy_text.insert('1.0', self.late_policies[selected])
                if selected != "Custom":
                    self.late_policy_text.config(state='disabled')
            
                # Update the checkbox in policies tab
                if hasattr(self, 'optional_policies'):
                    if selected == "No late work":
                        self.optional_policies["late_work"].set(False)
                    else:
                        self.optional_policies["late_work"].set(True)
        
        self.late_policy_var.trace_add('write', update_late_policy)
        
        # Extra Credit Policy
        extra_frame = ttk.Frame(policies_frame)
        extra_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(extra_frame, text="Extra Credit Policy:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.extra_credit_policies = {
            "Standard": "Extra credit opportunities may be announced during the semester. Points will be added to your mid-term exam grade.",
            "No extra credit": "No extra credit will be offered in this course.",
            "Optional assignments": "Students may complete optional assignments for extra credit, worth up to 3% of the final grade.",
            "Custom": ""
        }
        
        self.extra_credit_var = tk.StringVar()
        self.extra_combo = ttk.Combobox(extra_frame, textvariable=self.extra_credit_var,
                                      values=list(self.extra_credit_policies.keys()), width=30)
        self.extra_combo.pack(side=tk.LEFT)
        self.extra_combo.set("Select a policy...")
        
        self.extra_credit_text = scrolledtext.ScrolledText(policies_frame, width=60, height=3, wrap=tk.WORD)
        self.extra_credit_text.pack(fill=tk.X, padx=5, pady=(0, 5))
        
        def update_extra_credit(*args):
            selected = self.extra_credit_var.get()
            if selected in self.extra_credit_policies:
                self.extra_credit_text.config(state='normal')
                self.extra_credit_text.delete('1.0', tk.END)
                self.extra_credit_text.insert('1.0', self.extra_credit_policies[selected])
                if selected != "Custom":
                    self.extra_credit_text.config(state='disabled')
        
        self.extra_credit_var.trace_add('write', update_extra_credit)
        
        # Bind combobox events
        self.late_combo.bind('<<ComboboxSelected>>', lambda e: update_late_policy())
        self.extra_combo.bind('<<ComboboxSelected>>', lambda e: update_extra_credit())
        
        # Grading Scale Section
        scale_frame = ttk.LabelFrame(content_frame, text="Grading Scale")
        scale_frame.pack(fill=tk.X, padx=5, pady=5)
        
        grades_frame = ttk.Frame(scale_frame)
        grades_frame.pack(padx=5, pady=5)
        
        grades = [
            ("A", "93-100"), ("A-", "90-92"),
            ("B+", "87-89"), ("B", "83-86"),
            ("B-", "80-82"), ("C+", "77-79"),
            ("C", "73-76"), ("C-", "70-72"),
            ("D+", "67-69"), ("D", "63-66"),
            ("D-", "60-62"), ("E", "0-59")
        ]
        
        for i, (letter, number) in enumerate(grades):
            row = i // 4
            col = i % 4
            grade_frame = ttk.Frame(grades_frame)
            grade_frame.grid(row=row, column=col, padx=5, pady=2)
            
            ttk.Label(grade_frame, text=f"{letter}:").pack(side=tk.LEFT)
            entry = ttk.Entry(grade_frame, width=8)
            entry.insert(0, number)
            entry.pack(side=tk.LEFT, padx=2)
        
    def create_policies_tab(self):
        """Create the policies tab"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Policies")
        
        frame = self.create_scrollable_frame(tab)
        
        # Fixed policies (non-editable)
        ttk.Label(frame, text="Department Policies (Fixed):", style="Heading.TLabel").pack(anchor="w", padx=5, pady=(5,10))
        
        fixed_policies = [
            "Academic Honesty",
            "Recording Policy",
            "Accommodations",
            "Canvas Use Policy",
            "Course Evaluations"
        ]
        
        # Create variables for fixed policies
        self.fixed_policy_vars = {}
        for policy in fixed_policies:
            var = tk.BooleanVar(value=True)
            self.fixed_policy_vars[policy] = var
            ttk.Checkbutton(frame, text=policy, state="disabled", variable=var).pack(anchor="w", padx=20, pady=2)
        
        # Optional policies
        ttk.Label(frame, text="Optional Policies:", style="Heading.TLabel").pack(anchor="w", padx=5, pady=(20,10))
        
        self.optional_policies = {
            "attendance": tk.BooleanVar(value=True),
            "late_work": tk.BooleanVar(value=True),
            "technology": tk.BooleanVar(value=True),
            "extra_credit": tk.BooleanVar(value=False)
        }
        
        # Create checkbuttons and bind them to policy updates
        for policy, var in self.optional_policies.items():
            cb = ttk.Checkbutton(frame, text=policy.replace("_", " ").title(), variable=var)
            cb.pack(anchor="w", padx=20, pady=2)
            
            # Special handling for late work and extra credit policies
            if policy == "late_work":
                def on_late_work_toggle(*args):
                    if not var.get():  # If late work is disabled
                        if hasattr(self, 'late_combo'):
                            self.late_combo.set("No late work")
                            self.late_policy_text.config(state='normal')
                            self.late_policy_text.delete('1.0', tk.END)
                            self.late_policy_text.insert('1.0', self.late_policies["No late work"])
                            self.late_policy_text.config(state='disabled')
                    else:  # If late work is enabled
                        if hasattr(self, 'late_combo'):
                            self.late_combo.set("Standard (10% per day)")
                            self.late_policy_text.config(state='normal')
                            self.late_policy_text.delete('1.0', tk.END)
                            self.late_policy_text.insert('1.0', self.late_policies["Standard (10% per day)"])
                            self.late_policy_text.config(state='disabled')
            
                var.trace_add('write', on_late_work_toggle)
            
            elif policy == "extra_credit":
                def on_extra_credit_toggle(*args):
                    if not var.get():  # If extra credit is disabled
                        if hasattr(self, 'extra_combo'):
                            self.extra_combo.set("No extra credit")
                            self.extra_credit_text.config(state='normal')
                            self.extra_credit_text.delete('1.0', tk.END)
                            self.extra_credit_text.insert('1.0', self.extra_credit_policies["No extra credit"])
                            self.extra_credit_text.config(state='disabled')
                    else:  # If extra credit is enabled
                        if hasattr(self, 'extra_combo'):
                            self.extra_combo.set("Standard")
                            self.extra_credit_text.config(state='normal')
                            self.extra_credit_text.delete('1.0', tk.END)
                            self.extra_credit_text.insert('1.0', self.extra_credit_policies["Standard"])
                            self.extra_credit_text.config(state='disabled')
            
                var.trace_add('write', on_extra_credit_toggle)
        
    def create_action_buttons(self):
        """Create the bottom action buttons"""
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(button_frame, text="Save Template", command=self.save_template).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Load Template", command=self.load_template).pack(side=tk.LEFT, padx=5)
        
        # Export options
        export_frame = ttk.Frame(button_frame)
        export_frame.pack(side=tk.RIGHT)
        
        self.var_docx = tk.BooleanVar(value=True)
        self.var_pdf = tk.BooleanVar(value=False)
        
        ttk.Checkbutton(export_frame, text="Word (.docx)", variable=self.var_docx).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(export_frame, text="PDF", variable=self.var_pdf).pack(side=tk.LEFT, padx=5)
        ttk.Button(export_frame, text="Generate Syllabus", command=self.generate_syllabus).pack(side=tk.LEFT, padx=5)
        
    def add_ta(self):
        """Add a new TA entry"""
        frame = ttk.Frame(self.ta_container)
        frame.pack(anchor="w", pady=2)
        
        entries = []
        for label, width in [("Name:", 25), ("Email:", 30), ("Office Hours:", 45)]:
            ttk.Label(frame, text=label).pack(side=tk.LEFT)
            entry = ttk.Entry(frame, width=width)
            entry.pack(side=tk.LEFT, padx=5)
            entries.append(entry)
            
        self.ta_entries.append(entries)
        
        def remove_ta():
            frame.destroy()
            self.ta_entries.remove(entries)
            
        remove_btn = ttk.Button(frame, text="×", 
                              command=remove_ta,
                              style="Delete.TButton")
        remove_btn.pack(side=tk.LEFT, padx=2)
        
    def add_schedule_entry(self):
        """Add a new schedule entry with enhanced styling"""
        row = len(self.schedule_entries)
        
        # Date entry with modern styling
        date_entry = ttk.Entry(self.entries_frame, width=15)
        date_entry.grid(row=row, column=0, sticky="w", padx=(5, 10), pady=2)
        
        # Topic entry
        topic_entry = ttk.Entry(self.entries_frame, width=40)
        topic_entry.grid(row=row, column=1, sticky="ew", padx=5, pady=2)
        
        # Readings frame
        readings_frame = ttk.Frame(self.entries_frame)
        readings_frame.grid(row=row, column=2, sticky="ew", padx=5, pady=2)
        readings_frame.grid_columnconfigure(0, weight=1)
        
        # Readings text area with styled background
        readings_text = scrolledtext.ScrolledText(readings_frame, width=50, height=3, 
                                               wrap=tk.WORD, bg='#ffffff')
        readings_text.grid(row=0, column=0, sticky="ew")
        
        # Buttons frame
        buttons_frame = ttk.Frame(readings_frame)
        buttons_frame.grid(row=0, column=1, sticky="ns", padx=(5, 0))
        
        def insert_p_marker():
            readings_text.insert(tk.INSERT, "[P] ")
        
        def count_words():
            text = readings_text.get("1.0", tk.END).strip()
            word_count = len(text.split())
            readings_text.insert(tk.END, f" [{word_count} words]")
        
        # Compact buttons with modern styling
        ttk.Button(buttons_frame, text="[P]",
                  command=insert_p_marker,
                  style="Small.TButton").pack(side=tk.TOP, pady=(0, 2))
        ttk.Button(buttons_frame, text="#",
                  command=count_words,
                  style="Small.TButton").pack(side=tk.TOP)
        
        # Work Due entry
        work_due_entry = ttk.Entry(self.entries_frame, width=20)
        work_due_entry.grid(row=row, column=3, sticky="w", padx=5, pady=2)
        
        def remove_entry():
            date_entry.destroy()
            topic_entry.destroy()
            readings_frame.destroy()
            work_due_entry.destroy()
            delete_btn.destroy()
            self.schedule_entries.remove(entry_dict)
            self.repack_schedule_entries()
        
        # Delete button with red styling
        delete_btn = ttk.Button(self.entries_frame, text="×",
                              command=remove_entry,
                              style="Delete.TButton")
        delete_btn.grid(row=row, column=4, padx=(0, 5), pady=2)
        
        entry_dict = {
            "date": date_entry,
            "topic": topic_entry,
            "readings": readings_text,
            "work_due": work_due_entry,
            "delete_btn": delete_btn,
            "row": row
        }
        
        self.schedule_entries.append(entry_dict)
        return entry_dict

    def repack_schedule_entries(self):
        """Repack all schedule entries after a deletion"""
        for i, entry in enumerate(self.schedule_entries):
            entry["row"] = i  # Update row number
            entry["date"].grid(row=i, column=0, sticky="w", padx=(5, 10), pady=2)
            entry["topic"].grid(row=i, column=1, sticky="ew", padx=5, pady=2)
            entry["readings"].master.grid(row=i, column=2, sticky="ew", padx=5, pady=2)
            entry["work_due"].grid(row=i, column=3, sticky="w", padx=5, pady=2)
            entry["delete_btn"].grid(row=i, column=4, padx=(0, 5), pady=2)
        
    def on_template_selected(self, event):
        """Handle template selection"""
        selected = self.template_var.get()
        if selected in self.templates:
            template = self.templates[selected]
            self.load_template_content(template)
            
    def load_template_content(self, template):
        """Load the content from the selected template"""
        # Clear existing entries first
        self.clear_all_entries()
        
        # Basic course info
        self.entry_course_num.insert(0, template.course_code)
        self.entry_course_title.insert(0, template.title)
        self.txt_description.insert("1.0", template.description)

        # If it's the test template, populate all fields with sample data
        if hasattr(template, 'sample_data'):
            data = template.sample_data
            
            # Course Information
            self.entry_term.insert(0, data["course_info"]["term"])
            self.entry_credits.insert(0, data["course_info"]["credits"])
            self.entry_meeting_times.insert(0, data["course_info"]["meeting_times"])
            self.entry_location.insert(0, data["course_info"]["location"])
            
            # Materials
            if hasattr(self, 'materials_text'):
                self.materials_text.insert('1.0', data["course_info"]["required_materials"])
            if hasattr(self, 'fee_entry'):
                self.fee_entry.insert(0, data["course_info"]["materials_fee"])
            
            # Instructor Information
            instructor = data["instructor_info"]
            self.entry_instr_name.insert(0, instructor["name"])
            self.entry_instr_office.insert(0, instructor["office"])
            self.entry_instr_phone.insert(0, instructor["phone"])
            self.entry_instr_email.insert(0, instructor["email"])
            self.entry_instr_office_hours.insert(0, instructor["office_hours"])
            
            # Teaching Assistants
            for ta in data["tas"]:
                self.add_ta()
                ta_entries = self.ta_entries[-1]
                ta_entries[0].insert(0, ta["name"])
                ta_entries[1].insert(0, ta["email"])
                ta_entries[2].insert(0, ta["office_hours"])
            
            # Grading Categories
            for category in data["grading_categories"]:
                cat = self.add_category()
                cat["name"].insert(0, category["name"])
                cat["weight"].insert(0, category["weight"])
                cat["description"].insert('1.0', category["description"])
                
                # Add assignments to this category
                for assignment in category["assignments"]:
                    self.add_assignment_to_category(cat)
                    latest_assignment = cat["assignments"][-1]
                    latest_assignment["title"].insert(0, assignment["title"])
                    latest_assignment["due date"].insert(0, assignment["due"])
                    latest_assignment["points"].insert(0, assignment["points"])
            
            # Schedule
            for entry in data["schedule"]:
                schedule_entry = self.add_schedule_entry()
                schedule_entry["date"].insert(0, entry["date"])
                schedule_entry["topic"].insert(0, entry["topic"])
                schedule_entry["readings"].insert('1.0', entry["readings"])
                schedule_entry["work_due"].insert(0, entry["work_due"])
            
            # Policies
            policies = data["policies"]
            for policy_name, value in policies.items():
                if policy_name in self.optional_policies:
                    self.optional_policies[policy_name].set(value)
            
            # Update policy dropdowns
            if hasattr(self, 'late_combo') and "late_policy" in policies:
                self.late_combo.set(policies["late_policy"])
                self.late_policy_var.set(policies["late_policy"])
            
            if hasattr(self, 'extra_combo') and "extra_credit_policy" in policies:
                self.extra_combo.set(policies["extra_credit_policy"])
                self.extra_credit_var.set(policies["extra_credit_policy"])
        else:
            # For non-test templates, populate with template-specific data
            self.entry_term.insert(0, "Fall 2024")
            self.entry_credits.insert(0, "3")
            self.entry_meeting_times.insert(0, "MWF 10:40 AM - 11:30 AM")
            self.entry_location.insert(0, "TBA")

    def clear_all_entries(self):
        """Clear all entries in the form"""
        # Clear course info
        for entry in [self.entry_course_num, self.entry_course_title, 
                     self.entry_term, self.entry_credits, 
                     self.entry_meeting_times, self.entry_location]:
            entry.delete(0, tk.END)
        
        self.txt_description.delete("1.0", tk.END)
        
        # Clear instructor info with correct field names
        for field in ["name", "office", "phone", "email", "office_hours"]:
            entry = getattr(self, f"entry_instr_{field}", None)
            if entry:
                entry.delete(0, tk.END)
        
        # Clear TAs
        if hasattr(self, 'ta_entries'):
            for ta_entries in self.ta_entries:
                for entry in ta_entries:
                    entry.master.destroy()
            self.ta_entries.clear()
            
        # Recreate TA container
        if hasattr(self, 'ta_container'):
            self.ta_container.destroy()
            self.ta_container = ttk.Frame(self.ta_frame)
            self.ta_container.pack(fill=tk.X, padx=5, pady=5)
        
        # Clear materials
        if hasattr(self, 'materials_text'):
            self.materials_text.delete('1.0', tk.END)
        if hasattr(self, 'fee_entry'):
            self.fee_entry.delete(0, tk.END)
        
        # Clear grading categories
        if hasattr(self, 'category_frames'):
            for category in self.category_frames:
                category["frame"].destroy()
            self.category_frames.clear()
        
        # Clear schedule entries
        if hasattr(self, 'schedule_entries'):
            for entry in self.schedule_entries:
                for widget in [entry["date"], entry["topic"], entry["readings"].master, entry["work_due"], entry["delete_btn"]]:
                    if widget.winfo_exists():
                        widget.destroy()
            self.schedule_entries.clear()

    def add_assignment_to_category(self, category):
        """Add a new assignment to a category"""
        assignment_frame = ttk.Frame(category["frame"])
        assignment_frame.pack(fill=tk.X, padx=5, pady=2)
        
        ttk.Label(assignment_frame, text="Title:").pack(side=tk.LEFT)
        title_entry = ttk.Entry(assignment_frame, width=30)
        title_entry.pack(side=tk.LEFT, padx=2)
        
        ttk.Label(assignment_frame, text="Due:").pack(side=tk.LEFT)
        due_entry = ttk.Entry(assignment_frame, width=15)
        due_entry.pack(side=tk.LEFT, padx=2)
        
        ttk.Label(assignment_frame, text="Points:").pack(side=tk.LEFT)
        points_entry = ttk.Entry(assignment_frame, width=5)
        points_entry.pack(side=tk.LEFT, padx=2)
        
        def remove_assignment():
            assignment_frame.destroy()
            category["assignments"].remove(assignment_dict)
        
        remove_btn = ttk.Button(assignment_frame, text="×", 
                              command=remove_assignment,
                              style="Delete.TButton")
        remove_btn.pack(side=tk.LEFT, padx=2)
        
        assignment_dict = {
            "frame": assignment_frame,
            "title": title_entry,
            "due date": due_entry,
            "points": points_entry
        }
        
        if "assignments" not in category:
            category["assignments"] = []
        category["assignments"].append(assignment_dict)
        
        return assignment_dict

    def save_template(self):
        """Save the current settings as a template"""
        name = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            title="Save Template As"
        )
        if name:
            template_data = self.gather_template_data()
            with open(name, 'w') as f:
                json.dump(template_data, f, indent=2)
                
    def load_template(self):
        """Load a saved template"""
        name = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json")],
            title="Load Template"
        )
        if name:
            with open(name, 'r') as f:
                template_data = json.load(f)
            self.apply_template_data(template_data)
            
    def import_schedule(self):
        """Import schedule from CSV with enhanced formatting support"""
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")],
            title="Import Schedule"
        )
        if not file_path:
            return
        
        # Clear existing entries
        for entry in self.schedule_entries:
            entry["frame"].destroy()
        self.schedule_entries.clear()
        
        # Import based on file type
        if file_path.endswith('.csv'):
            import csv
            with open(file_path, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    entry = self.add_schedule_entry()
                    entry["date"].insert(0, row.get("Date", ""))
                    entry["topic"].insert(0, row.get("Topic", ""))
                    entry["readings"].insert("1.0", row.get("Readings/Preparation", ""))
                    entry["work_due"].insert(0, row.get("Work Due", ""))
        else:
            # Handle Excel import
            import pandas as pd
            df = pd.read_excel(file_path)
            for _, row in df.iterrows():
                entry = self.add_schedule_entry()
                entry["date"].insert(0, str(row.get("Date", "")))
                entry["topic"].insert(0, str(row.get("Topic", "")))
                entry["readings"].insert("1.0", str(row.get("Readings/Preparation", "")))
                entry["work_due"].insert(0, str(row.get("Work Due", "")))
            
    def export_schedule(self):
        """Export schedule to CSV with formatting preserved"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")],
            title="Export Schedule"
        )
        if not file_path:
            return
        
        # Prepare data
        data = []
        for entry in self.schedule_entries:
            data.append({
                "Date": entry["date"].get(),
                "Topic": entry["topic"].get(),
                "Readings/Preparation": entry["readings"].get("1.0", tk.END).strip(),
                "Work Due": entry["work_due"].get()
            })
        
        # Export based on file type
        if file_path.endswith('.csv'):
            import csv
            with open(file_path, 'w', encoding='utf-8', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=["Date", "Topic", "Readings/Preparation", "Work Due"])
                writer.writeheader()
                writer.writerows(data)
        else:
            # Handle Excel export
            import pandas as pd
            df = pd.DataFrame(data)
            df.to_excel(file_path, index=False)
            
    def generate_pdf(self, export_path, content):
        """Generate PDF directly using ReportLab"""
        doc = SimpleDocTemplate(export_path, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []
        
        # Custom styles
        styles.add(ParagraphStyle(
            name='CustomTitle',
            parent=styles['Heading1'],
            fontSize=16,
            spaceAfter=30,
            alignment=1  # Center alignment
        ))
        
        styles.add(ParagraphStyle(
            name='CustomHeading',
            parent=styles['Heading1'],
            fontSize=14,
            spaceAfter=12
        ))
        
        styles.add(ParagraphStyle(
            name='CustomBody',
            parent=styles['Normal'],
            fontSize=11,
            spaceAfter=6
        ))
        
        # Title
        title = f"{self.entry_course_num.get()}: {self.entry_course_title.get()}"
        story.append(Paragraph(title, styles['CustomTitle']))
        
        # Term and credits
        term_line = f"{self.entry_term.get()} ({self.entry_credits.get()} credits)"
        story.append(Paragraph(term_line, styles['CustomTitle']))
        story.append(Spacer(1, 12))
        
        # General Information
        story.append(Paragraph("I. General Information", styles['CustomHeading']))
        story.append(Paragraph(f"Meeting days and times: {self.entry_meeting_times.get()}", styles['CustomBody']))
        story.append(Paragraph(f"Class location: {self.entry_location.get()}", styles['CustomBody']))
        story.append(Spacer(1, 12))
        
        # Instructor Information
        story.append(Paragraph("Instructor:", styles['CustomBody']))
        instructor_info = [
            ("Name:", self.entry_instr_name.get()),
            ("Office:", self.entry_instr_office.get()),
            ("Phone:", self.entry_instr_phone.get()),
            ("Email:", self.entry_instr_email.get()),
            ("Office Hours:", self.entry_instr_office_hours.get())
        ]
        
        for label, value in instructor_info:
            story.append(Paragraph(f"    {label} {value}", styles['CustomBody']))
        
        # Teaching Assistants
        if self.ta_entries:
            story.append(Spacer(1, 12))
            story.append(Paragraph("Teaching Assistants:", styles['CustomBody']))
            for ta in self.ta_entries:
                story.append(Paragraph(f"    Name: {ta[0].get()}", styles['CustomBody']))
                story.append(Paragraph(f"    Email: {ta[1].get()}", styles['CustomBody']))
                story.append(Paragraph(f"    Office Hours: {ta[2].get()}", styles['CustomBody']))
                story.append(Spacer(1, 6))
        
        # Course Description
        story.append(Paragraph("Course Description", styles['CustomHeading']))
        story.append(Paragraph(self.txt_description.get("1.0", tk.END).strip(), styles['CustomBody']))
        
        # Schedule
        if self.schedule_entries:
            story.append(Paragraph("Course Schedule", styles['CustomHeading']))
            schedule_data = [['Date', 'Topic', 'Readings', 'Work Due']]
            for entry in self.schedule_entries:
                schedule_data.append([
                    entry['date'].get(),
                    entry['topic'].get(),
                    entry['readings'].get("1.0", tk.END).strip(),
                    entry['work_due'].get()
                ])
            
            schedule_table = Table(schedule_data, colWidths=[1*inch, 2*inch, 2.5*inch, 1.5*inch])
            schedule_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ]))
            story.append(schedule_table)
        
        # Build the PDF
        doc.build(story)

    def generate_syllabus(self):
        """Generate the final syllabus document"""
        if not self.validate_inputs():
            return

        # Get export path
        file_types = [("PDF Document", "*.pdf"), ("Word Document", "*.docx")]
        
        export_path = filedialog.asksaveasfilename(
            defaultextension=".pdf" if self.var_pdf.get() else ".docx",
            filetypes=file_types,
            title="Save Syllabus As"
        )
        
        if not export_path:
            return
        
        try:
            if export_path.lower().endswith('.pdf'):
                # Generate PDF directly
                self.generate_pdf(export_path, self.gather_content())
                messagebox.showinfo("Success", f"Syllabus saved as PDF: {export_path}")
            else:
                # Generate Word document
                doc = self.create_syllabus_document()
                doc.save(export_path)
                messagebox.showinfo("Success", f"Syllabus saved as: {export_path}")
                
        except Exception as e:
            messagebox.showerror("Error", 
                f"Failed to save syllabus:\n{str(e)}\n\n"
                "Please make sure you have write permissions and the file is not open in another program.")
            
    def validate_inputs(self):
        """Validate required inputs before generating syllabus"""
        required_fields = [
            (self.entry_course_num, "Course Number"),
            (self.entry_course_title, "Course Title"),
            (self.entry_term, "Term"),
            (self.entry_credits, "Credits"),
            (self.entry_meeting_times, "Meeting Times"),
            (self.entry_location, "Location"),
            (self.entry_instr_name, "Instructor Name"),
            (self.entry_instr_email, "Instructor Email")
        ]
        
        for field, name in required_fields:
            if not field.get().strip():
                messagebox.showerror("Error", f"{name} is required.")
                return False
                
        return True
        
    def create_syllabus_document(self):
        """Create the Word document for the syllabus following the exact AMH2020 format"""
        doc = Document()
        
        # Title and Course Info
        doc.add_heading(f"{self.entry_course_num.get()}: {self.entry_course_title.get()}", level=0)
        title_paragraph = doc.paragraphs[-1]
        title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Term and credits
        term_line = f"{self.entry_term.get()} ({self.entry_credits.get()} credits)"
        p = doc.add_paragraph(term_line)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_paragraph()  # Spacing
        
        # I. General Information
        doc.add_heading("I. General Information", level=1)
        p = doc.add_paragraph()
        p.add_run("Meeting days and times: ").bold = True
        p.add_run(self.entry_meeting_times.get())
        p = doc.add_paragraph()
        p.add_run("Class location: ").bold = True
        p.add_run(self.entry_location.get())
        
        doc.add_paragraph()  # Spacing
        
        # Instructor Information
        p = doc.add_paragraph()
        p.add_run("Instructor:").bold = True
        doc.add_paragraph()  # Spacing
        
        instructor_info = [
            ("Name:", self.entry_instr_name.get()),
            ("Office:", self.entry_instr_office.get()),
            ("Phone:", self.entry_instr_phone.get()),
            ("Email:", self.entry_instr_email.get()),
            ("Office Hours:", self.entry_instr_office_hours.get())
        ]
        
        for label, value in instructor_info:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.5)
            p.add_run(f"{label} ").bold = True
            p.add_run(value)
        
        # Teaching Assistants
        if self.ta_entries:
            doc.add_paragraph()  # Spacing
            p = doc.add_paragraph()
            p.add_run("Teaching Assistants:").bold = True
            doc.add_paragraph()  # Spacing
            
            for ta in self.ta_entries:
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.5)
                p.add_run(f"Name: ").bold = True
                p.add_run(ta[0].get())
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.5)
                p.add_run(f"Email: ").bold = True
                p.add_run(ta[1].get())
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.5)
                p.add_run(f"Office Hours: ").bold = True
                p.add_run(ta[2].get())
                doc.add_paragraph()  # Spacing
        
        # Course Description
        doc.add_heading("Course Description", level=2)
        doc.add_paragraph(course_description_default)
        
        # Prerequisites
        doc.add_heading("Prerequisites", level=2)
        doc.add_paragraph(prerequisites_default)
        
        # General Education Designation
        doc.add_heading("General Education Designation: Social and Behavioral Sciences (S)", level=2)
        doc.add_paragraph(gen_ed_default)
        
        # Course Objectives
        doc.add_heading("Course Objectives", level=2)
        doc.add_paragraph(course_objectives_default)
        
        # II. Student Learning Outcomes
        doc.add_heading("II. Student Learning Outcomes", level=1)
        doc.add_paragraph(student_learning_outcomes_default)
        
        # III. Graded Work
        doc.add_heading("III. Graded Work", level=1)
        
        # Add grading categories
        for category in self.category_frames:
            name = category["name"].get()
            weight = category["weight"].get()
            p = doc.add_paragraph()
            p.add_run(f"{name} ({weight}%): ").bold = True
            
            # Add assignments within category if any
            if category["assignments"]:
                for assignment in category["assignments"]:
                    p = doc.add_paragraph()
                    p.paragraph_format.left_indent = Inches(0.5)
                    p.add_run(f"{assignment['title'].get()}: Due {assignment['due date'].get()} ({assignment['points'].get()} points)")
        
        # Grading Scale
        doc.add_heading("Grading Scale", level=2)
        grades = [
            ("A", "100-93"),
            ("A-", "92-90"),
            ("B+", "89-87"),
            ("B", "86-83"),
            ("B-", "82-80"),
            ("C+", "79-77"),
            ("C", "76-73"),
            ("C-", "72-70"),
            ("D+", "69-67"),
            ("D", "66-63"),
            ("D-", "62-60"),
            ("E", "59-0")
        ]
        
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Table Grid'
        header_cells = table.rows[0].cells
        header_cells[0].text = 'Letter Grade'
        header_cells[1].text = 'Number Grade'
        
        for letter, number in grades:
            row_cells = table.add_row().cells
            row_cells[0].text = letter
            row_cells[1].text = number
        
        # IV. Evaluations
        doc.add_heading("IV. Evaluations", level=1)
        doc.add_paragraph(evaluations_default)
        
        # V. University Policies and Resources
        doc.add_heading("V. University Policies and Resources", level=1)
        
        # Add each policy section
        policies = [
            ("Students requiring accommodation", accommodations_default),
            ("University Honesty Policy", honesty_plagiarism_default),
            ("In-class recording", recording_policy_default),
            ("Procedure for conflict resolution", conflict_resolution_default),
            ("Campus Resources", campus_resources_default)
        ]
        
        for title, content in policies:
            doc.add_heading(title, level=2)
            doc.add_paragraph(content)
        
        # VI. Calendar
        doc.add_heading("VI. Calendar", level=1)
        
        # Create calendar table
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        header_cells = table.rows[0].cells
        header_cells[0].text = 'Date'
        header_cells[1].text = 'Topic'
        header_cells[2].text = 'Readings/Preparation'
        header_cells[3].text = 'Work Due'
        
        # Add schedule entries
        for entry in self.schedule_entries:
            row_cells = table.add_row().cells
            row_cells[0].text = entry['date'].get()
            row_cells[1].text = entry['topic'].get()
            row_cells[2].text = entry['readings'].get("1.0", tk.END).strip()
            row_cells[3].text = entry['work_due'].get()
        
        return doc
        
    def run(self):
        """Start the application"""
        self.root.mainloop()

    def add_category(self):
        """Add a new empty grading category"""
        frame = ttk.LabelFrame(self.categories_frame)
        frame.pack(fill=tk.X, pady=5)
        
        # Category header
        header_frame = ttk.Frame(frame)
        header_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(header_frame, text="Category Name:").pack(side=tk.LEFT)
        name_entry = ttk.Entry(header_frame, width=30)
        name_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(header_frame, text="Weight (%):").pack(side=tk.LEFT)
        weight_entry = ttk.Entry(header_frame, width=5)
        weight_entry.pack(side=tk.LEFT, padx=5)
        
        # Description
        ttk.Label(frame, text="Description:").pack(anchor="w", padx=5)
        desc_text = scrolledtext.ScrolledText(frame, width=60, height=4, wrap=tk.WORD)
        desc_text.pack(fill=tk.X, padx=5, pady=5)
        
        # Assignments section
        assignments_frame = ttk.Frame(frame)
        assignments_frame.pack(fill=tk.X, padx=5, pady=5)
        
        assignments = []
        
        def add_assignment():
            assignment_frame = ttk.Frame(assignments_frame)
            assignment_frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(assignment_frame, text="Title:").pack(side=tk.LEFT)
            title_entry = ttk.Entry(assignment_frame, width=30)
            title_entry.pack(side=tk.LEFT, padx=2)
            
            ttk.Label(assignment_frame, text="Due:").pack(side=tk.LEFT)
            due_entry = ttk.Entry(assignment_frame, width=15)
            due_entry.pack(side=tk.LEFT, padx=2)
            
            ttk.Label(assignment_frame, text="Points:").pack(side=tk.LEFT)
            points_entry = ttk.Entry(assignment_frame, width=5)
            points_entry.pack(side=tk.LEFT, padx=2)
            
            def remove_assignment():
                assignment_frame.destroy()
                assignments.remove(assignment_dict)
            
            remove_btn = ttk.Button(assignment_frame, text="×", 
                                  command=remove_assignment,
                                  style="Delete.TButton")
            remove_btn.pack(side=tk.LEFT, padx=2)
            
            assignment_dict = {
                "frame": assignment_frame,
                "title": title_entry,
                "due date": due_entry,
                "points": points_entry
            }
            assignments.append(assignment_dict)
        
        ttk.Button(frame, text="Add Assignment", command=add_assignment).pack(anchor="w", padx=5, pady=5)
        
        def remove_category():
            frame.destroy()
            self.category_frames.remove(category_dict)
        
        ttk.Button(frame, text="Remove Category", command=remove_category).pack(anchor="w", padx=5, pady=5)
        
        category_dict = {
            "frame": frame,
            "name": name_entry,
            "weight": weight_entry,
            "description": desc_text,
            "assignments": assignments
        }
        
        self.category_frames.append(category_dict)
        return category_dict

    def gather_content(self):
        """Gather all content from the form for syllabus generation"""
        content = {
            "course_info": {
                "number": self.entry_course_num.get(),
                "title": self.entry_course_title.get(),
                "term": self.entry_term.get(),
                "credits": self.entry_credits.get(),
                "meeting_times": self.entry_meeting_times.get(),
                "location": self.entry_location.get(),
                "description": self.txt_description.get("1.0", tk.END).strip()
            },
            "instructor_info": {
                "name": self.entry_instr_name.get(),
                "office": self.entry_instr_office.get(),
                "phone": self.entry_instr_phone.get(),
                "email": self.entry_instr_email.get(),
                "office_hours": self.entry_instr_office_hours.get()
            },
            "tas": [
                {
                    "name": ta[0].get(),
                    "email": ta[1].get(),
                    "office_hours": ta[2].get()
                }
                for ta in self.ta_entries
            ],
            "schedule": [
                {
                    "date": entry["date"].get(),
                    "topic": entry["topic"].get(),
                    "readings": entry["readings"].get("1.0", tk.END).strip(),
                    "work_due": entry["work_due"].get()
                }
                for entry in self.schedule_entries
            ],
            "grading": [
                {
                    "name": cat["name"].get(),
                    "weight": cat["weight"].get(),
                    "description": cat["description"].get("1.0", tk.END).strip(),
                    "assignments": [
                        {
                            "title": assg["title"].get(),
                            "due": assg["due date"].get(),
                            "points": assg["points"].get()
                        }
                        for assg in cat["assignments"]
                    ]
                }
                for cat in self.category_frames
            ],
            "policies": {
                policy: var.get()
                for policy, var in self.optional_policies.items()
            }
        }
        
        # Add materials info if available
        if hasattr(self, 'materials_text'):
            content["materials"] = {
                "required": self.materials_text.get("1.0", tk.END).strip(),
                "fee": self.fee_entry.get() if hasattr(self, 'fee_entry') else ""
            }
            
        # Add policy details if available
        if hasattr(self, 'late_policy_var'):
            content["policies"]["late_policy"] = self.late_policy_var.get()
            content["policies"]["late_policy_text"] = self.late_policy_text.get("1.0", tk.END).strip()
            
        if hasattr(self, 'extra_credit_var'):
            content["policies"]["extra_credit_policy"] = self.extra_credit_var.get()
            content["policies"]["extra_credit_text"] = self.extra_credit_text.get("1.0", tk.END).strip()
            
        return content

if __name__ == "__main__":
    app = HistorySyllabusGenerator()
    app.run()
