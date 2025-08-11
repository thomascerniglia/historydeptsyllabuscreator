"""
Document Preview Module for History Syllabus Generator
Contains all document preview methods and utilities
"""

import tkinter as tk
from tkinter import ttk, scrolledtext
import webbrowser
from constants import *

class DocumentPreviewMixin:
    """Mixin class containing all document preview methods"""
    
    def create_document_preview_tab(self):
        """Create a tab for previewing the entire document as it will appear in final form"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Document Preview")
        if hasattr(self, 'add_mousewheel_scrolling'):
            self.add_mousewheel_scrolling(tab, tab)
        
        # Create a scrollable preview area
        preview_frame = ttk.Frame(tab)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Controls at the top
        control_frame = ttk.Frame(preview_frame)
        control_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(control_frame, text="Document Preview", style="Heading.TLabel").pack(side=tk.LEFT, padx=5)
        
        refresh_btn = ttk.Button(control_frame, text="Refresh Preview", 
                              command=self.update_document_preview,
                              style="Action.TButton")
        refresh_btn.pack(side=tk.RIGHT, padx=5)
        
        # Create a canvas with scrollbar for the preview content
        canvas_frame = ttk.Frame(preview_frame, relief=tk.SUNKEN, borderwidth=1)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.preview_canvas = tk.Canvas(canvas_frame, bg="white")
        scrollbar_y = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.preview_canvas.yview)
        scrollbar_x = ttk.Scrollbar(canvas_frame, orient="horizontal", command=self.preview_canvas.xview)
        
        self.preview_content_frame = ttk.Frame(self.preview_canvas, style="Preview.TFrame")
        self.preview_content_frame.bind(
            "<Configure>",
            lambda e: self.preview_canvas.configure(
                scrollregion=self.preview_canvas.bbox("all")
            )
        )
        self.preview_canvas.create_window((0, 0), window=self.preview_content_frame, anchor="nw")
        self.preview_canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        # Pack the scrollbars and canvas
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.preview_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Add mousewheel scrolling
        if hasattr(self, 'add_mousewheel_scrolling'):
            self.add_mousewheel_scrolling(self.preview_canvas, self.preview_canvas)
            self.add_mousewheel_scrolling(self.preview_content_frame, self.preview_canvas)
        
        # Bind mousewheel events to the preview content so the canvas scrolls
        self.preview_canvas.bind("<Enter>", lambda e: self.preview_canvas.focus_set())
        self.preview_canvas.bind("<MouseWheel>", lambda event: self.preview_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units"))
        
        # Initial preview generation
        self.update_document_preview()

    def update_document_preview(self):
        """Update the document preview with current content"""
        # Clear previous content
        for widget in self.preview_content_frame.winfo_children():
            widget.destroy()

        try:
            # Create container for all preview sections
            content_container = ttk.Frame(self.preview_content_frame, style="Preview.TFrame")
            content_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

            # Title and Course Info
            title_text = f"{self.entry_course_num.get()}: {self.entry_course_title.get()}"
            title = ttk.Label(content_container, text=title_text,
                           font=("Times New Roman", 14, "bold"),
                           style="Preview.TLabel")
            title.pack(pady=5)

            term_text = f"{self.entry_term.get()} ({self.entry_credits.get()} credits)"
            term = ttk.Label(content_container, text=term_text,
                          font=("Times New Roman", 12),
                          style="Preview.TLabel")
            term.pack(pady=2)

            ttk.Separator(content_container).pack(fill=tk.X, pady=10)

            # I. General Information
            self._add_preview_section(content_container, "I. General Information", 12, "bold")

            self._add_preview_field(content_container, "Meeting days and times:", self.entry_meeting_times.get())
            self._add_preview_field(content_container, "Class location:", self.entry_location.get())

            # Instructor info
            self._add_preview_text(content_container, "\nInstructor:", bold=True)

            instructor_frame = ttk.Frame(content_container, style="Preview.TFrame")
            instructor_frame.pack(fill=tk.X, padx=20, pady=2)

            instructor_info = [
                ("Name:", self.entry_instr_name.get()),
                ("Office:", self.entry_instr_office.get()),
                ("Phone:", self.entry_instr_phone.get()),
                ("Email:", self.entry_instr_email.get()),
                ("Office Hours:", self.entry_instr_office_hours.get())
            ]

            for i, (label, value) in enumerate(instructor_info):
                info_line = ttk.Frame(instructor_frame, style="Preview.TFrame")
                info_line.pack(fill=tk.X, pady=1)
                ttk.Label(info_line, text=label, font=("Times New Roman", 10, "bold"),
                         width=12, anchor="w", background="white").pack(side=tk.LEFT)
                ttk.Label(info_line, text=value, font=("Times New Roman", 10),
                         background="white").pack(side=tk.LEFT)

            # Sections
            if hasattr(self, 'ta_entries') and self.ta_entries:
                self._add_preview_text(content_container, "\nSections:", bold=True)
                
                ta_frame = ttk.Frame(content_container, style="Preview.TFrame")
                ta_frame.pack(fill=tk.X, padx=20, pady=2)
                
                for ta in self.ta_entries:
                    if len(ta) >= 5:
                        ta_info = [
                            ("TA Name:", ta[0].get()),
                            ("Email:", ta[1].get()),
                            ("Office Hours:", ta[2].get()),
                            ("Section Meeting Place:", ta[3].get()),
                            ("Section Meeting Time:", ta[4].get())
                        ]
                        
                        for label, value in ta_info:
                            info_line = ttk.Frame(ta_frame, style="Preview.TFrame")
                            info_line.pack(fill=tk.X, pady=1)
                            ttk.Label(info_line, text=label, font=("Times New Roman", 10, "bold"),
                                     width=12, anchor="w", background="white").pack(side=tk.LEFT)
                            ttk.Label(info_line, text=value, font=("Times New Roman", 10),
                                     background="white").pack(side=tk.LEFT)

            # Course Description
            self._add_preview_section(content_container, "Course Description", 12, "bold")
            self._add_preview_text_with_link(content_container, self.txt_description.get("1.0", tk.END).strip())

            # Prerequisites
            self._add_preview_section(content_container, "Prerequisites", 12, "bold")
            prerequisites = self.entry_prerequisites.get().strip() if hasattr(self, 'entry_prerequisites') else "None"
            self._add_preview_text(content_container, prerequisites or "None")
            
            # Gen Ed (if applicable)
            if hasattr(self, 'show_gen_ed') and self.show_gen_ed.get():
                self._add_preview_section(content_container, "General Education Designation: Social and Behavioral Sciences (S)", 12, "bold")
                self._add_preview_text(content_container, gen_ed_default)
                
                course_num = self.entry_course_num.get()
                completion_text = f"Your successful completion of {course_num} with a grade of \"C\" or higher will count towards UF's General Education State Core in Social and Behavioral Sciences (S). It will also count towards the State of Florida's Civic Literacy requirement."
                self._add_preview_text(content_container, completion_text)
            
            # Course Objectives
            self._add_preview_section(content_container, "Course Objectives", 12, "bold")
            objectives_link_frame = ttk.Frame(content_container, style="Preview.TFrame")
            objectives_link_frame.pack(fill=tk.X, pady=2)
            self._add_preview_text(objectives_link_frame, "All General Education area objectives can be found ", end="")
            link_label = ttk.Label(objectives_link_frame, text="here", foreground="blue", cursor="hand2", 
                                  font=("Times New Roman", 10, "underline"))
            link_label.pack(side=tk.LEFT)
            link_label.bind("<Button-1>", lambda e: webbrowser.open("https://undergrad.aa.ufl.edu/general-education/gen-ed-program/subject-area-objectives/"))
            self._add_preview_text(objectives_link_frame, ".", start="")
            
            if hasattr(self, 'objective_entries') and self.objective_entries:
                objectives_frame = ttk.Frame(content_container, style="Preview.TFrame")
                objectives_frame.pack(fill=tk.X, padx=10, pady=2)
                for i, obj in enumerate(self.objective_entries, 1):
                    obj_text = obj["entry"].get().strip()
                    if obj_text:
                        self._add_preview_text(objectives_frame, f"{i}. {obj_text}")
            
            # II. Student Learning Outcomes
            self._add_preview_section(content_container, "II. Student Learning Outcomes", 12, "bold")
            self._add_preview_text(content_container, "A student who successfully completes this course will:")
            
            outcomes_frame = ttk.Frame(content_container, style="Preview.TFrame")
            outcomes_frame.pack(fill=tk.X, padx=10, pady=2)
            
            if hasattr(self, 'outcome_entries') and self.outcome_entries:
                for i, outcome_entry in enumerate(self.outcome_entries, 1):
                    outcome_text = outcome_entry["entry"].get().strip()
                    if outcome_text:
                        self._add_preview_text(outcomes_frame, f"{i}. {outcome_text}")
            else:
                default_outcomes = [
                    "Describe the factual details of the substantive historical episodes under study.",
                    "Identify and analyze foundational developments that shaped history using critical thinking skills.",
                    "Demonstrate an understanding of the primary ideas, values, and perceptions that have shaped history.",
                    "Demonstrate competency in civic literacy."
                ]
                for i, outcome in enumerate(default_outcomes, 1):
                    self._add_preview_text(outcomes_frame, f"{i}. {outcome}")
            
            # Learning Objectives Table
            if hasattr(self, 'learning_objectives_entries') and self.learning_objectives_entries:
                self._add_preview_section(content_container, "Objectives—General Education and Social and Behavioral Sciences (S)", 11, "bold")
                
                # Create table frame
                table_frame = ttk.Frame(content_container, relief=tk.SOLID, borderwidth=1)
                table_frame.pack(fill=tk.X, padx=10, pady=5)
                
                # Headers
                headers = ["CATEGORY", "SOCIAL SCIENCE SLOS", "STATE SLO ASSIGNMENTS", "COURSE-SPECIFIC"]
                for i, header in enumerate(headers):
                    header_cell = ttk.Label(table_frame, text=header, 
                                          background="#808080", foreground="white",
                                          font=("Arial", 9, "bold"))
                    header_cell.grid(row=0, column=i, sticky="nsew", padx=1, pady=1)
                    table_frame.columnconfigure(i, weight=1 if i == 0 else 2)
                
                # Content rows
                row = 1
                for category, entries in self.learning_objectives_entries.items():
                    if 'frame' in entries and not entries['frame'].winfo_exists():
                        continue
                    
                    # Get text from entries:
                    category_text = entries['category'].get() if 'category' in entries else category
                    slo_text = entries['slo'].get("1.0", tk.END).strip()
                    assignments_text = entries['assignments'].get("1.0", tk.END).strip()
                    course_specific_text = entries['course_specific'].get("1.0", tk.END).strip()
                    
                    # Create cells
                    cells = [
                        (category_text, 120),
                        (slo_text, 200),
                        (assignments_text, 200),
                        (course_specific_text, 200)
                    ]
                    for col, (text, wraplength) in enumerate(cells):
                        cell = ttk.Label(table_frame, text=text, background="white",
                                       relief="solid", borderwidth=1, wraplength=wraplength,
                                       padding=5)
                        cell.grid(row=row, column=col, sticky="nsew", padx=1, pady=1)
                    row += 1
            
            # III. Graded Work
            self._add_preview_section(content_container, "III. Graded Work", 12, "bold")
            
            # Required Materials
            if hasattr(self, 'materials_text') and self.materials_text.get("1.0", tk.END).strip():
                self._add_preview_section(content_container, "Required Materials", 11, "bold")
                materials_text = self.materials_text.get("1.0", tk.END).strip()
                self._add_preview_text_with_link(content_container, materials_text)
                
                # Materials Fee
                fee_value = self.fee_entry.get() if hasattr(self, 'fee_entry') else "0.00"
                self._add_preview_text(content_container, f"\nMaterials Fee: ${fee_value}", bold=True)
            
            # Grading Components
            if hasattr(self, 'category_frames') and self.category_frames:
                self._add_preview_section(content_container, "Grading Components", 11, "bold")
                
                # Create table for grade components
                grade_frame = ttk.Frame(content_container, style="Preview.TFrame")
                grade_frame.pack(fill=tk.X, padx=10, pady=5)
                
                # Header for grade table
                ttk.Label(grade_frame, text="Category", font=("Arial", 10, "bold"), width=25).grid(
                    row=0, column=0, sticky="w", padx=5, pady=2)
                ttk.Label(grade_frame, text="Weight", font=("Arial", 10, "bold"), width=10).grid(
                    row=0, column=1, sticky="w", padx=5, pady=2)
                ttk.Separator(grade_frame, orient=tk.HORIZONTAL).grid(
                    row=1, column=0, columnspan=2, sticky="ew", pady=2)
                
                # Add each category
                row = 2
                for i, category in enumerate(self.category_frames):
                    name = category["name"].get().strip()
                    weight = category["weight"].get().strip()
                    if name and weight:
                        ttk.Label(grade_frame, text=name, font=("Arial", 10), width=25).grid(
                            row=row, column=0, sticky="w", padx=5, pady=1)
                        ttk.Label(grade_frame, text=f"{weight}%", font=("Arial", 10), width=10).grid(
                            row=row, column=1, sticky="w", padx=5, pady=1)
                        row += 1
                
                # Check for descriptions and assignments
                for category in self.category_frames:
                    name = category["name"].get().strip()
                    desc = category["description"].get("1.0", tk.END).strip()
                    
                    if name and desc:
                        self._add_preview_text(content_container, f"\n{name}: ", bold=True, end="")
                        self._add_preview_text(content_container, desc, start="")
                    
                    # Assignments for this category
                    if "assignments" in category and category["assignments"]:
                        has_assignments = False
                        for assignment in category["assignments"]:
                            title = assignment["title"].get().strip()
                            due_date = assignment["due date"].get().strip()
                            points = assignment["points"].get().strip()
                            description = assignment["description"].get("1.0", tk.END).strip() if "description" in assignment else ""
                            
                            if title:
                                if not has_assignments:
                                    self._add_preview_text(content_container, f"{name} Assignments:", bold=True)
                                    has_assignments = True
                                
                                assignment_text = f"• {title}"
                                if due_date:
                                    assignment_text += f" (Due: {due_date})"
                                if points:
                                    assignment_text += f" - {points} points"
                                
                                self._add_preview_text(content_container, assignment_text, indent=10)
                                
                                if description:
                                    self._add_preview_text_with_link(content_container, description, indent=20)
            
            # Grading Scale
            self._add_preview_section(content_container, "Grading Scale", 11, "bold")
            
            # Create grading scale data
            grades = [
                ["Letter Grade", "Number Grade"],
                ["A", "100-93"],
                ["A-", "92-90"],
                ["B+", "89-87"],
                ["B", "86-83"],
                ["B-", "82-80"],
                ["C+", "79-77"],
                ["C", "76-73"],
                ["C-", "72-70"],
                ["D+", "69-67"],
                ["D", "66-63"],
                ["D-", "62-60"],
                ["E", "59-0"]
            ]
            
            # Create table frame
            table_frame = ttk.Frame(content_container, borderwidth=1, relief=tk.SOLID)
            table_frame.pack(fill=tk.X, pady=5, padx=10)
            
            # Table headers
            header_row = ttk.Frame(table_frame, style="Preview.TFrame")
            header_row.pack(fill=tk.X)
            
            for i, header in enumerate(grades[0]):
                cell = ttk.Label(header_row, text=header, background="#808080", foreground="white",
                              font=("Arial", 9, "bold"), padding=5)
                cell.grid(row=0, column=i, sticky="nsew", padx=1, pady=1)
                header_row.columnconfigure(i, weight=1)
            
            # Table content rows
            for i, grade in enumerate(grades[1:]):
                data_row = ttk.Frame(table_frame, style="Preview.TFrame")
                data_row.pack(fill=tk.X)
                
                for j, value in enumerate(grade):
                    cell = ttk.Label(data_row, text=value, background="white",
                                   font=("Arial", 9), padding=5, relief="solid", borderwidth=1)
                    cell.grid(row=0, column=j, sticky="nsew", padx=1, pady=1)
                    data_row.columnconfigure(j, weight=1)
            
            # Add the grading policy text with a hyperlink
            policy_text = "See the UF Catalog's \"Grades and Grading Policies\" for information on how UF assigns grade points."
            link = "https://catalog.ufl.edu/UGRD/academic-regulations/grades-grading-policies/"
            self._add_preview_text(content_container, policy_text)
            
            # Add rounding statement if enabled
            if hasattr(self, 'grading_rounding_var') and self.grading_rounding_var.get():
                from constants import grading_rounding_default
                self._add_preview_text(content_container, grading_rounding_default)
            
            # Add the note about minimum grades
            note_text = "Note: A minimum grade of C is required to earn General Education credit."
            self._add_preview_text(content_container, note_text)
            
            # Add Instructions for Submitting Written Assignments section
            self._add_preview_section(content_container, "Instructions for Submitting Written Assignments", 12, "bold")
            self._add_preview_text(content_container, "All written assignments must be submitted as Word documents (.doc or .docx) through the \"Assignments\" portal in Canvas by the specified deadlines. Do NOT send assignments as PDF files.")
                        
            # Add Late Submissions policy (show content from the input)
            if hasattr(self, 'late_submissions_policy_var') and self.late_submissions_policy_var.get():
                self._add_preview_section(content_container, "Late Submissions", 11, "bold")
                if hasattr(self, 'late_policy_text') and self.late_policy_text.get("1.0", tk.END).strip():
                    # Display the actual content from the late policy text input
                    self._add_preview_text_with_link(content_container, self.late_policy_text.get("1.0", tk.END).strip())
                elif hasattr(self, 'late_policy_var') and hasattr(self, 'late_policies') and self.late_policy_var.get() in self.late_policies:
                    # If there's a selected policy from the dropdown but no custom text, use the dropdown's standard text
                    selected_policy = self.late_policy_var.get()
                    self._add_preview_text(content_container, self.late_policies[selected_policy])
                else:
                    # Default text only if no input and no selection
                    self._add_preview_text(content_container, "Late submission policy not specified.")
            
            # Add Extra Credit policy (using the exact same approach as Late Submissions)
            if hasattr(self, 'extra_credit_policy_var') and self.extra_credit_policy_var.get():
                self._add_preview_section(content_container, "Extra Credit", 11, "bold")
                if hasattr(self, 'extra_credit_text') and self.extra_credit_text.get("1.0", tk.END).strip():
                    # Display the actual content from the extra credit text input
                    self._add_preview_text(content_container, self.extra_credit_text.get("1.0", tk.END).strip())
                elif hasattr(self, 'extra_credit_var') and hasattr(self, 'extra_credit_policies') and self.extra_credit_var.get() in self.extra_credit_policies:
                    # If there's a selected policy from the dropdown but no custom text, use the dropdown's standard text
                    selected_policy = self.extra_credit_var.get()
                    self._add_preview_text(content_container, self.extra_credit_policies[selected_policy])
                else:
                    # Default text only if no input and no selection
                    self._add_preview_text(content_container, "Extra credit policy not specified.")
            
            # Canvas Policy (always included)
            if hasattr(self, 'canvas_policy_var') and self.canvas_policy_var.get():
                self._add_preview_section(content_container, "Canvas", 11, "bold")
                if hasattr(self, 'canvas_policy_text') and self.canvas_policy_text.get("1.0", tk.END).strip() != "-Replace with your Canvas Policies-":
                    self._add_preview_text_with_link(content_container, self.canvas_policy_text.get("1.0", tk.END).strip())
                else:
                    self._add_preview_text(content_container, canvas_policy_default)
            
            # Add Technology Policy
            if hasattr(self, 'technology_policy_var') and self.technology_policy_var.get():
                self._add_preview_section(content_container, "Technology Policy", 11, "bold")
                if hasattr(self, 'technology_policy_text'):
                    self._add_preview_text_with_link(content_container, self.technology_policy_text.get("1.0", "end").strip())
            
            # Add Communication Policy
            if hasattr(self, 'communication_policy_var') and self.communication_policy_var.get():
                self._add_preview_section(content_container, "Communication Policy", 11, "bold")
                if hasattr(self, 'communication_policy_text'):
                    self._add_preview_text_with_link(content_container, self.communication_policy_text.get("1.0", "end").strip())
            
            # Add Assignment Support section if enabled
            if hasattr(self, 'outside_support_var') and self.outside_support_var.get():
                self._add_preview_section(content_container, "Assignment Support Outside the Classroom", 11, "bold")
                if hasattr(self, 'support_text'):
                    support_text = self.support_text.get("1.0", tk.END).strip()
                    self._add_preview_text_with_link(content_container, support_text)
            
            # Add "IV. University Policies and Resources" section (formerly V.)
            self._add_preview_section(content_container, "IV. University Policies and Resources", 12, "bold")
            
            # Check if simplified policies are enabled
            if hasattr(self, 'use_simplified_policies_var') and self.use_simplified_policies_var.get():
                # Use simplified UF policies with clickable link
                from constants import uf_policy_simplified
                self._add_preview_text_with_link(content_container, uf_policy_simplified)
            else:
                # Use original detailed policies
                # Students requiring accommodation
                self._add_preview_section(content_container, "Students requiring accommodation", 11, "bold")
                accommodations_text = (
                    "Students with disabilities who experience learning barriers and would like to request academic accommodations "
                    "should connect with the Disability Resource Center by visiting https://disability.ufl.edu/students/get-started/. "
                    "It is important for students to share their accommodation letter with the instructor and discuss their "
                    "access needs as early as possible in the semester."
                )
                self._add_preview_text_with_link(content_container, accommodations_text)
                
                # University Honesty Policy
                self._add_preview_section(content_container, "University Honesty Policy", 11, "bold")
                honesty_text = (
                    "UF students are bound by The Honor Pledge which states \"We, the members of the "
                    "University of Florida community, pledge to hold ourselves and our peers to the "
                    "highest standards of honor and integrity by abiding by the Honor Code.\" On all "
                    "work submitted for credit by students at the University of Florida, the "
                    "following pledge is either required or implied: \"On my honor, I have neither "
                    "given nor received unauthorized aid in doing this assignment.\" The Conduct Code "
                    "specifies a number of behaviors that are in violation of this code and the "
                    "possible sanctions. See the UF Conduct Code website for more information. If "
                    "you have any questions or concerns, please consult with the instructor or TAs "
                    "in this class."
                )
                self._add_preview_text(content_container, honesty_text)
                
                # Plagiarism section
                self._add_preview_section(content_container, "Plagiarism and Related Ethical Violations", 11, "bold")
                plagiarism_text = (
                    "Ethical violations such as plagiarism, cheating, academic misconduct (e.g. passing off others' work as your own, reusing old assignments, etc.) "
                    "will not be tolerated and will result in a failing grade in this course. Students must be especially wary of plagiarism. "
                    "The UF Student Honor Code defines plagiarism as follows: "
                    "A student shall not represent as the student's own work all or any portion of the work of another. "
                    "Plagiarism includes (but is not limited to): a. Quoting oral or written materials, whether published or unpublished, without proper attribution. "
                    "b. Submitting a document or assignment which in whole or in part is identical or substantially identical to a document or assignment not authored by the student."
                    " Note that plagiarism also includes the use of any artificial intelligence programs, such as ChatGPT."
                )
                self._add_preview_text(content_container, plagiarism_text)

            # V. Course Schedule (formerly VI.)
            self._add_preview_section(content_container, "V. Calendar", 12, "bold")
            
            # Schedule Table Preview
            if hasattr(self, 'schedule_entries') and self.schedule_entries:
                schedule_table_frame = ttk.Frame(content_container, style="Preview.TFrame")
                schedule_table_frame.pack(fill=tk.X, pady=5)
                
                # Header
                header_frame = ttk.Frame(schedule_table_frame, style="Preview.TFrame")
                header_frame.pack(fill=tk.X)
                ttk.Label(header_frame, text="Date", width=15, style="PreviewHeader.TLabel", borderwidth=1, relief="solid").pack(side=tk.LEFT)
                ttk.Label(header_frame, text="Topic", width=30, style="PreviewHeader.TLabel", borderwidth=1, relief="solid").pack(side=tk.LEFT)
                ttk.Label(header_frame, text="Readings/Preparation", width=50, style="PreviewHeader.TLabel", borderwidth=1, relief="solid").pack(side=tk.LEFT)
                ttk.Label(header_frame, text="Work Due", width=20, style="PreviewHeader.TLabel", borderwidth=1, relief="solid").pack(side=tk.LEFT)
                
                # Entries
                for entry in self.schedule_entries:
                    date_text = entry["date"].get().strip()
                    topic_text = entry["topic"].get().strip()
                    readings_text = entry["readings"].get("1.0", tk.END).strip()
                    work_due_text = entry["work_due"].get().strip()
                    
                    # Skip empty rows
                    if not any([date_text, topic_text, readings_text, work_due_text]):
                        continue
                    
                    entry_frame = ttk.Frame(schedule_table_frame, style="Preview.TFrame")
                    entry_frame.pack(fill=tk.X)
                    
                    ttk.Label(entry_frame, text=date_text, width=15, wraplength=100, borderwidth=1, relief="solid").pack(side=tk.LEFT)
                    ttk.Label(entry_frame, text=topic_text, width=30, wraplength=200, borderwidth=1, relief="solid").pack(side=tk.LEFT)
                    ttk.Label(entry_frame, text=readings_text, width=50, wraplength=350, borderwidth=1, relief="solid").pack(side=tk.LEFT)
                    ttk.Label(entry_frame, text=work_due_text, width=20, wraplength=150, borderwidth=1, relief="solid").pack(side=tk.LEFT)
            else:
                self._add_preview_text(content_container, "Schedule will be provided separately.")
            
            # Update scroll region after adding all content
            self.preview_content_frame.update_idletasks()
            self.preview_canvas.config(scrollregion=self.preview_canvas.bbox("all"))
        
        except Exception as e:
            print(f"Error updating document preview: {e}")
            import traceback
            traceback.print_exc()

        # At the very end of the function, after all content is added:
        try:
            # Add footer with page indicator
            footer_frame = ttk.Frame(content_container, style="Preview.TFrame")
            footer_frame.pack(fill=tk.X, pady=20)
            
            # Add note about pagination
            ttk.Label(footer_frame, 
                     text="Note: Page numbers will appear in the generated document",
                     font=("Times New Roman", 9, "italic"),
                     background="white").pack(side=tk.LEFT)
            
            # Show sample page number on right
            ttk.Label(footer_frame, 
                     text="Page X", 
                     font=("Times New Roman", 9),
                     background="white").pack(side=tk.RIGHT)
            
            # Add separator above footer
            ttk.Separator(content_container, orient=tk.HORIZONTAL).pack(
                fill=tk.X, pady=5, before=footer_frame)
        except Exception as e:
            print(f"Error adding page number preview: {e}")

    def _add_preview_section(self, parent, text, font_size=12, font_weight="normal"):
        """Add a section heading to the preview"""
        section = ttk.Label(parent, text=text,
                            font=("Times New Roman", font_size, font_weight),
                            background="white")
        section.pack(anchor="w", pady=5, fill=tk.X)
        return section

    def _add_preview_field(self, parent, label, value, indent=20):
        """Add a labeled field to the preview"""
        field_frame = ttk.Frame(parent, style="Preview.TFrame")
        field_frame.pack(fill=tk.X, padx=indent, pady=1)
        
        label_widget = ttk.Label(field_frame, text=label, 
                                 font=("Times New Roman", 10, "bold"),
                                 width=15, anchor="w",
                                 background="white")
        label_widget.pack(side=tk.LEFT)
        
        value_widget = ttk.Label(field_frame, text=value,
                                 font=("Times New Roman", 10),
                                 background="white",
                                 justify=tk.LEFT,
                                 wraplength=400)
        value_widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        return field_frame

    def _add_preview_text(self, parent, text, bold=False, indent=0, pady=2, end=None, start=None):
        """
        Add plain text to the preview
        
        Parameters:
        - parent: parent widget
        - text: text to display
        - bold: whether to bold the text
        - indent: left indent in pixels
        - pady: vertical padding
        - end: if set to "", don't add newline at end (pack with side=tk.LEFT)
        - start: if set to "", continue on same line (pack with side=tk.LEFT)
        """
        text_widget = ttk.Label(parent, text=text,
                                font=("Times New Roman", 10, bold and "bold" or "normal"),
                                background="white",
                                justify=tk.LEFT,
                                wraplength=600)
        
        # Determine how to pack the widget based on end/start parameters
        if end == "" or start == "":
            text_widget.pack(side=tk.LEFT, padx=indent, pady=pady)
        else:
            text_widget.pack(anchor="w", padx=indent, pady=pady, fill=tk.X)

    def _add_preview_text_with_link(self, parent, text, bold=False, indent=0, pady=2):
        """
        Add text to the preview with clickable links
        
        Parameters:
        - parent: parent widget
        - text: text to display (may contain URLs)
        - bold: whether to bold the text
        - indent: left indent in pixels
        - pady: vertical padding
        """
        import re
        import webbrowser
        
        # Create a frame to hold the text
        frame = tk.Frame(parent, bg="white")
        frame.pack(anchor="w", padx=indent, pady=pady, fill=tk.X)
        
        # Find URLs in the text
        url_pattern = r'(https?://[^\s]+)'
        parts = re.split(url_pattern, text)
        
        # Create a text widget to handle proper text wrapping
        text_widget = tk.Text(frame, 
                            font=("Times New Roman", 10, bold and "bold" or "normal"),
                            bg="white", 
                            wrap=tk.WORD,
                            height=1,  # Start with minimal height
                            bd=0,
                            highlightthickness=0)
        text_widget.pack(fill=tk.X, expand=True)
        
        # Insert text and create clickable links
        for i, part in enumerate(parts):
            if re.match(url_pattern, part):
                # This is a URL - make it clickable
                start_index = text_widget.index(tk.INSERT)
                text_widget.insert(tk.END, part)
                end_index = text_widget.index(tk.INSERT)
                
                # Configure the URL as clickable
                text_widget.tag_add(f"link_{i}", start_index, end_index)
                text_widget.tag_config(f"link_{i}", foreground="blue", underline=True)
                text_widget.tag_bind(f"link_{i}", "<Button-1>", lambda e, url=part: webbrowser.open(url))
                text_widget.tag_bind(f"link_{i}", "<Enter>", lambda e: text_widget.config(cursor="hand2"))
                text_widget.tag_bind(f"link_{i}", "<Leave>", lambda e: text_widget.config(cursor=""))
            else:
                # This is regular text
                if part:  # Include all parts, even if just whitespace for proper formatting
                    text_widget.insert(tk.END, part)
        
        # Make text widget read-only
        text_widget.config(state=tk.DISABLED)
        
        # Adjust height based on content
        text_widget.update()
        lines = int(text_widget.index('end-1c').split('.')[0])
        text_widget.config(height=lines)
