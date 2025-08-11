import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from tkinter import scrolledtext
from constants import *

class UITabsMixin:
    """Mixin class containing all UI tab creation methods"""
    
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

        # Add mousewheel scrolling
        self.add_mousewheel_scrolling(canvas, canvas)
        self.add_mousewheel_scrolling(scrollable_frame, canvas)
        
        return scrollable_frame
        
    def add_mousewheel_scrolling(self, widget, canvas=None):
        """Add mousewheel scrolling to a widget"""
        def _on_mousewheel(event):
            if canvas:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            else:
                widget.yview_scroll(int(-1 * (event.delta / 120)), "units")

        # Bind mousewheel on Windows/MacOS
        widget.bind("<MouseWheel>", _on_mousewheel)
        # For Linux, bind Button-4 and Button-5
        widget.bind("<Button-4>", lambda e: _on_mousewheel(type('Event', (), {'delta': 120})))
        widget.bind("<Button-5>", lambda e: _on_mousewheel(type('Event', (), {'delta': -120})))

    def add_schedule_entry(self, date="", topic="", readings="", work_due=""):
        """Add a new schedule entry with enhanced styling"""
        row = len(self.schedule_entries)
        
        # Date entry
        date_entry = ttk.Entry(self.entries_frame, width=15)
        date_entry.grid(row=row, column=0, sticky="w", padx=(5, 10), pady=2)
        date_entry.insert(0, date)  # Populate with provided data
        
        # Topic entry
        topic_entry = ttk.Entry(self.entries_frame, width=40)
        topic_entry.grid(row=row, column=1, sticky="ew", padx=5, pady=2)
        topic_entry.insert(0, topic)  # Populate with provided data
        
        # Readings frame
        readings_frame = ttk.Frame(self.entries_frame)
        readings_frame.grid(row=row, column=2, sticky="ew", padx=5, pady=2)
        readings_frame.grid_columnconfigure(0, weight=1)
        
        # Readings text area
        readings_text = scrolledtext.ScrolledText(readings_frame, width=50, height=3, wrap=tk.WORD)
        readings_text.grid(row=0, column=0, sticky="ew")
        readings_text.insert("1.0", readings)  # Populate with provided data
        self.add_mousewheel_scrolling(readings_text)
        
        # Buttons frame
        buttons_frame = ttk.Frame(readings_frame)
        buttons_frame.grid(row=0, column=1, sticky="ns", padx=(5, 0))
        
        def insert_p_marker():
            readings_text.insert(tk.INSERT, "[P] ")
        
        def count_words():
            text = readings_text.get("1.0", tk.END).strip()
            word_count = len(text.split())
            readings_text.insert(tk.END, f" [{word_count} words]")
        
        ttk.Button(buttons_frame, text="[P]", command=insert_p_marker, style="Small.TButton").pack(side=tk.TOP, pady=(0, 2))
        ttk.Button(buttons_frame, text="#", command=count_words, style="Small.TButton").pack(side=tk.TOP)
        
        # Work Due entry
        work_due_entry = ttk.Entry(self.entries_frame, width=20)
        work_due_entry.grid(row=row, column=3, sticky="w", padx=5, pady=2)
        work_due_entry.insert(0, work_due)  # Populate with provided data
        
        def remove_entry():
            date_entry.destroy()
            topic_entry.destroy()
            readings_frame.destroy()
            work_due_entry.destroy()
            delete_btn.destroy()
            self.schedule_entries.remove(entry_dict)
            self.repack_schedule_entries()
        
        # Delete button
        delete_btn = ttk.Button(self.entries_frame, text="X", command=remove_entry, style="Delete.TButton")
        delete_btn.grid(row=row, column=4, padx=(0, 5), pady=2)
        
        entry_dict = {
            "date": date_entry,
            "topic": topic_entry,
            "readings": readings_text,
            "work_due": work_due_entry,
            "delete_btn": delete_btn,
            "readings_frame": readings_frame
        }
        
        self.schedule_entries.append(entry_dict)
        
        # Bind update events
        for widget in [date_entry, topic_entry, work_due_entry]:
            widget.bind("<KeyRelease>", lambda e: self.update_document_preview() if hasattr(self, 'update_document_preview') else None)
        readings_text.bind("<KeyRelease>", lambda e: self.update_document_preview() if hasattr(self, 'update_document_preview') else None)
        
        return entry_dict

    def repack_schedule_entries(self):
        """Repack all schedule entries after a deletion"""
        for i, entry in enumerate(self.schedule_entries):
            entry["date"].grid(row=i, column=0, sticky="w", padx=(5, 10), pady=2)
            entry["topic"].grid(row=i, column=1, sticky="ew", padx=5, pady=2)
            entry["readings_frame"].grid(row=i, column=2, sticky="ew", padx=5, pady=2)
            entry["work_due"].grid(row=i, column=3, sticky="w", padx=5, pady=2)
            entry["delete_btn"].grid(row=i, column=4, padx=(0, 5), pady=2)

    def get_outcomes_range(self):
        """Get the range of outcomes based on current number of outcomes"""
        num_outcomes = len(self.outcome_entries) if hasattr(self, 'outcome_entries') else 0
        return f"Outcomes 1-{num_outcomes}" if num_outcomes > 0 else "No outcomes defined"

    def update_outcomes_references(self):
        """Update all references to outcomes in the learning objectives table"""
        outcomes_range = self.get_outcomes_range()
        if hasattr(self, 'learning_objectives_entries'):
            for entries in self.learning_objectives_entries.values():
                if 'assignments' in entries and entries['assignments'].winfo_exists():
                    current_state = entries['assignments'].cget('state')
                    
                    if current_state == 'disabled':
                        entries['assignments'].config(state='normal')
                    current_text = entries['assignments'].get("1.0", tk.END).strip()
                    if current_text.startswith("Outcome"):
                        entries['assignments'].delete("1.0", tk.END)
                        entries['assignments'].insert("1.0", outcomes_range)
                    if current_state == 'disabled':
                        entries['assignments'].config(state='disabled')

    def add_learning_objective_row(self, category="", slo="", assignments=""):
        """Add a new row to the Learning Objectives table"""
        row_frame = ttk.LabelFrame(self.lo_entries_frame, text=category or "New Category")
        row_frame.pack(fill=tk.X, pady=5, padx=5)
        
        # Default SLO text based on category
        default_slos = {
            "Content": "Identify, describe, and explain key themes, principles, and terminology; the history, theory and/or methodologies used; and social institutions, structures or processes.",
            "Critical Thinking": "Apply formal and informal qualitative or quantitative analysis effectively to examine the processes and means by which individuals make personal and group decisions. Assess and analyze ethical perspectives in individual and societal decisions.",
            "Communication": "Communication is the development and expression of ideas in written and oral forms."
        }
        
        # Category name
        name_frame = ttk.Frame(row_frame)
        name_frame.pack(fill=tk.X, pady=2)
        ttk.Label(name_frame, text="Category:").pack(side=tk.LEFT, padx=5)
        category_entry = ttk.Entry(name_frame, width=30)
        category_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        if category:
            category_entry.insert(0, category)
        
        # Social Science SLOs
        slo_frame = ttk.LabelFrame(row_frame, text="Social Science SLOs")
        slo_frame.pack(fill=tk.X, pady=2, padx=5)
        slo_text = scrolledtext.ScrolledText(slo_frame, width=40, height=4, wrap=tk.WORD)
        slo_text.pack(padx=5, pady=5)
        # Insert default SLO text based on category
        if category in default_slos:
            slo_text.insert("1.0", default_slos[category])
            slo_text.config(state="disabled")
        elif slo:
            slo_text.insert("1.0", slo)
        self.add_mousewheel_scrolling(slo_text)
        
        # State SLO Assignments
        assignments_frame = ttk.LabelFrame(row_frame, text="State SLO Assignments")
        assignments_frame.pack(fill=tk.X, pady=2, padx=5)
        assignments_text = scrolledtext.ScrolledText(assignments_frame, width=40, height=4, wrap=tk.WORD)
        assignments_text.pack(padx=5, pady=5)
        assignments_text.insert("1.0", assignments or self.get_outcomes_range())
        if category in default_slos:  # Only disable for predefined categories
            assignments_text.config(state="disabled")
        self.add_mousewheel_scrolling(assignments_text)
        
        # Course-Specific
        course_specific_frame = ttk.LabelFrame(row_frame, text="Course-Specific")
        course_specific_frame.pack(fill=tk.X, pady=2, padx=5)
        course_specific_text = scrolledtext.ScrolledText(course_specific_frame, width=40, height=4, wrap=tk.WORD)
        course_specific_text.pack(padx=5, pady=5)
        self.add_mousewheel_scrolling(course_specific_text)
        
        # Only show remove button for custom categories
        if category not in default_slos:
            ttk.Button(row_frame, text="Remove Category", 
                      command=lambda: self.remove_lo_category(row_frame, category),
                      style="Delete.TButton").pack(pady=5)
        
        # Store references
        if not hasattr(self, 'learning_objectives_entries'):
            self.learning_objectives_entries = {}
        self.learning_objectives_entries[category or f"category_{len(self.learning_objectives_entries)}"] = {
            "frame": row_frame,
            "name_entry": category_entry,
            "category": category_entry,
            "slo": slo_text,
            "assignments": assignments_text,
            "course_specific": course_specific_text
        }
        
        # Bind updates only for editable fields
        for widget in [category_entry, course_specific_text]:
            widget.bind("<KeyRelease>", lambda e: self.update_lo_preview() if hasattr(self, 'update_lo_preview') else None)
        
        return row_frame

    def remove_lo_category(self, frame, category):
        """Remove a learning objective category"""
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to remove this category?"):
            frame.destroy()
            if hasattr(self, 'learning_objectives_entries') and category in self.learning_objectives_entries:
                del self.learning_objectives_entries[category]
            if hasattr(self, 'update_lo_preview'):
                self.update_lo_preview()

    def update_lo_preview(self):
        """Update the Learning Objectives preview in the right panel"""
        if hasattr(self, 'preview_frame'):
            for widget in self.preview_frame.winfo_children():
                widget.destroy()
            
            # Title
            ttk.Label(self.preview_frame, 
                     text="Objectives—General Education and Social and Behavioral Sciences (S)", 
                     style="Heading.TLabel").pack(anchor="w", padx=5, pady=5)
            
            # Create table frame
            table_frame = ttk.Frame(self.preview_frame, relief=tk.SOLID, borderwidth=1)
            table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
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
            if hasattr(self, 'learning_objectives_entries'):
                for category, entries in self.learning_objectives_entries.items():
                    if 'frame' in entries and not entries['frame'].winfo_exists():
                        continue
                    
                    # Get text from entries:
                    category_text = entries['category'].get()
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
            
            # Show placeholder if no entries
            if row == 1:
                ttk.Label(table_frame, text="No learning objectives defined",
                         style="Italic.TLabel").grid(row=1, column=0, columnspan=4, pady=10)

    def add_objective_entry(self, default_text=""):
        """Add a new course objective entry with a number"""
        frame = ttk.Frame(self.objective_entries_frame)
        frame.pack(fill=tk.X, pady=2)
        
        number_label = ttk.Label(frame, text=f"{len(self.objective_entries)+1}.", width=3)
        number_label.pack(side=tk.LEFT, padx=(5,0))
        
        entry = ttk.Entry(frame, width=60)
        entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        if default_text:
            entry.insert(0, default_text)
        
        def remove_objective():
            frame.destroy()
            self.objective_entries.remove(obj_dict)
            self.renumber_objectives()
            if hasattr(self, 'update_document_preview'):
                self.update_document_preview()
        
        remove_btn = ttk.Button(frame, text="X", command=remove_objective, style="Delete.TButton")
        remove_btn.pack(side=tk.LEFT, padx=5)
        
        obj_dict = {"frame": frame, "number": number_label, "entry": entry}
        self.objective_entries.append(obj_dict)
        
        entry.bind("<KeyRelease>", lambda e: self.update_document_preview() if hasattr(self, 'update_document_preview') else None)
        
        return obj_dict

    def renumber_objectives(self):
        """Update numbering for course objectives"""
        if hasattr(self, 'objective_entries'):
            for i, obj in enumerate(self.objective_entries):
                obj["number"].config(text=f"{i+1}.")

    def add_outcome_entry(self, default_text=""):
        """Add a new Student Learning Outcome entry with a number"""
        # Create frame for the entry
        frame = ttk.Frame(self.outcome_entries_frame)
        frame.pack(fill=tk.X, pady=2)
        
        # Add number label for the entry
        number_label = ttk.Label(frame, text=f"{len(self.outcome_entries)+1}.", width=3)
        number_label.pack(side=tk.LEFT, padx=(5,0))
        
        # Add entry field
        entry = ttk.Entry(frame, width=60)
        entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        if default_text:
            entry.insert(0, default_text)
        
        def remove_outcome():
            frame.destroy()
            self.outcome_entries.remove(entry_dict)
            self.renumber_outcomes()
            self.update_outcomes_references()  # Update the references after removing
            if hasattr(self, 'update_lo_preview'):
                self.update_lo_preview()
        
        remove_btn = ttk.Button(frame, text="X", 
                              command=remove_outcome,
                              style="Delete.TButton")
        remove_btn.pack(side=tk.LEFT, padx=5)
        
        # Create dictionary to store entry information
        entry_dict = {
            "frame": frame,
            "number": number_label,
            "entry": entry
        }
        
        # Add to list of entries
        self.outcome_entries.append(entry_dict)
        
        # Bind update event
        entry.bind("<KeyRelease>", lambda e: self.update_all_previews() if hasattr(self, 'update_all_previews') else None)
        
        # Update references immediately
        self.update_outcomes_references()
        
        return entry_dict

    def update_all_previews(self):
        """Update both previews"""
        self.update_outcomes_references()
        if hasattr(self, 'update_lo_preview'):
            self.update_lo_preview()

    def renumber_outcomes(self):
        """Update the numbering of outcome entries after one is removed"""
        if hasattr(self, 'outcome_entries'):
            for i, entry in enumerate(self.outcome_entries):
                entry["number"].config(text=f"{i + 1}.")
        
        # Also update the preview
        if hasattr(self, 'update_lo_preview'):
            self.update_lo_preview()

    def create_course_info_tab(self):
        """Create the course information tab"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="★ Course Information")
        frame = self.create_scrollable_frame(tab)
        self.add_mousewheel_scrolling(tab, tab)

        # Course info fields
        fields = [
            ("Course Number:", "course_num", 15),
            ("Course Title:", "course_title", 40),
            ("Term:", "term", 20),
            ("Credits:", "credits", 5),
            ("Prerequisites:", "prerequisites", 40),
            ("Meeting Days/Times:", "meeting_times", 40),
            ("Location:", "location", 30)
        ]
        current_row = 0
        for label, field_name, width in fields:
            ttk.Label(frame, text=label).grid(row=current_row, column=0, sticky="e", padx=5, pady=5)
            entry = ttk.Entry(frame, width=width)
            entry.grid(row=current_row, column=1, sticky="w", padx=5, pady=5)
            setattr(self, f"entry_{field_name}", entry)
            current_row += 1

        # General Education (fixed to Social and Behavioral Sciences (S))
        ttk.Label(frame, text="General Education:").grid(row=current_row, column=0, sticky="e", padx=5, pady=5)
        gen_ed_frame = ttk.Frame(frame)
        gen_ed_frame.grid(row=current_row, column=1, sticky="w", padx=5, pady=5)
        ttk.Label(gen_ed_frame, text="Social and Behavioral Sciences (S)", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)
        # Add checkbox for showing Gen Ed section in exports/preview
        gen_ed_check = ttk.Checkbutton(
            gen_ed_frame,
            text="Show General Education Designation in syllabus",
            variable=self.show_gen_ed,
            command=self.update_document_preview
        )
        gen_ed_check.pack(side=tk.LEFT, padx=10)
        current_row += 1

        # Course description
        ttk.Label(frame, text="Course Description:").grid(row=current_row, column=0, sticky="ne", padx=5, pady=5)
        self.txt_description = scrolledtext.ScrolledText(frame, width=60, height=6, wrap=tk.WORD)
        self.txt_description.grid(row=current_row, column=1, sticky="w", padx=5, pady=5)
        self.add_mousewheel_scrolling(self.txt_description)
        current_row += 1

        # Course Objectives as numbered list input
        ttk.Label(frame, text="Course Objectives:").grid(row=current_row, column=0, sticky="nw", padx=5, pady=5)
        self.objective_entries_frame = ttk.Frame(frame)
        self.objective_entries_frame.grid(row=current_row, column=1, sticky="w", padx=5, pady=5)
        self.objective_entries = []  # Start with an empty list so teachers can enter objectives
        add_objective_btn = ttk.Button(frame, text="Add Objective", command=self.add_objective_entry, style="Action.TButton")
        add_objective_btn.grid(row=current_row+1, column=1, sticky="w", padx=5, pady=5)
        current_row += 2

    def create_learning_objectives_tab(self):
        """Create a tab for customizable learning objectives table"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Learning Objectives & SLOs")
        self.add_mousewheel_scrolling(tab, tab)

        # Create vertical paned window to split content
        main_paned = ttk.PanedWindow(tab, orient=tk.VERTICAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Top section - Student Learning Outcomes
        top_frame = ttk.LabelFrame(main_paned, text="Student Learning Outcomes")
        main_paned.add(top_frame, weight=1)
        
        ttk.Label(top_frame, text="A student who successfully completes this course will:", 
                 style="Italic.TLabel").pack(anchor="w", padx=10, pady=(5,10))
        
        # Container for SLO entries
        self.outcome_entries_frame = ttk.Frame(top_frame)
        self.outcome_entries_frame.pack(fill=tk.X, expand=True, padx=10)
        self.outcome_entries = []
        # Add default SLO entries with empty content
        default_outcomes = [
            "",  # Empty strings for blank entries
            "",
            "",
            ""
        ]
        # Create initial 4 empty outcome entries
        for outcome in default_outcomes:
            self.add_outcome_entry(outcome)
        
        # Add button for new outcomes
        ttk.Button(top_frame, text="Add Learning Outcome", 
                   command=self.add_outcome_entry,
                   style="Action.TButton").pack(pady=10, padx=10)
        
        # Bottom section - Learning Objectives Table
        bottom_paned = ttk.PanedWindow(main_paned, orient=tk.HORIZONTAL)
        main_paned.add(bottom_paned, weight=2)
        
        # Left side - editing
        edit_frame = ttk.LabelFrame(bottom_paned, text="Learning Objectives Table")
        bottom_paned.add(edit_frame, weight=1)
        
        ttk.Label(edit_frame, text="Fill in the course-specific information for each category:", 
                 style="Italic.TLabel").pack(anchor="w", padx=10, pady=(5,10))
        
        # Scrollable container for learning objectives
        canvas = tk.Canvas(edit_frame)
        scrollbar = ttk.Scrollbar(edit_frame, orient="vertical", command=canvas.yview)
        
        self.lo_entries_frame = ttk.Frame(canvas)
        self.lo_entries_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.lo_entries_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack canvas and scrollbar
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Add default categories with blank content
        default_categories = {
            "Content": {
                "slo": "",
                "assignments": self.get_outcomes_range(),  # Dynamic based on number of outcomes
                "course_specific": ""
            },
            "Critical Thinking": {
                "slo": "",
                "assignments": self.get_outcomes_range(),
                "course_specific": ""
            },
            "Communication": {
                "slo": "",
                "assignments": self.get_outcomes_range(),
                "course_specific": ""
            }
        }
        for category, content in default_categories.items():
            self.add_learning_objective_row(category, content["slo"], content["assignments"])
        
        ttk.Button(edit_frame, text="Add New Category", 
                   command=lambda: self.add_learning_objective_row(),
                   style="Action.TButton").pack(pady=10)
        
        # Right side - preview
        preview_frame = ttk.LabelFrame(bottom_paned, text="Preview")
        bottom_paned.add(preview_frame, weight=1)
        
        self.preview_frame = ttk.Frame(preview_frame)
        self.preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Initial preview
        self.update_lo_preview()

    def create_instructor_info_tab(self):
        """Create the instructor information tab with proper layout"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="★ Instructor Information")
        self.add_mousewheel_scrolling(tab, tab)

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
        
        # Sections
        self.ta_frame = ttk.LabelFrame(content_frame, text="Sections")
        self.ta_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.ta_container = ttk.Frame(self.ta_frame)
        self.ta_container.pack(fill=tk.X, padx=5, pady=5)
        self.ta_entries = []
        
        ttk.Button(self.ta_frame, text="Add Section", 
                  command=self.add_ta,
                  style="Action.TButton").pack(pady=5)
        
    def create_schedule_tab(self):
        """Create the course schedule tab with proper layout"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="★ Calendar/Course Schedule") 
        self.add_mousewheel_scrolling(tab, tab)

        # Main container
        main_frame = ttk.Frame(tab, padding="5")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Top controls frame
        controls_frame = ttk.Frame(main_frame)
        controls_frame.pack(fill=tk.X, pady=(0, 5))
        
        # Left side buttons
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
        
        # Add to the buttons_frame in create_schedule_tab method:
        # Update the button text in create_schedule_tab method:
        example_csv_btn = ttk.Button(
            buttons_frame, text="Save CSV Template Example", 
            command=self.export_schedule_example,
            style="Action.TButton"
        )
        example_csv_btn.pack(side=tk.LEFT, padx=5)
        
        # Right side helper text
        helper_frame = ttk.Frame(controls_frame)
        helper_frame.pack(side=tk.RIGHT)
        
        ttk.Label(helper_frame, text="Use [P] to mark primary sources", 
                  style="Italic.TLabel").pack(side=tk.RIGHT, padx=(0, 10))
        ttk.Label(helper_frame, text="⚠️ Include page/word/time counts for all readings/films", 
                  style="Italic.TLabel", foreground="red").pack(side=tk.RIGHT, padx=(0, 10))
        
        # Fixed header frame for column headings
        header_frame = ttk.Frame(main_frame, style="Preview.TFrame")
        header_frame.pack(fill=tk.X, pady=(0, 5))  # Use pack instead of grid
        
        # Define column headings
        ttk.Label(header_frame, text="Date", style="Heading.TLabel").pack(side=tk.LEFT, padx=(5, 10))
        ttk.Label(header_frame, text="Topic", style="Heading.TLabel").pack(side=tk.LEFT, padx=150)
        ttk.Label(header_frame, text="Readings/Preparation", style="Heading.TLabel").pack(side=tk.LEFT, padx=5)
        ttk.Label(header_frame, text="Work Due", style="Heading.TLabel").pack(side=tk.LEFT, padx=175)
        
        # Scrollable frame for schedule entries
        canvas = tk.Canvas(main_frame, bg='#f5f5f5')  # Light gray background
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        
        # Create a frame inside the canvas for entries
        self.entries_frame = ttk.Frame(canvas, style="Template.TFrame")
        self.entries_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        # Configure the canvas
        canvas.create_window((0, 0), window=self.entries_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Pack the canvas and scrollbar
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=5)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)

        # Enable mousewheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        self.schedule_entries = []
        
    def create_assignments_tab(self):
        """Create the assignments and grading tab with proper layout"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="★ Assignments & Grading [Required]")
        self.add_mousewheel_scrolling(tab, tab)

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

        # --- Formatting Toolbar ---
        toolbar = ttk.Frame(materials_frame)
        toolbar.pack(fill=tk.X, padx=5, pady=(5, 0))

        def insert_markup(markup_type):
            text_widget = self.materials_text
            try:
                sel_start = text_widget.index(tk.SEL_FIRST)
                sel_end = text_widget.index(tk.SEL_LAST)
                selected = text_widget.get(sel_start, sel_end)
            except tk.TclError:
                sel_start = sel_end = None
                selected = ""
            if markup_type == "bold":
                tag = "**"
                new_text = f"{tag}{selected or 'bold text'}{tag}"
            elif markup_type == "italic":
                tag = "*"
                new_text = f"{tag}{selected or 'italic text'}{tag}"
            elif markup_type == "link":
                new_text = f"[{selected or 'link text'}](http://example.com)"
            else:
                return
            if sel_start and sel_end:
                text_widget.delete(sel_start, sel_end)
                text_widget.insert(sel_start, new_text)
            else:
                text_widget.insert(tk.INSERT, new_text)

        bold_btn = ttk.Button(toolbar, text="B", width=2, command=lambda: insert_markup("bold"))
        bold_btn.pack(side=tk.LEFT, padx=(0, 2))
        italic_btn = ttk.Button(toolbar, text="I", width=2, command=lambda: insert_markup("italic"))
        italic_btn.pack(side=tk.LEFT, padx=(0, 2))
        link_btn = ttk.Button(toolbar, text="🔗", width=2, command=lambda: insert_markup("link"))
        link_btn.pack(side=tk.LEFT, padx=(0, 2))

        # Materials text area
        self.materials_text = scrolledtext.ScrolledText(materials_frame, width=60, height=6, wrap=tk.WORD)
        self.materials_text.pack(fill=tk.X, padx=5, pady=5)
        self.materials_text.insert("1.0", "**Required** textbook *and* materials will be *listed* here.")
        self.add_mousewheel_scrolling(self.materials_text)
        
        fee_frame = ttk.Frame(materials_frame)
        fee_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(fee_frame, text="Materials Fee: $").pack(side=tk.LEFT)
        self.fee_entry = ttk.Entry(fee_frame, width=10)
        self.fee_entry.pack(side=tk.LEFT)
        
        # Add this to the relevant section of your UI (e.g., in the Required Materials section)
        help_button = ttk.Button(materials_frame, text="Formatting Help", command=self.show_formatting_help)
        help_button.pack(side=tk.RIGHT, padx=5, pady=5)

        # Grading Components Section
        components_frame = ttk.LabelFrame(content_frame, text="Graded Components")
        components_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.categories_frame = ttk.Frame(components_frame)
        self.categories_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.category_frames = []
        
        ttk.Button(components_frame, text="Add Category", 
                  command=self.add_category,
                  style="Action.TButton").pack(padx=5, pady=5)
        
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
        self.add_mousewheel_scrolling(tab, tab)
        
        frame = self.create_scrollable_frame(tab)
        
        # Add information header
        info_frame = ttk.LabelFrame(frame, text="Policy Information")
        info_frame.pack(fill=tk.X, padx=20, pady=10)
        
        info_label = ttk.Label(
            info_frame,
            text="All necessary UF academic policies, including University Assessment Policies, \nEvaluations, Recording Policy, Campus Resources, and Academic Resources \nwill be automatically included via the UF Policy Link below.",
            font=('Arial', 10),
            foreground='blue'
        )
        info_label.pack(padx=10, pady=10)
        
        # UF Policies Section - Always enabled by default
        uf_policies_frame = ttk.LabelFrame(frame, text="UF Academic Policies")
        uf_policies_frame.pack(fill=tk.X, padx=20, pady=5)
        
        # Set this to always be enabled
        self.use_simplified_policies_var.set(True)
        
        uf_policy_label = ttk.Label(
            uf_policies_frame,
            text="✓ UF Academic Policies (automatically included)",
            font=('Arial', 10, 'bold'),
            foreground='green'
        )
        uf_policy_label.pack(anchor="w", padx=10, pady=5)
        
        ttk.Label(uf_policies_frame, 
                 text="This includes all required UF policies via link: https://syllabus.ufl.edu/syllabus-policy/uf-syllabus-policy-links/", 
                 style="Italic.TLabel").pack(anchor="w", padx=20, pady=(0, 10))

        # Canvas Policy checkbox (move above the Canvas Policy entry box)
        canvas_check = ttk.Checkbutton(
            frame,
            text="Include Canvas Policy",
            variable=self.canvas_policy_var,
            command=self.update_document_preview
        )
        canvas_check.pack(anchor="w", padx=5, pady=2)

        # Canvas Policy entry box
        canvas_frame = ttk.LabelFrame(frame, text="Canvas Policy")
        canvas_frame.pack(fill=tk.X, padx=20, pady=(10, 5))
        
        self.canvas_policy_text = scrolledtext.ScrolledText(canvas_frame, width=60, height=4, wrap=tk.WORD)
        self.canvas_policy_text.pack(fill=tk.X, padx=5, pady=5)
        self.canvas_policy_text.insert("1.0", canvas_policy_default)  # Changed from placeholder
        self.add_mousewheel_scrolling(self.canvas_policy_text)
        
        ttk.Label(canvas_frame, text="This will appear under the Extra Credit policy in the syllabus.", 
                 style="Italic.TLabel").pack(anchor="w", padx=5, pady=(0, 5))
        
        # Technology Policy checkbox (move above the Technology Policy entry box)
        technology_check = ttk.Checkbutton(
            frame,
            text="Include Technology Policy",
            variable=self.technology_policy_var,
            command=self.update_document_preview
        )
        technology_check.pack(anchor="w", padx=5, pady=2)

        # Technology Policy entry box
        tech_frame = ttk.LabelFrame(frame, text="Technology in the Classroom Policy")
        tech_frame.pack(fill=tk.X, padx=20, pady=(10, 5))
        
        self.technology_policy_text = scrolledtext.ScrolledText(tech_frame, width=60, height=4, wrap=tk.WORD)
        self.technology_policy_text.pack(fill=tk.X, padx=5, pady=5)
        self.technology_policy_text.insert("1.0", technology_policy_default)  # Changed from placeholder
        self.add_mousewheel_scrolling(self.technology_policy_text)
        
        ttk.Label(tech_frame, text="This will appear after the Canvas policy in the syllabus.", 
                 style="Italic.TLabel").pack(anchor="w", padx=5, pady=(0, 5))
        
        # Class Communication Policy checkbox (move above the Class Communication Policy entry box)
        communication_check = ttk.Checkbutton(
            frame,
            text="Include Class Communication Policy",
            variable=self.communication_policy_var,
            command=self.update_document_preview
        )
        communication_check.pack(anchor="w", padx=5, pady=2)

        # Class Communication Policy entry box
        comm_frame = ttk.LabelFrame(frame, text="Class Communication Policy")
        comm_frame.pack(fill=tk.X, padx=20, pady=(10, 5))
        
        self.communication_policy_text = scrolledtext.ScrolledText(comm_frame, width=60, height=4, wrap=tk.WORD)
        self.communication_policy_text.pack(fill=tk.X, padx=5, pady=5)
        self.communication_policy_text.insert("1.0", class_communication_policy_default)  # Changed from placeholder
        self.add_mousewheel_scrolling(self.communication_policy_text)
        
        ttk.Label(comm_frame, text="This will appear after the Technology policy in the syllabus.", 
                 style="Italic.TLabel").pack(anchor="w", padx=5, pady=(0, 5))
        
        # Optional policies (moved below comm)
        support_check = ttk.Checkbutton(
            frame,
            text="Include Assignment Support Outside the Classroom",
            variable=self.outside_support_var,
            command=self.update_document_preview
        )
        support_check.pack(anchor="w", padx=5, pady=2)

        # Then add the support_frame
        support_frame = ttk.LabelFrame(frame, text="Assignment Support Details")
        support_frame.pack(fill=tk.X, padx=20, pady=(10, 5))
        self.support_text = scrolledtext.ScrolledText(support_frame, width=60, height=4, wrap=tk.WORD)
        self.support_text.pack(fill=tk.X, padx=5, pady=5)
        self.support_text.insert("1.0", assignment_support_default)
        self.add_mousewheel_scrolling(self.support_text)

        # Add separator for new section
        ttk.Separator(frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        # Grading Options Section
        grading_options_frame = ttk.LabelFrame(frame, text="Grading Options")
        grading_options_frame.pack(fill=tk.X, padx=20, pady=5)
        
        # Grading rounding option
        grading_rounding_check = ttk.Checkbutton(
            grading_options_frame,
            text="Include grading rounding statement",
            variable=self.grading_rounding_var,
            command=self.update_document_preview
        )
        grading_rounding_check.pack(anchor="w", padx=5, pady=2)
        
        ttk.Label(grading_options_frame, 
                 text="\"All non-whole number grades .5 and above will be rounded up (example: 89.5 → 90)\"", 
                 style="Italic.TLabel").pack(anchor="w", padx=20, pady=(0, 5))
        
        # Add separator for new section
        ttk.Separator(frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)

        # --- Move Late Submissions Policy Section here ---
        # Late Submissions Policy checkbox (above the late_frame)
        late_policy_check = ttk.Checkbutton(
            frame,
            text="Include Late Submissions Policy",
            variable=self.late_submissions_policy_var,
            command=self.update_document_preview
        )
        late_policy_check.pack(anchor="w", padx=5, pady=2)

        late_frame = ttk.LabelFrame(frame, text="Late Submissions Policy")
        late_frame.pack(fill=tk.X, padx=20, pady=5)
        ttk.Label(late_frame, text="Late Submissions Policy:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.late_policy_var = tk.StringVar()
        # Default late policy text that should be used for Custom option
        default_late_text = "Late assignments will be penalized 10% per day late unless prior arrangements have been made with the instructor due to documented emergency or illness. Contact instructor as soon as possible if you anticipate being unable to meet a deadline."
        self.late_policies = {
            "Standard (10% per day)": "Unless an extension is granted, assignments will incur a 10-point penalty for every day they are late.",
            "No late work": "No late work will be accepted without prior approval.",
            "48-hour grace": "Students have a 48-hour grace period for submissions, after which no late work will be accepted.",
            "Custom": default_late_text  # Use the default text instead of empty string
        }
        self.late_combo = ttk.Combobox(late_frame, 
                                      textvariable=self.late_policy_var,
                                      values=list(self.late_policies.keys()), 
                                      width=30)
        self.late_combo.pack(side=tk.LEFT)
        self.late_combo.set("Custom")  # Set to Custom to match template default
        
        self.late_policy_text = scrolledtext.ScrolledText(late_frame, width=60, height=3, wrap=tk.WORD)
        self.late_policy_text.pack(fill=tk.X, padx=5, pady=(0, 5))
        # Initialize with the default policy text (same as Custom option)
        self.late_policy_text.insert('1.0', default_late_text)
        
        def update_late_policy(*args):
            selected = self.late_policy_var.get()
            if selected in self.late_policies:
                # Enable text widget
                self.late_policy_text.config(state='normal')
                self.late_policy_text.delete('1.0', tk.END)
                self.late_policy_text.insert('1.0', self.late_policies[selected])
                if selected != "Custom":
                    self.late_policy_text.config(state='disabled')
                # Update optional policy state just by setting it:
                if hasattr(self, 'optional_policies'):
                    # Check if the key exists before setting, although it should now
                    if "late_work" in self.optional_policies:
                        if selected == "No late work":
                            self.optional_policies["late_work"].set(False)
                        else:
                            self.optional_policies["late_work"].set(True)
        
        self.late_policy_var.trace_add('write', update_late_policy)
        self.late_combo.bind('<<ComboboxSelected>>', lambda e: update_late_policy())

        # --- Move Extra Credit Policy Section here ---
        # Extra Credit Policy checkbox (above the extra_frame)
        extra_credit_check = ttk.Checkbutton(
            frame,
            text="Include Extra Credit Policy",
            variable=self.extra_credit_policy_var,
            command=self.update_document_preview
        )
        extra_credit_check.pack(anchor="w", padx=5, pady=2)

        extra_frame = ttk.LabelFrame(frame, text="Extra Credit Policy")
        extra_frame.pack(fill=tk.X, padx=20, pady=5)
        ttk.Label(extra_frame, text="Extra Credit Policy:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.extra_credit_var = tk.StringVar()
        # Default extra credit text that should be used for Custom option
        default_extra_credit_text = "Extra credit opportunities may be available at the instructor's discretion. These will be announced in class and posted on Canvas when available."
        self.extra_credit_policies = {
            "Standard": "Extra credit opportunities may be announced during the semester. Points will be added to your mid-term exam grade.",
            "No extra credit": "No extra credit will be offered in this course.",
            "Optional assignments": "Students may complete optional assignments for extra credit, worth up to 3% of the final grade.",
            "Custom": default_extra_credit_text  # Use the default text instead of empty string
        }
        self.extra_combo = ttk.Combobox(extra_frame, textvariable=self.extra_credit_var,
                                        values=list(self.extra_credit_policies.keys()), width=30)
        self.extra_combo.pack(side=tk.LEFT)
        self.extra_combo.set("Custom")  # Set to Custom to match template default
        
        self.extra_credit_text = scrolledtext.ScrolledText(extra_frame, width=60, height=3, wrap=tk.WORD)
        self.extra_credit_text.pack(fill=tk.X, padx=5, pady=(0, 5))
        # Initialize with the default extra credit text (same as Custom option)
        self.extra_credit_text.insert('1.0', default_extra_credit_text)
        
        def update_extra_credit(*args):
            selected = self.extra_credit_var.get()
            if selected in self.extra_credit_policies:
                self.extra_credit_text.config(state='normal')
                self.extra_credit_text.delete('1.0', tk.END)
                self.extra_credit_text.insert('1.0', self.extra_credit_policies[selected])
                if selected != "Custom":
                    self.extra_credit_text.config(state='disabled')
        
        self.extra_credit_var.trace_add('write', update_extra_credit)
        self.extra_combo.bind('<<ComboboxSelected>>', lambda e: update_extra_credit())

    def create_action_buttons(self):
        """Create action buttons at the bottom of the interface"""
        # Create a frame at the bottom that will stay fixed
        # Use self.root (window) directly instead of main_container
        self.action_frame = tk.Frame(self.root, bg="#f0f0f0")  
        self.action_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=10)
        
        # Add a separator above the action frame to visually separate it
        separator = ttk.Separator(self.root, orient='horizontal')
        separator.pack(side=tk.BOTTOM, fill=tk.X, before=self.action_frame)

        # Generate syllabus button (Word)
        gen_syllabus_btn = ttk.Button(
            self.action_frame, text="Generate Syllabus (Word)", style="Action.TButton",
            command=lambda: self.generate_syllabus(export_format="docx")
        )
        gen_syllabus_btn.pack(side=tk.RIGHT, padx=5)
        
        # Optional: Add save/load project buttons on the left side
        project_frame = ttk.Frame(self.action_frame)
        project_frame.pack(side=tk.LEFT, padx=5)
        
        save_project_btn = ttk.Button(
            project_frame, text="Save Project", style="Action.TButton",
            command=self.save_template  # Using existing save_template method
        )
        save_project_btn.pack(side=tk.LEFT, padx=5)

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
            if "readings_frame" in entry:
                entry["readings_frame"].destroy()
            if "delete_btn" in entry:
                entry["delete_btn"].destroy()
            entry["date"].destroy()
            entry["topic"].destroy()
            entry["work_due"].destroy()
        self.schedule_entries.clear()
        
        # Import based on file type
        if file_path.endswith('.csv'):
            import csv
            with open(file_path, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    self.add_schedule_entry(
                        row.get("Date", ""),
                        row.get("Topic", ""),
                        row.get("Readings/Preparation", ""),
                        row.get("Work Due", "")
                    )
        else:
            # Handle Excel import
            try:
                import pandas as pd
                df = pd.read_excel(file_path)
                for _, row in df.iterrows():
                    self.add_schedule_entry(
                        str(row.get("Date", "")),
                        str(row.get("Topic", "")),
                        str(row.get("Readings/Preparation", "")),
                        str(row.get("Work Due", ""))
                    )
            except ImportError:
                messagebox.showerror("Error", "pandas is required for Excel import. Please install it or use CSV format.")
        
        messagebox.showinfo("Success", f"Schedule imported from {file_path}")
            
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
            try:
                import pandas as pd
                df = pd.DataFrame(data)
                df.to_excel(file_path, index=False)
            except ImportError:
                messagebox.showerror("Error", "pandas is required for Excel export. Please install it or use CSV format.")
                return
                
        messagebox.showinfo("Success", f"Schedule exported to {file_path}")
            
    def export_schedule_example(self):
        """Create an example CSV schedule file to show proper format"""
        # Define example data with clear formatting examples
        example_data = [
            ["Date", "Topic", "Readings/Preparation", "Work Due"],
            ["January 10, 2025", "Introduction to Course", "Syllabus [1000 words]", "None"],
            ["January 17, 2025", "The Progressive Era", "American Yawp, Ch. 20 [8400 words]\nTheodore Roosevelt, 'The New Nationalism' [P]", "Reading Response #1"],
            ["January 24, 2025", "World War I", "American Yawp, Ch. 21 [8750 words]\nWoodrow Wilson, 'War Message to Congress' [P]", "Paper Proposal Due"],
            ["January 31, 2025", "The 1920s", "American Yawp, Ch. 22 [7800 words]\nF. Scott Fitzgerald, excerpt from The Great Gatsby [P]", "Discussion Post"],
            ["February 7, 2025", "Great Depression", "American Yawp, Ch. 23 [8200 words]\nFDR, First Inaugural Address [P]", "Reading Response #2"]
        ]
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Save Example Schedule CSV"
        )
        
        if not file_path:
            return
        
        try:
            import csv
            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerows(example_data)
            
            messagebox.showinfo("Success", 
                "Example schedule CSV template has been saved to:\n\n"
                f"{file_path}\n\n"
                "Please use this file as a template/guideline when creating your own schedule.\n"
                "You can edit it directly in Excel or any spreadsheet program, then import it back "
                "using the 'Import Schedule' button.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save example: {str(e)}")
