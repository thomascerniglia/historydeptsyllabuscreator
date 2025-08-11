"""
History Syllabus Generator - Main Application File
This is the main entry point for the History Syllabus Generator application.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from tkinter import scrolledtext
import os
import sys

# Import the modules we've created
from constants import *
from templates import SyllabusTemplate, load_default_templates
from ui_tabs import UITabsMixin
from document_generation import DocumentGenerationMixin
from document_preview import DocumentPreviewMixin

class HistorySyllabusGenerator(UITabsMixin, DocumentGenerationMixin, DocumentPreviewMixin):
    """Main application class for the History Syllabus Generator"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("History Syllabus Generator")
        self.root.state('zoomed')
        self.current_template = None
        self.templates = []
        self.template_names = []
        # Store Sections (keeping ta_entries variable name for compatibility)
        self.ta_entries = []
        # Store schedule entries
        self.schedule_entries = []
        # Store category frames
        self.category_frames = []
        # Store learning objective entries
        self.learning_objectives_entries = {}
        # Set up styles
        self.setup_styles()
        # Add variable for Gen Ed toggle
        self.show_gen_ed = tk.BooleanVar(value=True)
        self.outside_support_var = tk.BooleanVar(value=True)
        # Keep only instructor-specific policy variables (UF policies will be automatic)
        self.late_submissions_policy_var = tk.BooleanVar(value=True)
        self.extra_credit_policy_var = tk.BooleanVar(value=True)
        self.canvas_policy_var = tk.BooleanVar(value=True)
        self.technology_policy_var = tk.BooleanVar(value=True)
        self.communication_policy_var = tk.BooleanVar(value=True)
        # Add grading rounding option
        self.grading_rounding_var = tk.BooleanVar(value=False)
        # UF policies are now always included via link (not optional)
        self.use_simplified_policies_var = tk.BooleanVar(value=True)
        self.optional_policies = {
            "late_submissions": self.late_submissions_policy_var,
            "extra_credit": self.extra_credit_policy_var,
            "canvas": self.canvas_policy_var,
            "technology": self.technology_policy_var,
            "communication": self.communication_policy_var
        }
        
        # Initialize late policies dictionary
        self.late_policies = {
            "Standard (10% per day)": "Late assignments will be penalized 10% per day late unless prior arrangements have been made with the instructor.",
            "Strict (no late work)": "Late assignments will not be accepted unless prior arrangements have been made with the instructor due to documented emergency or illness.",
            "Flexible (reduced points)": "Late assignments will be accepted with reduced points. Contact instructor for specific penalties.",
            "Custom": ""
        }
        
        # Initialize extra credit policies dictionary
        self.extra_credit_policies = {
            "Standard": "Extra credit opportunities may be available at the instructor's discretion. These will be announced in class and posted on Canvas.",
            "None available": "No extra credit opportunities will be offered in this course.",
            "Project-based": "Extra credit may be earned through additional research projects or presentations. Contact instructor for details.",
            "Custom": ""
        }
        
        # Create action frame FIRST before main_container
        # This ensures it's always at the bottom regardless of content
        self.create_action_buttons()
        
        # Then create the main interface
        self.create_main_interface()
        
        self.templates = load_default_templates()
        
        self.template_names = [f"{t.course_code}: {t.title}" for t in self.templates]
        self.template_combo['values'] = ["Clear Template"] + self.template_names

    def setup_styles(self):
        """Set up ttk styles for the application"""
        style = ttk.Style()
        
        # Configure tab styles
        style.configure('TNotebook.Tab', padding=[20, 8], font=('Arial', 10, 'bold'))
        style.configure('TNotebook', tabposition='n')
        
        # Configure button styles
        style.configure('Action.TButton', font=('Arial', 11, 'bold'), padding=[10, 5])
        
        # Configure label styles
        style.configure('Heading.TLabel', font=('Arial', 12, 'bold'))
        style.configure('Bold.TLabel', font=('Arial', 10, 'bold'))

    def create_action_buttons(self):
        """Create the action buttons at the bottom of the window"""
        # Create bottom action frame
        self.action_frame = tk.Frame(self.root, bg='lightgray', height=60)
        self.action_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=5)
        self.action_frame.pack_propagate(False)
        
        # Generate buttons frame (centered)
        generate_frame = tk.Frame(self.action_frame, bg='lightgray')
        generate_frame.pack(expand=True, pady=10)
        
        ttk.Button(generate_frame, text="Generate Word Document", 
                  command=lambda: self.generate_syllabus("docx"), 
                  style='Action.TButton').pack(side=tk.LEFT, padx=5)
        
        ttk.Button(generate_frame, text="Generate PDF Document", 
                  command=lambda: self.generate_syllabus("pdf"), 
                  style='Action.TButton').pack(side=tk.LEFT, padx=5)

    def create_main_interface(self):
        """Create the main tabbed interface"""
        # Create main container for content (above action buttons)
        self.main_container = tk.Frame(self.root)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 0))
        
        # Add template selection at the top
        template_frame = tk.Frame(self.main_container)
        template_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(template_frame, text="Load Template:", font=('Arial', 10, 'bold')).pack(side=tk.LEFT)
        
        self.template_combo = ttk.Combobox(template_frame, values=["Clear Template"], state="readonly", width=25)
        self.template_combo.pack(side=tk.LEFT, padx=(10, 0))
        self.template_combo.bind("<<ComboboxSelected>>", self.on_template_selected)
        self.template_combo.set("Clear Template")
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Create all tabs using methods from UITabsMixin
        self.create_course_info_tab()
        self.create_instructor_info_tab()
        self.create_schedule_tab()
        self.create_assignments_tab()
        self.create_policies_tab()
        self.create_document_preview_tab()

    def on_template_selected(self, event=None):
        """Handle template selection from dropdown"""
        selected = self.template_combo.get()
        
        if selected == "Clear Template":
            self.clear_all_fields()
        else:
            # Find the template
            template_index = self.template_names.index(selected)
            template = self.templates[template_index]
            self.load_template_content(template)

    def load_template_content(self, template):
        """Load template content into form fields"""
        try:
            # Clear existing content first using the comprehensive clearing method
            self.clear_all_fields()
            
            # Course Info
            if hasattr(template, 'course_code'):
                self.entry_course_num.insert(0, template.course_code)
            if hasattr(template, 'title'):
                self.entry_course_title.insert(0, template.title)
            if hasattr(template, 'prerequisites'):
                self.entry_prerequisites.insert(0, template.prerequisites)
            # Additional course info fields
            if hasattr(template, 'semester'):
                self.entry_term.insert(0, template.semester)
            if hasattr(template, 'credits'):
                self.entry_credits.insert(0, template.credits)
            if hasattr(template, 'class_days') and hasattr(template, 'class_times'):
                meeting_times = f"{template.class_days} {template.class_times}"
                self.entry_meeting_times.insert(0, meeting_times)
            if hasattr(template, 'classroom'):
                self.entry_location.insert(0, template.classroom)

            # Instructor Info
            if hasattr(template, 'instructor_name'):
                self.entry_instr_name.insert(0, template.instructor_name)
            if hasattr(template, 'instructor_office'):
                self.entry_instr_office.insert(0, template.instructor_office)
            if hasattr(template, 'instructor_phone'):
                self.entry_instr_phone.insert(0, template.instructor_phone)
            if hasattr(template, 'instructor_email'):
                self.entry_instr_email.insert(0, template.instructor_email)
            if hasattr(template, 'instructor_office_hours'):
                self.entry_instr_office_hours.insert(0, template.instructor_office_hours)
                
            # Sections
            if hasattr(template, 'tas') and template.tas:
                # Add each Section from the template
                for ta_data in template.tas:
                    # Create new Section entry
                    self.add_ta()
                    # Get the latest Section entry (the one we just added)
                    if self.ta_entries:
                        latest_ta = self.ta_entries[-1]
                        latest_ta[0].insert(0, ta_data.get('name', ''))
                        latest_ta[1].insert(0, ta_data.get('email', ''))
                        latest_ta[2].insert(0, ta_data.get('office_hours', ''))
                        latest_ta[3].insert(0, ta_data.get('class_room', ''))
                        latest_ta[4].insert(0, ta_data.get('class_time', ''))
                    
            # Course Description & Objectives
            if hasattr(template, 'description'):
                self.txt_description.insert("1.0", template.description)
            if hasattr(template, 'objectives'):
                for obj in template.objectives:
                    self.add_objective_entry(obj)
                    
            # Student Learning Outcomes
            if hasattr(template, 'outcomes'):
                # Only load outcomes if the UI frame exists (tab has been created)
                if hasattr(self, 'outcome_entries_frame'):
                    for outcome in template.outcomes:
                        self.add_outcome_entry(outcome)
                
            # Schedule
            if hasattr(template, 'schedule') and template.schedule:
                for entry in template.schedule:
                    self.add_schedule_entry(
                        entry.get('date', ''),
                        entry.get('topic', ''),
                        entry.get('readings', ''),
                        entry.get('work_due', '')
                    )
                    
            # Learning Objectives Table
            if hasattr(template, 'learning_objectives') and template.learning_objectives:
                for category, data in template.learning_objectives.items():
                    self.add_learning_objective_row(
                        category,
                        data.get('slo', ''),
                        data.get('assignments', '')
                    )
                    
            # Load policy text content - clear first, then load template content
            if hasattr(template, 'canvas_policy') and hasattr(self, 'canvas_policy_text'):
                self.canvas_policy_text.delete("1.0", tk.END)
                self.canvas_policy_text.insert("1.0", template.canvas_policy)
            if hasattr(template, 'technology_policy') and hasattr(self, 'technology_policy_text'):
                self.technology_policy_text.delete("1.0", tk.END)
                self.technology_policy_text.insert("1.0", template.technology_policy)
            if hasattr(template, 'communication_policy') and hasattr(self, 'communication_policy_text'):
                self.communication_policy_text.delete("1.0", tk.END)
                self.communication_policy_text.insert("1.0", template.communication_policy)
            if hasattr(template, 'support_policy') and hasattr(self, 'support_text'):
                self.support_text.delete("1.0", tk.END)
                self.support_text.insert("1.0", template.support_policy)
            
            # Policy dropdowns and boolean variables (set first to avoid trace conflicts)
            self._set_policy_dropdowns(template)
            
            # IMPORTANT: Only load custom policy text if it exists in template
            # If no custom text, preserve the existing default text in the UI
            if hasattr(template, 'late_policy_text') and hasattr(self, 'late_policy_text'):
                def set_late_policy_text():
                    self.late_policy_text.config(state='normal')
                    self.late_policy_text.delete("1.0", tk.END)
                    self.late_policy_text.insert("1.0", template.late_policy_text)
                    print(f"DEBUG: Loaded late policy text: {template.late_policy_text[:50]}...")
                self.root.after(10, set_late_policy_text)  # Delay 10ms
            else:
                print("DEBUG: No custom late policy text in template, preserving default UI text")
            
            if hasattr(template, 'extra_credit_policy_text') and hasattr(self, 'extra_credit_text'):
                def set_extra_credit_text():
                    self.extra_credit_text.config(state='normal') 
                    self.extra_credit_text.delete("1.0", tk.END)
                    self.extra_credit_text.insert("1.0", template.extra_credit_policy_text)
                    print(f"DEBUG: Loaded extra credit text: {template.extra_credit_policy_text[:50]}...")
                self.root.after(10, set_extra_credit_text)  # Delay 10ms
            else:
                print("DEBUG: No custom extra credit text in template, preserving default UI text")
            
        except Exception as e:
            print(f"Error loading template: {e}")
            import traceback
            traceback.print_exc()

    def clear_all_fields(self):
        """Clear all form fields and reset to default state"""
        # Course info fields
        if hasattr(self, 'entry_course_num'):
            self.entry_course_num.delete(0, tk.END)
        if hasattr(self, 'entry_course_title'):
            self.entry_course_title.delete(0, tk.END)
        if hasattr(self, 'entry_term'):
            self.entry_term.delete(0, tk.END)
        if hasattr(self, 'entry_credits'):
            self.entry_credits.delete(0, tk.END)
        if hasattr(self, 'entry_prerequisites'):
            self.entry_prerequisites.delete(0, tk.END)
        if hasattr(self, 'entry_meeting_times'):
            self.entry_meeting_times.delete(0, tk.END)
        if hasattr(self, 'entry_location'):
            self.entry_location.delete(0, tk.END)
        if hasattr(self, 'txt_description'):
            self.txt_description.delete("1.0", tk.END)
        
        # Instructor info fields
        if hasattr(self, 'entry_instr_name'):
            self.entry_instr_name.delete(0, tk.END)
        if hasattr(self, 'entry_instr_office'):
            self.entry_instr_office.delete(0, tk.END)
        if hasattr(self, 'entry_instr_phone'):
            self.entry_instr_phone.delete(0, tk.END)
        if hasattr(self, 'entry_instr_email'):
            self.entry_instr_email.delete(0, tk.END)
        if hasattr(self, 'entry_instr_office_hours'):
            self.entry_instr_office_hours.delete(0, tk.END)
        
        # Clear TAs/Sections properly
        if hasattr(self, 'ta_entries'):
            for ta_widgets in self.ta_entries[:]:  # Create a copy to iterate over
                if hasattr(ta_widgets[0], 'master'):  # Check if widget still exists
                    ta_widgets[0].master.destroy()  # Destroy the parent frame
            self.ta_entries.clear()
        
        # Clear course objectives properly
        if hasattr(self, 'objective_entries'):
            for obj_dict in self.objective_entries[:]:
                if obj_dict["frame"].winfo_exists():
                    obj_dict["frame"].destroy()
            self.objective_entries.clear()
        
        # Clear Student Learning Outcomes properly
        if hasattr(self, 'outcome_entries'):
            for outcome_dict in self.outcome_entries[:]:
                if outcome_dict["frame"].winfo_exists():
                    outcome_dict["frame"].destroy()
            self.outcome_entries.clear()
        
        # Clear schedule properly
        if hasattr(self, 'schedule_entries'):
            for entry_dict in self.schedule_entries[:]:
                # Handle the new grid-based schedule entries
                if "readings_frame" in entry_dict:
                    entry_dict["readings_frame"].destroy()
                if "delete_btn" in entry_dict:
                    entry_dict["delete_btn"].destroy()
                entry_dict["date"].destroy()
                entry_dict["topic"].destroy()
                entry_dict["work_due"].destroy()
            self.schedule_entries.clear()
        
        # Clear assignment categories properly
        if hasattr(self, 'category_frames'):
            for category_dict in self.category_frames[:]:
                if category_dict["frame"].winfo_exists():
                    category_dict["frame"].destroy()
            self.category_frames.clear()
        
        # Clear learning objectives table properly
        if hasattr(self, 'learning_objectives_entries'):
            for category, entries in self.learning_objectives_entries.items():
                if 'frame' in entries and entries['frame'].winfo_exists():
                    entries['frame'].destroy()
            self.learning_objectives_entries.clear()
        
        # Clear any materials/required materials fields
        if hasattr(self, 'materials_text'):
            self.materials_text.delete("1.0", tk.END)
        if hasattr(self, 'fee_entry'):
            self.fee_entry.delete(0, tk.END)
        
        # Clear policy text fields and restore defaults
        if hasattr(self, 'canvas_policy_text'):
            self.canvas_policy_text.delete("1.0", tk.END)
            # Import default values
            from constants import canvas_policy_default
            self.canvas_policy_text.insert("1.0", canvas_policy_default)
        if hasattr(self, 'technology_policy_text'):
            self.technology_policy_text.delete("1.0", tk.END)
            from constants import technology_policy_default
            self.technology_policy_text.insert("1.0", technology_policy_default)
        if hasattr(self, 'communication_policy_text'):
            self.communication_policy_text.delete("1.0", tk.END)
            from constants import class_communication_policy_default
            self.communication_policy_text.insert("1.0", class_communication_policy_default)
        if hasattr(self, 'support_text'):
            self.support_text.delete("1.0", tk.END)
            from constants import assignment_support_default
            self.support_text.insert("1.0", assignment_support_default)
        if hasattr(self, 'late_policy_text'):
            self.late_policy_text.delete("1.0", tk.END)
            # Restore default late policy text
            default_late_text = "Late assignments will be penalized 10% per day late unless prior arrangements have been made with the instructor due to documented emergency or illness. Contact instructor as soon as possible if you anticipate being unable to meet a deadline."
            self.late_policy_text.insert("1.0", default_late_text)
        if hasattr(self, 'extra_credit_text'):
            self.extra_credit_text.delete("1.0", tk.END)
            # Restore default extra credit text
            default_extra_credit_text = "Extra credit opportunities may be available at the instructor's discretion. These will be announced in class and posted on Canvas when available."
            self.extra_credit_text.insert("1.0", default_extra_credit_text)
            
        # Reset dropdown selections to default values
        if hasattr(self, 'late_policy_var'):
            self.late_policy_var.set("Custom")
        if hasattr(self, 'extra_credit_var'):
            self.extra_credit_var.set("Custom")
    
        # Update any previews that might be affected
        if hasattr(self, 'update_document_preview'):
            self.update_document_preview()
        if hasattr(self, 'update_lo_preview'):
            self.update_lo_preview()
        if hasattr(self, 'update_outcomes_references'):
            self.update_outcomes_references()

    def load_template(self, template):
        """Load a template into the form fields"""
        # Clear existing data first
        self.clear_all_fields()
        
        # Load course info
        if hasattr(self, 'entry_course_num'):
            self.entry_course_num.insert(0, template.course_code or "")
        if hasattr(self, 'entry_course_title'):
            self.entry_course_title.insert(0, template.title or "")
        if hasattr(self, 'entry_term'):
            self.entry_term.insert(0, template.semester or "")
        if hasattr(self, 'entry_credits'):
            self.entry_credits.insert(0, template.credits or "")
        if hasattr(self, 'entry_prerequisites'):
            self.entry_prerequisites.insert(0, template.prerequisites or "")
        if hasattr(self, 'entry_meeting_times'):
            self.entry_meeting_times.insert(0, f"{template.class_days} {template.class_times}" if template.class_days and template.class_times else "")
        if hasattr(self, 'entry_location'):
            self.entry_location.insert(0, template.classroom or "")
        if hasattr(self, 'txt_description'):
            self.txt_description.insert("1.0", template.description or "")
        
        # Load instructor info
        if hasattr(self, 'entry_instr_name'):
            self.entry_instr_name.insert(0, template.instructor_name or "")
        if hasattr(self, 'entry_instr_office'):
            self.entry_instr_office.insert(0, template.instructor_office or "")
        if hasattr(self, 'entry_instr_phone'):
            self.entry_instr_phone.insert(0, template.instructor_phone or "")
        if hasattr(self, 'entry_instr_email'):
            self.entry_instr_email.insert(0, template.instructor_email or "")
        if hasattr(self, 'entry_instr_office_hours'):
            self.entry_instr_office_hours.insert(0, template.instructor_office_hours or "")
        
        # Load objectives and outcomes if available
        if hasattr(template, 'objectives') and template.objectives:
            for obj in template.objectives:
                if hasattr(self, 'objective_entries'):
                    # Add objective entry
                    self.add_objective_entry(obj)
        
        if hasattr(template, 'outcomes') and template.outcomes:
            for outcome in template.outcomes:
                if hasattr(self, 'outcome_entries'):
                    # Add outcome entry
                    self.add_outcome_entry(outcome)
        
        # Load TAs if available
        if hasattr(template, 'tas') and template.tas:
            for ta in template.tas:
                if hasattr(self, 'ta_entries'):
                    # Add TA entry
                    self.add_ta_entry(ta.get('name', ''), ta.get('email', ''), ta.get('office_hours', ''), '', '')

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
            self.update_document_preview()
        
        remove_btn = ttk.Button(frame, text="X", command=remove_objective, style="Delete.TButton")
        remove_btn.pack(side=tk.LEFT, padx=5)
        
        obj_dict = {"frame": frame, "number": number_label, "entry": entry}
        self.objective_entries.append(obj_dict)
        
        entry.bind("<KeyRelease>", lambda e: self.update_document_preview())
        
        return obj_dict

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
        entry.bind("<KeyRelease>", lambda e: self.update_all_previews())
        
        # Update references immediately
        self.update_outcomes_references()
        
        return entry_dict

    def add_ta_entry(self, name="", email="", office_hours="", class_room="", class_time=""):
        """Add a TA entry with given information"""
        # This would be implemented in the UI tabs mixin
        pass

    def update_all_previews(self):
        """Update both previews"""
        self.update_outcomes_references()
        self.update_lo_preview()

    def renumber_outcomes(self):
        """Update the numbering of outcome entries after one is removed"""
        for i, entry in enumerate(self.outcome_entries):
            entry["number"].config(text=f"{i + 1}.")
        
        # Also update the preview
        self.update_lo_preview()

    def renumber_objectives(self):
        """Update numbering for course objectives"""
        for i, obj in enumerate(self.objective_entries):
            obj["number"].config(text=f"{i+1}.")

    def setup_styles(self):
        """Set up custom styles for the application"""
        style = ttk.Style()

        # General styles
        style.configure("TButton", padding=5, font=("Arial", 10))
        style.configure("TLabel", font=("Arial", 10))
        style.configure("TEntry", padding=5)
        style.configure("TFrame", background="#f5f5f5")
        
        # Custom styles
        style.configure("Heading.TLabel", font=("Arial", 12, "bold"))
        style.configure("Action.TButton", foreground="black", background="#0078D7", font=("Arial", 10, "bold"))
        style.map("Action.TButton", background=[("active", "#005A9E")])
        
        # Improve delete button appearance - changed foreground to black for better visibility
        style.configure("Delete.TButton", foreground="black", background="#D9534F", font=("Arial", 10, "bold"), 
                       padding=3, width=3)
        style.map("Delete.TButton", background=[("active", "#C9302C")])
        
        style.configure("Small.TButton", font=("Arial", 8))
        
        # Preview styles
        style.configure("Preview.TFrame", background="white")
        style.configure("Italic.TLabel", font=("Arial", 10, "italic"))

    def add_ta_entry(self, name="", email="", office_hours="", class_room="", class_time=""):
        """Add a Section entry with given information"""
        frame = ttk.Frame(self.ta_container)
        frame.pack(anchor="w", pady=2)
        
        entries = []
        values = [name, email, office_hours, class_room, class_time]
        for i, (label, width) in enumerate([("TA Name:", 25), ("Email:", 30), ("Office Hours:", 45), ("Section Meeting Place:",30),("Section Meeting Time:",30)]):
            ttk.Label(frame, text=label).pack(side=tk.LEFT)
            entry = ttk.Entry(frame, width=width)
            entry.pack(side=tk.LEFT, padx=5)
            if i < len(values):
                entry.insert(0, values[i])
            entries.append(entry)
            
        self.ta_entries.append(entries)
        
        def remove_ta():
            frame.destroy()
            self.ta_entries.remove(entries)
        
        remove_btn = ttk.Button(frame, text="X", 
                              command=remove_ta,
                              style="Delete.TButton")
        remove_btn.pack(side=tk.LEFT, padx=2)

    def add_ta(self):
        """Add a new Section entry"""
        self.add_ta_entry()

    def run(self):
        """Start the application main loop"""
        self.root.mainloop()

    def on_tab_changed(self, event):
        """Update preview when tab changes"""
        try:
            selected_tab = self.notebook.select()
            tab_id = self.notebook.index(selected_tab)
            # Document Preview tab
            if tab_id == 6:  # Document Preview tab
                if hasattr(self, 'update_document_preview'):
                    self.update_document_preview()
            elif tab_id == 2:  # Learning Objectives tab
                if hasattr(self, 'update_lo_preview'):
                    self.update_lo_preview()
        except Exception as e:
            print(f"Error changing tabs: {e}")

    def gather_content(self):
        """Gather all content from form fields"""
        try:
            content = {
                "course_info": {
                    "course_num": self.entry_course_num.get() if hasattr(self, 'entry_course_num') else "",
                    "course_title": self.entry_course_title.get() if hasattr(self, 'entry_course_title') else "",
                    "term": self.entry_term.get() if hasattr(self, 'entry_term') else "",
                    "credits": self.entry_credits.get() if hasattr(self, 'entry_credits') else "",
                    "prerequisites": self.entry_prerequisites.get() if hasattr(self, 'entry_prerequisites') else "",
                    "meeting_times": self.entry_meeting_times.get() if hasattr(self, 'entry_meeting_times') else "",
                    "location": self.entry_location.get() if hasattr(self, 'entry_location') else "",
                    "description": self.txt_description.get("1.0", tk.END).strip() if hasattr(self, 'txt_description') else "",
                    "objectives": "\n".join([obj["entry"].get().strip() for obj in self.objective_entries if obj["entry"].get().strip()]) if hasattr(self, 'objective_entries') else ""
                },
                "instructor_info": {
                    "name": self.entry_instr_name.get() if hasattr(self, 'entry_instr_name') else "",
                    "office": self.entry_instr_office.get() if hasattr(self, 'entry_instr_office') else "",
                    "phone": self.entry_instr_phone.get() if hasattr(self, 'entry_instr_phone') else "",
                    "email": self.entry_instr_email.get() if hasattr(self, 'entry_instr_email') else "",
                    "office_hours": self.entry_instr_office_hours.get() if hasattr(self, 'entry_instr_office_hours') else ""
                },
                "tas": [],
                "outcomes": [],
                "schedule": [],
                "grading_categories": [],
                "optional_policies": {},
                "late_policy": "",
                "extra_credit_policy": "",
                "canvas_policy": self.canvas_policy_text.get("1.0", tk.END).strip() if hasattr(self, 'canvas_policy_text') else "",
                "technology_policy": self.technology_policy_text.get("1.0", tk.END).strip() if hasattr(self, 'technology_policy_text') else "",
                "communication_policy": self.communication_policy_text.get("1.0", tk.END).strip() if hasattr(self, 'communication_policy_text') else "",
                "support_policy": self.support_text.get("1.0", tk.END).strip() if hasattr(self, 'support_text') else "",
                "learning_objectives": {}
            }

            # Gather outcomes from the new numbered list
            if hasattr(self, 'outcome_entries'):
                content["outcomes"] = [
                    {"text": outcome_entry["entry"].get().strip()}
                    for outcome_entry in self.outcome_entries if outcome_entry["entry"].get().strip()
                ]

            # Gather TAs
            if hasattr(self, 'ta_entries'):
                for ta_entry_widgets in self.ta_entries:
                    if len(ta_entry_widgets) == 5:
                         content["tas"].append({
                            "name": ta_entry_widgets[0].get(),
                            "email": ta_entry_widgets[1].get(),
                            "office_hours": ta_entry_widgets[2].get(),
                            "class_room": ta_entry_widgets[3].get(),
                            "class_time": ta_entry_widgets[4].get()
                        })

            # Gather schedule
            if hasattr(self, 'schedule_entries'):
                for entry in self.schedule_entries:
                    content["schedule"].append({
                        "date": entry["date"].get(),
                        "topic": entry["topic"].get(),
                        "readings": entry["readings"].get("1.0", tk.END).strip(),
                        "work_due": entry["work_due"].get()
                    })

            # Gather assignment categories
            if hasattr(self, 'category_frames'):
                for category in self.category_frames:
                    assignments_data = []
                    if "assignments" in category:
                        for assignment in category["assignments"]:
                            assignments_data.append({
                                "title": assignment["title"].get(),
                                "due_date": assignment["due date"].get(),
                                "points": assignment["points"].get(),
                                "description": assignment["description"].get("1.0", tk.END).strip()
                            })
                    content["grading_categories"].append({
                        "name": category["name"].get(),
                        "weight": category["weight"].get(),
                        "description": category["description"].get("1.0", tk.END).strip(),
                        "assignments": assignments_data
                    })

            # Gather learning objectives table data
            if hasattr(self, 'learning_objectives_entries'):
                for category_key, entries in self.learning_objectives_entries.items():
                    category_name = entries.get('name_entry').get() if 'name_entry' in entries else category_key
                    if category_name:
                        content["learning_objectives"][category_name] = {
                            "slo": entries['slo'].get("1.0", tk.END).strip(),
                            "assignments": entries['assignments'].get("1.0", tk.END).strip(),
                            "course_specific": entries['course_specific'].get("1.0", tk.END).strip()
                        }

            # Gather optional policies boolean values
            if hasattr(self, 'optional_policies'):
                for policy_name, policy_var in self.optional_policies.items():
                    content["optional_policies"][policy_name] = policy_var.get()

            # Keep only the instructor-specific policy options
            content["optional_policies"]["outside_support"] = self.outside_support_var.get() if hasattr(self, 'outside_support_var') else True
            content["optional_policies"]["show_gen_ed"] = self.show_gen_ed.get() if hasattr(self, 'show_gen_ed') else True

            if hasattr(self, 'materials_text'):
                content["materials"] = {
                    "required": self.materials_text.get("1.0", tk.END).strip(),
                    "fee": self.fee_entry.get() if hasattr(self, 'fee_entry') else ""
                }

            # Add policy details if available
            if hasattr(self, 'late_policy_var'):
                content["late_policy"] = self.late_policy_var.get()
                if hasattr(self, 'late_policy_text'):
                     content["late_policy_text"] = self.late_policy_text.get("1.0", tk.END).strip()

            if hasattr(self, 'extra_credit_var'):
                content["extra_credit_policy"] = self.extra_credit_var.get()
                if hasattr(self, 'extra_credit_text'):
                    content["extra_credit_policy_text"] = self.extra_credit_text.get("1.0", tk.END).strip()

            return content

        except Exception as e:
            print(f"Error gathering content: {e}")
            import traceback
            traceback.print_exc()
            return {
                "course_info": {
                    "course_num": "", "course_title": "", "term": "", "credits": "", "prerequisites": "",
                    "meeting_times": "", "location": "", "description": "", "objectives": ""
                },
                "instructor_info": {"name": "", "office": "", "phone": "", "email": "", "office_hours": ""},
                "tas": [], "outcomes": [], "schedule": [], "grading_categories": [],
                "optional_policies": {}, "learning_objectives": {}, "materials": {"required": "", "fee": ""},
                "late_policy": "", "extra_credit_policy": "", "canvas_policy": "", "technology_policy": "",
                "communication_policy": "", "support_policy": ""
            }

    def add_category(self, name="", weight="", description=""):
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
        if hasattr(self, 'add_mousewheel_scrolling'):
            self.add_mousewheel_scrolling(desc_text)
        
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
            
            # Description for the assignment
            ttk.Label(assignment_frame, text="Description:").pack(side=tk.LEFT)
            description_text = scrolledtext.ScrolledText(assignment_frame, width=40, height=3, wrap=tk.WORD)
            description_text.pack(side=tk.LEFT, padx=2)
            
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
                "points": points_entry,
                "description": description_text
            }
            assignments.append(assignment_dict)
        
        ttk.Button(frame, text="Add Assignment", command=add_assignment).pack(anchor="w", padx=5, pady=5)
        
        def remove_category():
            frame.destroy()
            if hasattr(self, 'category_frames'):
                self.category_frames.remove(category_dict)
        
        ttk.Button(frame, text="Remove Category", command=remove_category).pack(anchor="w", padx=5, pady=5)
        
        category_dict = {
            "frame": frame,
            "name": name_entry,
            "weight": weight_entry,
            "description": desc_text,
            "assignments": assignments
        }
        
        if not hasattr(self, 'category_frames'):
            self.category_frames = []
        self.category_frames.append(category_dict)
        return category_dict

    def add_assignment_to_category(self, category, title="", due_date="", points=""):
        """Add assignment to existing category"""
        # This method may be used by template loading
        pass

    def clear_all_entries(self):
        """Clear all form entries"""
        # Clear basic course info
        if hasattr(self, 'entry_course_num'):
            self.entry_course_num.delete(0, tk.END)
        if hasattr(self, 'entry_course_title'):
            self.entry_course_title.delete(0, tk.END)
        if hasattr(self, 'entry_term'):
            self.entry_term.delete(0, tk.END)
        if hasattr(self, 'entry_credits'):
            self.entry_credits.delete(0, tk.END)
        if hasattr(self, 'entry_prerequisites'):
            self.entry_prerequisites.delete(0, tk.END)
        if hasattr(self, 'entry_meeting_times'):
            self.entry_meeting_times.delete(0, tk.END)
        if hasattr(self, 'entry_location'):
            self.entry_location.delete(0, tk.END)
        if hasattr(self, 'txt_description'):
            self.txt_description.delete("1.0", tk.END)

        # Clear instructor info
        if hasattr(self, 'entry_instr_name'):
            self.entry_instr_name.delete(0, tk.END)
        if hasattr(self, 'entry_instr_office'):
            self.entry_instr_office.delete(0, tk.END)
        if hasattr(self, 'entry_instr_phone'):
            self.entry_instr_phone.delete(0, tk.END)
        if hasattr(self, 'entry_instr_email'):
            self.entry_instr_email.delete(0, tk.END)
        if hasattr(self, 'entry_instr_office_hours'):
            self.entry_instr_office_hours.delete(0, tk.END)

        # Clear dynamic lists
        if hasattr(self, 'ta_entries'):
            for ta_widgets in self.ta_entries[:]:  # Create a copy to iterate over
                if hasattr(ta_widgets[0], 'master'):  # Check if widget still exists
                    ta_widgets[0].master.destroy()  # Destroy the parent frame
            self.ta_entries.clear()

        if hasattr(self, 'outcome_entries'):
            for outcome_dict in self.outcome_entries[:]:
                if outcome_dict["frame"].winfo_exists():
                    outcome_dict["frame"].destroy()
            self.outcome_entries.clear()

        if hasattr(self, 'objective_entries'):
            for obj_dict in self.objective_entries[:]:
                if obj_dict["frame"].winfo_exists():
                    obj_dict["frame"].destroy()
            self.objective_entries.clear()

        if hasattr(self, 'schedule_entries'):
            for entry_dict in self.schedule_entries[:]:
                if entry_dict["frame"].winfo_exists():
                    entry_dict["frame"].destroy()
            self.schedule_entries.clear()

        if hasattr(self, 'category_frames'):
            for category_dict in self.category_frames[:]:
                if category_dict["frame"].winfo_exists():
                    category_dict["frame"].destroy()
            self.category_frames.clear()

        if hasattr(self, 'learning_objectives_entries'):
            for category, entries in self.learning_objectives_entries.items():
                if 'frame' in entries and entries['frame'].winfo_exists():
                    entries['frame'].destroy()
            self.learning_objectives_entries.clear()

    def _set_policy_dropdowns(self, template):
        """Set policy dropdown selections and boolean variables based on template"""
        try:
            # Set optional policy boolean variables if they exist in the template
            if hasattr(template, 'optional_policies') and template.optional_policies:
                for policy_name, value in template.optional_policies.items():
                    # Map the template policy names to our boolean variables (only instructor-specific ones)
                    if policy_name == 'late_submissions' and hasattr(self, 'late_submissions_policy_var'):
                        self.late_submissions_policy_var.set(value)
                    elif policy_name == 'extra_credit' and hasattr(self, 'extra_credit_policy_var'):
                        self.extra_credit_policy_var.set(value)
                    elif policy_name == 'canvas' and hasattr(self, 'canvas_policy_var'):
                        self.canvas_policy_var.set(value)
                    elif policy_name == 'technology' and hasattr(self, 'technology_policy_var'):
                        self.technology_policy_var.set(value)
                    elif policy_name == 'communication' and hasattr(self, 'communication_policy_var'):
                        self.communication_policy_var.set(value)
                    elif policy_name == 'outside_support' and hasattr(self, 'outside_support_var'):
                        self.outside_support_var.set(value)
                    elif policy_name == 'show_gen_ed' and hasattr(self, 'show_gen_ed'):
                        self.show_gen_ed.set(value)
            
            # Set individual policy variables if they exist directly on the template
            if hasattr(template, 'grading_rounding') and hasattr(self, 'grading_rounding_var'):
                self.grading_rounding_var.set(template.grading_rounding)
            if hasattr(template, 'use_simplified_policies') and hasattr(self, 'use_simplified_policies_var'):
                self.use_simplified_policies_var.set(template.use_simplified_policies)
                
            # Set dropdown policy selections
            if hasattr(template, 'late_policy') and hasattr(self, 'late_policy_var'):
                if template.late_policy in self.late_policies:
                    self.late_policy_var.set(template.late_policy)
            if hasattr(template, 'extra_credit_policy') and hasattr(self, 'extra_credit_var'):
                if template.extra_credit_policy in self.extra_credit_policies:
                    self.extra_credit_var.set(template.extra_credit_policy)
                    
        except Exception as e:
            print(f"Error setting policy dropdowns: {e}")
            import traceback
            traceback.print_exc()

    def save_template(self):
        """Save current form content as a template"""
        # This method can be implemented later if needed
        pass

    def import_schedule(self):
        """Import schedule from file"""
        # Delegate to the UI tabs implementation
        if hasattr(self, 'import_schedule') and hasattr(UITabsMixin, 'import_schedule'):
            UITabsMixin.import_schedule(self)

    def export_schedule(self):
        """Export schedule to file"""
        # Delegate to the UI tabs implementation  
        if hasattr(self, 'export_schedule') and hasattr(UITabsMixin, 'export_schedule'):
            UITabsMixin.export_schedule(self)

    def export_schedule_example(self):
        """Export an example schedule"""
        # Delegate to the UI tabs implementation
        if hasattr(self, 'export_schedule_example') and hasattr(UITabsMixin, 'export_schedule_example'):
            UITabsMixin.export_schedule_example(self)

if __name__ == "__main__":
    app = HistorySyllabusGenerator()
    app.run()
