"""
Document Generation Module for History Syllabus Generator
Contains all document generation methods and utilities
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import tempfile
import os
import traceback
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
import docx
try:
    from docx2pdf import convert
except ImportError:
    def convert(input_path, output_path):
        raise RuntimeError("docx2pdf library is not installed. Cannot convert to PDF.")

from constants import *

def add_hyperlink(paragraph, text, url):
    """Add a hyperlink to a paragraph."""
    # This gets the relationship ID for the hyperlink
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    
    # Create the hyperlink element
    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(docx.oxml.shared.qn('r:id'), r_id)
    
    # Create a new run
    new_run = docx.oxml.shared.OxmlElement('w:r')
    rPr = docx.oxml.shared.OxmlElement('w:rPr')
    
    # Add color and underline to the run
    c = docx.oxml.shared.OxmlElement('w:color')
    c.set(docx.oxml.shared.qn('w:val'), '0000FF')  # Blue color
    rPr.append(c)
    
    u = docx.oxml.shared.OxmlElement('w:u')
    u.set(docx.oxml.shared.qn('w:val'), 'single')  # Single underline
    rPr.append(u)
    
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    
    return hyperlink

def process_text_with_hyperlinks(paragraph, text):
    """Process text containing URLs and convert them to hyperlinks"""
    import re
    
    # Better regex patterns
    url_pattern = re.compile(r'(https?://[^\s\'"<>]+[^\s\'"<>.,;:])')
    email_pattern = re.compile(r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})')
    
    # Process URLs first
    url_parts = url_pattern.split(text)
    result_parts = []
    
    for part in url_parts:
        if url_pattern.match(part):
            # This is a URL
            try:
                add_hyperlink(paragraph, part, part)
            except Exception as e:
                # Fallback to plain text if hyperlink creation fails
                paragraph.add_run(part)
        else:
            # Process emails in non-URL text
            email_parts = email_pattern.split(part)
            for email_part in email_parts:
                if email_pattern.match(email_part):
                    # This is an email
                    try:
                        add_hyperlink(paragraph, email_part, f"mailto:{email_part}")
                    except Exception as e:
                        # Fallback to plain text
                        paragraph.add_run(email_part)
                else:
                    # Regular text
                    if email_part.strip():
                        paragraph.add_run(email_part)

class DocumentGenerationMixin:
    """Mixin class containing all document generation methods"""
    
    def generate_pdf(self, export_path, content):
        """Generate PDF directly using ReportLab"""
        # Configure the PDF document with proper page setup
        doc = SimpleDocTemplate(
            export_path,
            pagesize=letter,
            rightMargin=36,
            leftMargin=36,
            topMargin=36,
            bottomMargin=36
        )
        
        # Calculate available width for content
        available_width = letter[0] - 72  # 72 points = 1 inch (36pt margin on each side)
        
        styles = getSampleStyleSheet()
        story = []
    
        # VI. Course Schedule
        story.append(Paragraph("VI. Calendar", styles['CustomHeading1']))
        
        # Create schedule table with preprocessed data
        schedule_data = [["Date", "Topic", "Readings/Preparation", "Work Due"]]
        
        # Process the schedule entries first to better handle text wrapping
        processed_schedule_data = []
        for entry in self.schedule_entries:
            date = entry['date'].get()
            topic = entry['topic'].get()
            readings = entry['readings'].get("1.0", tk.END).strip()
            work_due = entry['work_due'].get()
            
            # Insert intelligent line breaks for better text wrapping in readings column
            if len(readings) > 80:  # If content is long
                # Add line breaks at natural points
                readings = readings.replace(". ", ".\n")
                readings = readings.replace("; ", ";\n")
                readings = readings.replace("] ", "]\n")
                readings = readings.replace("words)", "words)\n")
            
            # Insert breaks for topic if needed
            if len(topic) > 20:
                topic = topic.replace("; ", ";\n")
                topic = topic.replace(": ", ":\n")
            
            processed_schedule_data.append([date, topic, readings, work_due])
        
        schedule_data.extend(processed_schedule_data)
        
        # Set optimized column widths - adjusted for better balance
        col_widths = [
            0.6*inch,   # Date column (slightly narrower)
            1.1*inch,   # Topic column (slightly narrower)
            3.0*inch,   # Readings column (wider for more content)
            1.0*inch    # Work Due column
        ]
        
        # Create schedule table with the adjusted column widths
        schedule_table = Table(schedule_data, colWidths=col_widths, repeatRows=1)
        
        # Apply table styles with improved word wrapping settings
        schedule_table.setStyle(TableStyle([
            # Header row styling
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            
            # Content rows styling
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 7),  # Smaller font for better fit
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # Align text to top
            ('ALIGN', (0, 1), (-1, -1), 'LEFT'),
            
            # Reduced padding for more content space
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            
            # Critical for word wrapping
            ('WORDWRAP', (0, 0), (-1, -1), True)
        ]))
        
        # Add the table to the story
        story.append(schedule_table)
        
        # Define a page template with page numbers
        def add_page_number(canvas, doc):
            """Add page numbers to each page"""
            canvas.saveState()
            canvas.setFont('Helvetica', 9)
            # Position the page number at the bottom center of the page
            canvas.drawCentredString(letter[0]/2, 20, f"Page {canvas.getPageNumber()}")
            canvas.restoreState()
        
        # Build the document with page numbers
        doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)

    def generate_syllabus(self, export_format="docx"):
        """Generate the final syllabus document"""
        if not self.validate_inputs():
            return

        if export_format == "pdf":
            default_ext = ".pdf"
            file_types = [("PDF Document", "*.pdf"), ("Word Document", "*.docx")]
        else:
            default_ext = ".docx"
            file_types = [("Word Document", "*.docx"), ("PDF Document", "*.pdf")]

        export_path = filedialog.asksaveasfilename(
            defaultextension=default_ext,
            filetypes=file_types,
            title="Save Syllabus As"
        )

        if not export_path:
            return

        try:
            if export_path.lower().endswith('.pdf') or export_format == "pdf":
                # 1. Create a temporary .docx file
                with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_docx:
                    docx_path = tmp_docx.name
                doc = self.create_syllabus_document()
                doc.save(docx_path)
                # 2. Convert to PDF
                convert(docx_path, export_path)
                # 3. Remove the temporary .docx
                os.remove(docx_path)
                messagebox.showinfo("Success", f"Syllabus saved as PDF: {export_path}")
            else:
                doc = self.create_syllabus_document()
                doc.save(export_path)
                messagebox.showinfo("Success", f"Syllabus saved as Word document: {export_path}")
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
        """Create the Word document for the syllabus following the exact format from the example"""
        try:
            doc = Document()
            
            # Add page numbers in the footer
            section = doc.sections[0]
            footer = section.footer
            paragraph = footer.paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.text = "Page "
            
            # Add a field for the page number
            run = paragraph.add_run()
            fldChar1 = OxmlElement('w:fldChar')
            fldChar1.set(qn('w:fldCharType'), 'begin')
            run._element.append(fldChar1)
            
            instrText = OxmlElement('w:instrText')
            instrText.set(qn('xml:space'), 'preserve')
            instrText.text = "PAGE"
            run._element.append(instrText)
            
            fldChar2 = OxmlElement('w:fldChar')
            fldChar2.set(qn('w:fldCharType'), 'end')
            run._element.append(fldChar2)
            
            # Continue with the rest of the document creation code
            # Title and Course Info (centered)
            title = doc.add_heading(f"{self.entry_course_num.get()}: {self.entry_course_title.get()}", level=0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            term = doc.add_paragraph(f"{self.entry_term.get()} ({self.entry_credits.get()} credits)")
            term.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            doc.add_paragraph()  # Single space after title
            
            # I. General Information
            doc.add_heading("I. General Information", level=1)
            
            # Meeting times and location - no extra spacing
            p = doc.add_paragraph()
            p.add_run("Meeting days and times: ").bold = True
            p.add_run(self.entry_meeting_times.get())
            
            p = doc.add_paragraph()
            p.add_run("Class location: ").bold = True
            p.add_run(self.entry_location.get())
            
            # Instructor info - compact format
            p = doc.add_paragraph()
            p.add_run("\nInstructor:").bold = True
            
            instructor_info = [
                ("Name:", self.entry_instr_name.get()),
                ("Office:", self.entry_instr_office.get()),
                ("Phone:", self.entry_instr_phone.get()),
                ("Email:", self.entry_instr_email.get()),
                ("Office Hours:", self.entry_instr_office_hours.get())
            ]
            for label, value in instructor_info:
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.25)  # Reduced indentation
                p.paragraph_format.space_after = Pt(0)  # Remove spacing after paragraphs
                p.add_run(f"{label} ").bold = True
                
                if label == "Email:":
                    # Add email as mailto hyperlink
                    try:
                        add_hyperlink(p, value, f"mailto:{value}")
                    except Exception:
                        # Fallback to plain text if hyperlink creation fails
                        p.add_run(value)
                else:
                    # Other fields remain as plain text
                    p.add_run(value)
            
            # Sections - compact format
            if self.ta_entries:
                p = doc.add_paragraph()
                p.add_run("\nSections:").bold = True
                
                for ta in self.ta_entries:
                    p = doc.add_paragraph()
                    p.paragraph_format.left_indent = Inches(0.25)
                    p.paragraph_format.space_after = Pt(0)
                    p.add_run("Name: ").bold = True
                    p.add_run(ta[0].get())
                    
                    p = doc.add_paragraph()
                    p.paragraph_format.left_indent = Inches(0.25)
                    p.paragraph_format.space_after = Pt(0)
                    p.add_run("Email: ").bold = True
                    
                    # Add email as mailto hyperlink
                    try:
                        add_hyperlink(p, ta[1].get(), f"mailto:{ta[1].get()}")
                    except Exception:
                        # Fallback to plain text
                        p.add_run(ta[1].get())
                        
                    p = doc.add_paragraph()
                    p.paragraph_format.left_indent = Inches(0.25)
                    p.paragraph_format.space_after = Pt(0)
                    p.add_run("Office Hours: ").bold = True
                    p.add_run(ta[2].get())

                    p = doc.add_paragraph()
                    p.paragraph_format.left_indent = Inches(0.25)
                    p.paragraph_format.space_after = Pt(0)
                    p.add_run("Class Room: ").bold = True
                    p.add_run(ta[3].get())

                    p = doc.add_paragraph()
                    p.paragraph_format.left_indent = Inches(0.25)
                    p.paragraph_format.space_after = Pt(0)
                    p.add_run("Class Time: ").bold = True
                    p.add_run(ta[4].get())
            
            # Course Description
            doc.add_heading("Course Description", level=1)
            doc.add_paragraph(self.txt_description.get("1.0", tk.END).strip())
            
            # Prerequisites (moved before General Education)
            doc.add_heading("Prerequisites", level=1)
            doc.add_paragraph(self.entry_prerequisites.get().strip() if hasattr(self, 'entry_prerequisites') else "None")
            
            # --- Add General Education Designation (moved after Prerequisites) ---
            if self.show_gen_ed.get():
                # Define the designation variable
                designation = "Social and Behavioral Sciences (S)"  # Default value
                
                # Add the General Education heading and note
                doc.add_heading(f"General Education Designation: {designation}", level=1)
                
                # Add the full General Education description text
                doc.add_paragraph(gen_ed_default)
                
                # Then continue with the existing code for the success message
                course_num = self.entry_course_num.get()
                doc.add_paragraph(f"Your successful completion of {course_num} with a grade of \"C\" or higher will count towards UF's General Education State Core in {designation}. It will also count towards the State of Florida's Civic Literacy requirement.")
            
            # Course Objectives - Add this section (moved after General Education)
            doc.add_heading("Course Objectives", level=1)
            p = doc.add_paragraph("All General Education area objectives can be found ")
            add_hyperlink(p, "here", "https://undergrad.aa.ufl.edu/general-education/gen-ed-program/subject-area-objectives/")
            p.add_run(".")
            
            if hasattr(self, 'objective_entries') and any(obj["entry"].get().strip() for obj in self.objective_entries):
                for i, obj in enumerate(self.objective_entries, 1):
                    obj_text = obj["entry"].get().strip()
                    if obj_text:
                        p = doc.add_paragraph(f"{i}. {obj_text}")
                        p.paragraph_format.left_indent = Inches(0.25)
            
            # Add Student Learning Outcomes
            doc.add_heading("II. Student Learning Outcomes", level=1)
            doc.add_paragraph("A student who successfully completes this course will:")
            
            if hasattr(self, 'outcome_entries') and self.outcome_entries:
                for i, outcome_entry in enumerate(self.outcome_entries, 1):
                    outcome_text = outcome_entry["entry"].get().strip()
                    if outcome_text:
                        p = doc.add_paragraph(f"{i}. {outcome_text}")
                        p.paragraph_format.left_indent = Inches(0.25)
            else:
                # Default outcomes if none provided
                default_outcomes = [
                    "Describe the factual details of the substantive historical episodes under study.",
                    "Identify and analyze foundational developments that shaped history using critical thinking skills.",
                    "Demonstrate an understanding of the primary ideas, values, and perceptions that have shaped history.",
                    "Demonstrate competency in civic literacy."
                ]
                for i, outcome in enumerate(default_outcomes, 1):
                    p = doc.add_paragraph(f"{i}. {outcome}")
                    p.paragraph_format.left_indent = Inches(0.25)
            
            # If General Education is enabled, add the objectives table after the Student Learning Outcomes
            if self.show_gen_ed.get():
                # Add Learning Outcomes table
                doc.add_paragraph()
                doc.add_paragraph(f"Objectives—General Education and {designation}")
                
                # Create table with the correct number of rows and columns
                table_rows = 1  # Header row
                
                # Count rows based on learning objectives entries if they exist
                if hasattr(self, 'learning_objectives_entries') and self.learning_objectives_entries:
                    for category, entries in self.learning_objectives_entries.items():
                        # Skip entries that were removed
                        if 'frame' in entries and not entries['frame'].winfo_exists():
                            continue
                        table_rows += 1
                else:
                    # Default 3 categories if no custom entries
                    table_rows += 3
                    
                table = doc.add_table(rows=table_rows, cols=4)
                table.style = 'Table Grid'
                
                # Add header row
                header_cells = table.rows[0].cells
                header_cells[0].text = "CATEGORY"
                
                # Customize the second header based on designation
                slo_header = "SOCIAL SCIENCE SLOS"
                if "Humanities" in designation:
                    slo_header = "HUMANITIES SLOS"
                elif "International" in designation:
                    slo_header = "INTERNATIONAL SLOS"
                elif "Diversity" in designation:
                    slo_header = "DIVERSITY SLOS"
                elif "Biological" in designation:
                    slo_header = "BIOLOGICAL SCIENCES SLOS"
                elif "Physical" in designation:
                    slo_header = "PHYSICAL SCIENCES SLOS"
                elif "Mathematics" in designation:
                    slo_header = "MATHEMATICS SLOS"
                    
                header_cells[1].text = slo_header
                header_cells[2].text = "STATE SLO ASSIGNMENTS"
                header_cells[3].text = "COURSE-SPECIFIC"
                
                # Make header row bold
                for cell in header_cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.bold = True
                
                # Add data rows
                row_idx = 1
                if hasattr(self, 'learning_objectives_entries') and self.learning_objectives_entries:
                    for category, entries in self.learning_objectives_entries.items():
                        # Skip entries that were removed
                        if 'frame' in entries and not entries['frame'].winfo_exists():
                            continue
                        
                        # For custom categories, use the name_entry value
                        if 'name_entry' in entries:
                            category_name = entries['name_entry'].get()
                        else:
                            category_name = category
                        
                        slo_text = entries['slo'].get("1.0", tk.END).strip()
                        assignments_text = entries['assignments'].get("1.0", tk.END).strip()
                        course_specific_text = entries['course_specific'].get("1.0", tk.END).strip()
                        
                        if row_idx < len(table.rows):
                            row = table.rows[row_idx]
                            row.cells[0].text = category_name
                            row.cells[1].text = slo_text
                            row.cells[2].text = assignments_text
                            row.cells[3].text = course_specific_text
                            row_idx += 1
                else:
                    # Fallback to default table if no custom entries exist
                    default_data = [
                        ("Content", 
                         "Identify, describe, and explain key themes, principles, and terminology; the history, theory and/or methodologies used; and social institutions, structures and processes.", 
                         "Outcomes 1-4\n\nStudents will demonstrate their knowledge of the details of the substantive historical episodes by analyzing primary and secondary sources in short papers, homework assignments, exams, and in-class discussion."),
                        ("Critical Thinking", 
                         "Apply formal and informal qualitative or quantitative analysis effectively to examine the processes and means by which individuals make personal and group decisions. Assess and analyze ethical perspectives in individual and societal decisions.", 
                         "Outcomes 1-4\n\nStudents will demonstrate their ability in applying qualitative and quantitative methods by analyzing primary and secondary sources in short papers, homework assignments, and exams by using critical thinking skills."),
                        ("Communication", 
                         "Communication is the development and expression of ideas in written and oral forms.", 
                         "Outcomes 1-4\n\nStudents will identify and explain key developments that shaped history in written assignments and class discussion.\n\nStudents will demonstrate their understandings of the primary ideas, values, and perceptions that have shaped history and will describe them in written assignments, exams, and class discussion.")
                    ]
                    
                    for i, (category, slo, assignments) in enumerate(default_data):
                        if row_idx < len(table.rows):
                            row = table.rows[row_idx]
                            row.cells[0].text = category
                            row.cells[1].text = slo
                            row.cells[2].text = assignments
                            row.cells[3].text = ""
                            row_idx += 1
            
            # III. Graded Work
            doc.add_heading("III. Graded Work", level=1)
            
            # Materials (if provided)
            if hasattr(self, 'materials_text') and self.materials_text.get("1.0", tk.END).strip():
                doc.add_heading("Required Materials", level=2)
                materials_text = self.materials_text.get("1.0", tk.END).strip()
                # --- Use markup parser for formatted output ---
                self.parse_materials_markup(materials_text, doc=doc)
                # Always include the Materials Fee value from the fee_entry
                fee_value = ""
                if hasattr(self, 'fee_entry') and self.fee_entry.get().strip():
                    fee_value = self.fee_entry.get().strip()
                else:
                    fee_value = "0.00"
                p = doc.add_paragraph()
                p.add_run("\nMaterials Fee: $").bold = True
                p.add_run(fee_value)

            # ==== CONTINUING AFTER REQUIRED MATERIALS SECTION ====
            
            # Grading Components (Categories and Assignments)
            if hasattr(self, 'category_frames') and self.category_frames:
                doc.add_heading("Grading Components", level=2)
                
                # Create a table for categories and weights
                table = doc.add_table(rows=1, cols=2)
                table.style = 'Table Grid'
                
                # Add header row
                header_cells = table.rows[0].cells
                header_cells[0].text = "Category"
                header_cells[1].text = "Weight"
                
                # Make header row bold
                for cell in header_cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.bold = True
                
                # Add data rows
                for category in self.category_frames:
                    name = category["name"].get().strip()
                    weight = category["weight"].get().strip()
                    
                    if name and weight:
                        row_cells = table.add_row().cells
                        row_cells[0].text = name
                        row_cells[1].text = f"{weight}%"
                
                # Category descriptions and assignments
                for category in self.category_frames:
                    name = category["name"].get().strip()
                    desc = category["description"].get("1.0", tk.END).strip()
                    
                    if name and desc:
                        p = doc.add_paragraph()
                        p.add_run(f"\n{name}: ").bold = True
                        p.add_run(desc)
                    
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
                                    p = doc.add_paragraph()
                                    p.add_run(f"{name} Assignments:").bold = True
                                    has_assignments = True
                                
                                p = doc.add_paragraph()
                                p.paragraph_format.left_indent = Inches(0.25)
                                p.add_run(f"• {title}")
                                
                                if due_date:
                                    p.add_run(f" (Due: {due_date})")
                                if points:
                                    p.add_run(f" - {points} points")
                                
                                if description:
                                    p = doc.add_paragraph(description)
                                    p.paragraph_format.left_indent = Inches(0.5)
            
            # Grading Scale
            doc.add_heading("Grading Scale", level=2)
            
            # Create table for grading scale
            table = doc.add_table(rows=1, cols=2)
            table.style = 'Table Grid'
            
            # Header row
            header_cells = table.rows[0].cells
            header_cells[0].text = "Letter Grade"
            header_cells[1].text = "Number Grade"
            
            # Make header row bold
            for cell in header_cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True
            
            # Grade scale data
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
            
            for letter, number in grades:
                row_cells = table.add_row().cells
                row_cells[0].text = letter
                row_cells[1].text = number
            
            # Add UF grading policies note
            p = doc.add_paragraph()
            p.add_run("See the UF Catalog's ") 
            add_hyperlink(p,"Grades and Grading Policies", "https://catalog.ufl.edu/UGRD/academic-regulations/grades-grading-policies/")
            p.add_run(" for information on how UF assigns grade points.")

            # Add rounding statement if enabled
            if hasattr(self, 'grading_rounding_var') and self.grading_rounding_var.get():
                from constants import grading_rounding_default
                p = doc.add_paragraph(grading_rounding_default)

            # Add minimum grade note
            p = doc.add_paragraph()
            p.add_run("Note: A minimum grade of C is required to earn General Education credit.")
            
            # Add University Assessment Policy if checked
            if hasattr(self, 'assessment_policy_var') and self.assessment_policy_var.get():
                doc.add_heading("University Assessment Policies", level=2)
                p = doc.add_paragraph()
                p.add_run("Requirements for make-up exams, assignments, and other work in this course are consistent with university policies that can be found in the ")
                add_hyperlink(p, "Catalog", "https://catalog.ufl.edu/ugrad/current/regulations/info/attendance.aspx")
                p.add_run(".")

            # Instructions for Submitting Written Assignments
            doc.add_heading("Instructions for Submitting Written Assignments", level=1)
            doc.add_paragraph("All written assignments must be submitted as Word documents (.doc or .docx) through the \"Assignments\" portal in Canvas by the specified deadlines. Do NOT send assignments as PDF files.")
            
            # Add Extensions & Make-Up Exams if checked
            if hasattr(self, 'extensions_policy_var') and self.extensions_policy_var.get():
                doc.add_heading("Extensions & Make-Up Exams", level=2)
                p = doc.add_paragraph()
                extensions_text = (
                    "Only the professor can authorize an extension or make-up exam, and all requests must be "
                    "supported by documentation from a medical provider, Student Health Services, the Disability Resource Center, "
                    "or the Dean of Students Office. Requirements for attendance and make-up exams, assignments, and other work "
                    "in this course are consistent with university policies: https://catalog.ufl.edu/ugrad/current/regulations/info/attendance.aspx"
                )
                process_text_with_hyperlinks(p, extensions_text)

            # Add Late Submissions policy
            if (self.late_submissions_policy_var.get() and
                ((hasattr(self, 'late_policy_text') and 
                  self.late_policy_text.get("1.0", tk.END).strip() != "") or
                 (hasattr(self, 'late_policy_var') and 
                  self.late_policy_var.get() in self.late_policies))):
                doc.add_heading("Late Submissions", level=2)
                if hasattr(self, 'late_policy_text') and self.late_policy_text.get("1.0", tk.END).strip():
                    doc.add_paragraph(self.late_policy_text.get("1.0", tk.END).strip())
                elif hasattr(self, 'late_policy_var') and self.late_policy_var.get() in self.late_policies:
                    selected_policy = self.late_policy_var.get()
                    doc.add_paragraph(self.late_policies[selected_policy])
                else:
                    doc.add_paragraph("Late submission policy not specified.")
            
            # Add Extra Credit policy
            if (self.extra_credit_policy_var.get() and
                ((hasattr(self, 'extra_credit_text') and 
                  self.extra_credit_text.get("1.0", tk.END).strip() != "") or
                 (hasattr(self, 'extra_credit_var') and 
                  self.extra_credit_var.get() in self.extra_credit_policies))):
                doc.add_heading("Extra Credit", level=2)
                if hasattr(self, 'extra_credit_text') and self.extra_credit_text.get("1.0", tk.END).strip():
                    doc.add_paragraph(self.extra_credit_text.get("1.0", tk.END).strip())
                elif hasattr(self, 'extra_credit_var') and self.extra_credit_var.get() in self.extra_credit_policies:
                    selected_policy = self.extra_credit_var.get()
                    doc.add_paragraph(self.extra_credit_policies[selected_policy])
                else:
                    doc.add_paragraph("Extra credit policy not specified.")
            

            # Canvas Policy
            if (self.canvas_policy_var.get() and hasattr(self, 'canvas_policy_text')):
                doc.add_heading("Canvas", level=2)
                doc.add_paragraph(self.canvas_policy_text.get("1.0", tk.END).strip())


            # Technology Policy
            if (self.technology_policy_var.get() and hasattr(self, 'technology_policy_text')):
                doc.add_heading("Technology in the Classroom", level=2)
                doc.add_paragraph(self.technology_policy_text.get("1.0", tk.END).strip())


            # Communication Policy
            if (self.communication_policy_var.get() and hasattr(self, 'communication_policy_text')):
                doc.add_heading("Class Communication Policy", level=2)
                doc.add_paragraph(self.communication_policy_text.get("1.0", tk.END).strip())


            # Assignment Support section
            if (self.outside_support_var.get() and hasattr(self, 'support_text')):
                doc.add_heading("Assignment Support Outside the Classroom", level=2)
                doc.add_paragraph(self.support_text.get("1.0", tk.END).strip())
                        
            # IV. Evaluations section
            doc.add_heading("IV. Evaluations", level=1)
            p = doc.add_paragraph()
            eval_text = (
                "UF course evaluation process\n"
                "Students are expected to provide professional and respectful feedback on the quality of "
                "instruction in this course by completing course evaluations online. Students can complete "
                "evaluations in three ways:\n\n"
                "1. The email they receive from GatorEvals,\n"
                "2. Their Canvas course menu under GatorEvals, or\n"
                "3. The central portal at https://my-ufl.bluera.com\n\n"
                "Guidance on how to provide constructive feedback is available at "
                "https://gatorevals.aa.ufl.edu/students/. Students will be notified when the evaluation "
                "period opens. Summaries of course evaluation results are available to students at "
                "https://gatorevals.aa.ufl.edu/public-results/."
            )
            process_text_with_hyperlinks(p, eval_text)
            
            # V. University Policies and Resources
            doc.add_heading("V. University Policies and Resources", level=1)
            
            # Check if simplified policies are enabled
            if hasattr(self, 'use_simplified_policies_var') and self.use_simplified_policies_var.get():
                # Use simplified UF policies
                from constants import uf_policy_simplified
                p = doc.add_paragraph()
                p.add_run(uf_policy_simplified.split(": ")[0] + ": ")
                add_hyperlink(p, "this link", uf_policy_simplified.split(": ")[1])
                p.add_run(".")
            else:
                # Use original detailed policies
                # Accommodations policy
                doc.add_heading("Students requiring accommodation", level=2)
                p = doc.add_paragraph()
                accommodations_text = (
                    "Students with disabilities who experience learning barriers and would like to request academic accommodations "
                    "should connect with the Disability Resource Center by visiting https://disability.ufl.edu/students/get-started/. "
                    "It is important for students to share their accommodation letter with the instructor and discuss their "
                    "access needs as early as possible in the semester."
                )
                process_text_with_hyperlinks(p, accommodations_text)

                # University Honesty Policy
                doc.add_heading("University Honesty Policy", level=2)
                p = doc.add_paragraph()
                p.add_run("UF students are bound by The Honor Pledge which states \"We, the members of the University of Florida community, pledge to hold ourselves and our peers to the highest standards of honor and integrity by abiding by the Honor Code.\" " +
                "On all work submitted for credit by students at the University of Florida, the following pledge is either required or implied: " +
                "\"On my honor, I have neither given nor received unauthorized aid in doing this assignment.\" " +
                "The Conduct Code specifies a number of behaviors that are in violation of this code and the possible sanctions.")
                add_hyperlink(p, " See the UF Conduct Code website for more information", "https://sccr.dso.ufl.edu/process/student-conduct-code/")
                p.add_run(". If you have any questions or concerns, please consult with the instructor or TAs in this class.")

                doc.add_heading("Plagiarism and Related Ethical Violations ", level=2)
                p = doc.add_paragraph()
                p.add_run("Ethical violations such as plagiarism, cheating, academic misconduct (e.g. passing off others' work as your own, reusing old assignments, etc.) " \
                "will not be tolerated and will result in a failing grade in this course. Students must be especially wary of plagiarism. " \
                "The UF Student Honor Code defines plagiarism as follows: "
                "A student shall not represent as the student's own work all or any portion of the work of another. "
                "Plagiarism includes (but is not limited to): a. Quoting oral or written materials, whether published or unpublished, without proper attribution. "
                "b. Submitting a document or assignment which in whole or in part is identical or substantially identical to a document or assignment not authored by the student."
                " Note that plagiarism also includes the use of any artificial intelligence programs, such as ChatGPT. ")
                
                # In-class recording policy (if enabled)
                if hasattr(self, 'in_class_recording_var') and self.in_class_recording_var.get():
                    doc.add_heading("In-class recording", level=2)
                    doc.add_paragraph(recording_policy_default)
                
                # Conflict resolution (if enabled)
                if hasattr(self, 'conflict_resolution_var') and self.conflict_resolution_var.get():
                    doc.add_heading("Procedure for conflict resolution", level=2)
                    p = doc.add_paragraph()
                    p.add_run("Any classroom issues, disagreements or grade disputes should be discussed first between the instructor and the student. ")
                    p.add_run("If the problem cannot be resolved, please contact Nina Caputo (Associate Chair) (")
                    add_hyperlink(p, "ncaputo@ufl.edu", "mailto:ncaputo@ufl.edu")
                    p.add_run(", 352-273-3379). ")
                    p.add_run("Be prepared to provide documentation of the problem, as well as all graded materials for the semester. ")
                    p.add_run("Issues that cannot be resolved departmentally will be referred to the University Ombuds Office (")
                    add_hyperlink(p,"http://www.ombuds.ufl.edu", "http://www.ombuds.ufl.edu")
                    p.add_run("; 352-392-1308) or the Dean of Students Office (")
                    add_hyperlink(p, "http://www.dso.ufl.edu", "http://www.dso.ufl.edu")
                    p.add_run("; 352-392-1261).")
                
                # Campus Resources (if enabled)
                if hasattr(self, 'campus_resources_var') and self.campus_resources_var.get():
                    doc.add_heading("Campus Resources", level=2)
                    
                    # U Matter, We Care
                    p = doc.add_paragraph()
                    p.add_run("U Matter, We Care: ")
                    p.add_run("If you or someone you know is in distress, please contact ")
                    add_hyperlink(p, "umatter@ufl.edu", "mailto:umatter@ufl.edu"), 
                    p.add_run(" 352-392-1575, or visit ")
                    add_hyperlink(p, "U Matter, We Care website", "https://umatter.ufl.edu/")
                    p.add_run(", to refer or report a concern and a team member will reach out to the student in distress.")

                    # Counseling and Wellness Center
                    p = doc.add_paragraph()
                    p.add_run("Counseling and Wellness Center: ")
                p.add_run("Visit the ") 
                add_hyperlink(p, "Counseling and Wellness Center website", "https://counseling.ufl.edu/") 
                p.add_run(" or call 352-392-1575 for information on crisis services as well as non-crisis services.")
                
                # Student Health Care Center
                p = doc.add_paragraph()
                p.add_run("Student Health Care Center: ")
                p.add_run("Call 352-392-1161 for 24/7 information to help you find the care you need, or visit the ")
                add_hyperlink(p, "Student Health Care Center website", "https://shcc.ufl.edu/")
                p.add_run(".")
                
                # University Police Department
                p = doc.add_paragraph()
                p.add_run("University Police Department: ")
                p.add_run("Visit ")
                add_hyperlink(p, "UF Police Department website", "https://police.ufl.edu/") 
                p.add_run(" or call 352-392-1111 (or 9-1-1 for emergencies).")
                
                # UF Health Shands Emergency Room
                p = doc.add_paragraph()
                p.add_run("UF Health Shands Emergency Room / Trauma Center: ")
                p.add_run("For immediate medical care call 352-733-0111 or go to the emergency room at 1515 SW Archer Road, Gainesville, FL 32608; Visit the ")
                add_hyperlink(p, "UF Health Emergency Room and Trauma Center website", "https://ufhealth.org/emergency-room-trauma-center")
                p.add_run(".")
                
                # GatorWell Health Promotion Services
                p = doc.add_paragraph()
                p.add_run("GatorWell Health Promotion Services: ")
                p.add_run("For prevention services focused on optimal wellbeing, including Wellness Coaching for Academic Success, visit the ")
                add_hyperlink(p, "GatorWell website", "https://gatorwell.ufsa.ufl.edu/")
                p.add_run(" or call 352-273-4450.") 
                
                # Student Success Initiative
                p = doc.add_paragraph()
                p.add_run("Student Success Initiative, ")
                add_hyperlink(p, "https://studentsuccess.ufl.edu/", "https://studentsuccess.ufl.edu")
                
                # Field and Fork Pantry
                p = doc.add_paragraph()
                add_hyperlink(p, "Field and Fork Pantry", "https://pantry.fieldandfork.ufl.edu/")
                p.add_run(". Food and toiletries for students experiencing food insecurity.")
                
                # Dean of Students Office
                p = doc.add_paragraph()
                add_hyperlink(p, "Dean of Students Office", "https://care.dso.ufl.edu/")
                p.add_run(". 202 Peabody Hall, 392-1261. Among other services, the DSO assists students who are experiencing situations that compromises their ability to attend classes. This includes family emergencies and medical issues (including mental health crises).")
                
                # Academic Resources (if enabled)
                if self.academic_resources_var.get():
                    doc.add_heading("Academic Resources", level=2)
                    
                    #Career Connections Cernter
                    p = doc.add_paragraph()
                    add_hyperlink(p, "Career Connections Center", "https://career.ufl.edu/")
                    p.add_run(": Reitz Union Suite 1300, 352-392-1601. Career assistance and counseling services.  ")

                    # E-learning support
                    p = doc.add_paragraph()
                    p.add_run("E-learning technical support: Contact the ") 
                    add_hyperlink(p, "UF Computing Help Desk", "http://helpdesk.ufl.edu/")
                    p.add_run(" at 352-392-4357 or via e-mail at ")
                    add_hyperlink(p, "helpdesk@ufl.edu", "mailto:helpdesk@ufl.edu")
                    p.add_run(".")
                    
                    # Library Support
                    p = doc.add_paragraph()
                    add_hyperlink(p, "Library Support", "https://cms.uflib.ufl.edu/ask")
                    p.add_run(": Various ways to receive assistance with respect to using the libraries or finding resources.")
                    
                    # Teaching Center
                    p = doc.add_paragraph()
                    add_hyperlink(p, "Teaching Center", "https://teachingcenter.ufl.edu/")
                    p.add_run(": Broward Hall, 352-392-2010 or to make an appointment 352-392-6420. General study skills and tutoring.")
                    
                    # Writing Studio
                    p = doc.add_paragraph()
                    add_hyperlink(p, "Writing Studio", "https://writing.ufl.edu/writing-studio/")
                    p.add_run(": 2215 Turlington Hall, 352-846-1138. Help brainstorming, formatting, and writing papers.")
                    
                    # Student Complaints On-Campus
                    p = doc.add_paragraph()
                    p.add_run("Student Complaints On-Campus: Visit the ")
                    add_hyperlink(p, "Student Honor Code and Student Conduct Code webpage", "https://sccr.dso.ufl.edu/policies/student-honor-%20code-student-conduct-code/")
                    p.add_run(" for more information.")
            
            # Academic Resources (if enabled)
            
            # VI. Course Schedule (Calendar)
            doc.add_heading("VI. Calendar", level=1)
            
            # Create schedule table if there are entries
            if hasattr(self, 'schedule_entries') and self.schedule_entries:
                # Create table with headers
                table = doc.add_table(rows=1, cols=4)
                table.style = 'Table Grid'
                
                # Set column headers
                header_cells = table.rows[0].cells
                header_cells[0].text = "Date"
                header_cells[1].text = "Topic"
                header_cells[2].text = "Readings/Preparation"
                header_cells[3].text = "Work Due"
                
                # Make header row bold
                for cell in header_cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.bold = True
                
                # Add each schedule entry as a row
                for entry in self.schedule_entries:
                    date_text = entry["date"].get().strip()
                    topic_text = entry["topic"].get().strip()
                    readings_text = entry["readings"].get("1.0", tk.END).strip()
                    work_due_text = entry["work_due"].get().strip()
                    
                    # Skip empty rows
                    if not any([date_text, topic_text, readings_text, work_due_text]):
                        continue
                    
                    row_cells = table.add_row().cells
                    row_cells[0].text = date_text
                    row_cells[1].text = topic_text
                    row_cells[2].text = readings_text
                    row_cells[3].text = work_due_text
            else:
                doc.add_paragraph("Schedule will be provided separately.")

            return doc
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while generating the document: {e}")
            import traceback
            traceback.print_exc()
            # Return an empty document so we don't crash
            return Document()

    def parse_materials_markup(self, text, doc=None):
        """
        Parse Markdown-like markup in Required Materials and add to Word doc.
        - Supports the following syntax:
          - *italic* -> Text between single asterisks (*) will be italicized.
          - **bold** -> Text between double asterisks (**) will be bolded.
          - [text](url) -> Text inside square brackets ([text]) followed by a URL in parentheses (url) will become a hyperlink.
        """
        import re
        
        if doc is None:
            # Just return the text if no document is provided
            return text
        
        # Create paragraph for the content
        paragraph = doc.add_paragraph()
        
        # Regular expression to find markdown patterns
        pattern = re.compile(r'(\*\*.*?\*\*|\*.*?\*|\[.*?\]\(.*?\))')
        
        # Split the text based on the pattern
        parts = pattern.split(text)
        
        # Process each part
        for part in parts:
            if part.startswith('**') and part.endswith('**'):
                # Bold text
                run = paragraph.add_run(part[2:-2])
                run.bold = True
            elif part.startswith('*') and part.endswith('*'):
                # Italic text
                run = paragraph.add_run(part[1:-1])
                run.italic = True
            elif part.startswith('[') and '](' in part and part.endswith(')'):
                # Hyperlink
                # Extract the text and URL
                link_text = part[1:part.index('](')]
                url = part[part.index('](')+2:-1]
                
                # Add hyperlink
                run = paragraph.add_run(link_text)
                run.font.color.rgb = docx.shared.RGBColor(0, 0, 255)  # Blue color
                run.font.underline = True
                
                # Add the actual hyperlink
                # This is a simplification - true hyperlinks need more work with XML
                # For now, we'll just make it look like a hyperlink
            else:
                # Regular text
                paragraph.add_run(part)
        
        return paragraph

    def show_formatting_help(self):
        """Display a popup with formatting help."""
        help_window = tk.Toplevel(self.root)
        help_window.title("Formatting Help")
        help_window.geometry("400x300")

        explanation = """
        Markdown-like Formatting Guide:
        - *italic* -> Text between single asterisks (*) will be italicized.
        - **bold** -> Text between double asterisks (**) will be bolded.
        - [text](url) -> Text inside square brackets ([text]) followed by a URL in parentheses (url) will become a hyperlink.

        Example:
        Input: "This is *italic*, **bold*, and [a link](http://example.com)."
        Output in Word:
          - "italic" will appear italicized.
          - "bold" will appear bolded.
          - "a link" will appear as a clickable hyperlink pointing to "http://example.com".
        """

        text_widget = tk.Text(help_window, wrap=tk.WORD, font=("Arial", 10))
        text_widget.insert(tk.END, explanation)
        text_widget.config(state=tk.DISABLED)  # Make the text read-only
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
