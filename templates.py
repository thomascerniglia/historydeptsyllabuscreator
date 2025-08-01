from constants import *

class SyllabusTemplate:
    """Class to represent a syllabus template"""
    def __init__(self, course_code, title, description="", objectives=None, outcomes=None):
        self.course_code = course_code
        self.title = title
        self.description = description
        self.objectives = objectives or []
        self.outcomes = outcomes or []
        self.schedule = []
        self.grading_categories = []
        self.learning_objectives = {}

def load_default_templates():
    """Load a single default template for testing purposes"""
    try:
        # Create and initialize a test template
        test_template = SyllabusTemplate(
            course_code="TEST1000",
            title="TEST TEMPLATE",
            description="This is a test template for the History Syllabus Generator.",
            objectives=[
                "Understand the principles of historical research.",
                "Develop critical thinking skills through analysis of primary sources.",
                "Learn to construct historical arguments based on evidence."
            ],
            outcomes=[
                "Demonstrate ability to analyze primary sources.",
                "Write clear and coherent historical essays.",
                "Present research findings effectively."
            ]
        )
        # Then set all the properties
        test_template.instructor_name = "Dr. Jane Doe"
        test_template.prerequisites = "HIS100 or instructor permission" 
        test_template.instructor_office = "Room 123, History Building"
        test_template.instructor_phone = "555-123-4567"
        test_template.instructor_email = "jane.doe@university.edu"
        test_template.instructor_office_hours = "MWF 2:00 PM - 4:00 PM" 
        test_template.semester = "Spring 2025"
        test_template.credits = "3"
        test_template.class_days = "MWF"
        test_template.class_times = "10:00 AM - 10:50 AM"
        test_template.classroom = "History Building 202"
        # Add TAs with complete information
        test_template.tas = [
            {
                "name": "John Smith",
                "email": "john.smith@university.edu",
                "office_hours": "TTh 10:00 AM - 12:00 PM"
            },
            {
                "name": "Emily Johnson",
                "email": "emily.johnson@university.edu",
                "office_hours": "MW 1:00 PM - 3:00 PM"
            }
        ]
        test_template.learning_objectives = {
            "Content": {
                "slo": "Identify, describe, and explain key themes, principles, and terminology; the history, theory and/or methodologies used; and social institutions, structures and processes.",
                "assignments": "Outcomes 1-4",
                "course_specific": "Students will analyze primary and secondary sources in short papers, homework assignments, exams, and in-class discussion."
            },
            "Critical Thinking": {
                "slo": "Apply formal and informal qualitative or quantitative analysis effectively to examine the processes and means by which individuals make personal and group decisions. Assess and analyze ethical perspectives in individual and societal decisions.",
                "assignments": "Outcomes 1-4",
                "course_specific": "Students will apply critical thinking skills in written assignments and exams."
            },
            "Communication": {
                "slo": "Communication is the development and expression of ideas in written and oral forms.",
                "assignments": "Outcomes 1-4",
                "course_specific": "Students will present research findings and participate in class discussions."
            }
        }
        # Set default choices for late submissions and extra credit policy dropdowns
        test_template.late_policy = "Standard (10% per day)"
        test_template.extra_credit_policy = "Standard"

        # --- Add sample schedule entries for testing ---
        test_template.schedule = [
            {
                "date": "January 13, 2025",
                "topic": "Syllabus Review; Reconstruction",
                "readings": "AMH 2020 Syllabus [825 words]\n'Reconstruction,' Chapter 15, American Yawp [10390 words]",
                "work_due": "Syllabus Quiz due by 11:59pm"
            },
            {
                "date": "January 15, 2025",
                "topic": "Reconstruction",
                "readings": "Frederick Douglass, 'Remembering the Civil War' (1878)\npp. canonsociety.org/the-civil-war-1867 [1006 words]",
                "work_due": "Reading Response #1"
            },
            {
                "date": "January 17, 2025",
                "topic": "TA Session #1",
                "readings": "All January 15 Readings",
                "work_due": "Discussion Board Post"
            },
            {
                "date": "January 20, 2025",
                "topic": "No Class (Holiday)",
                "readings": "No readings assigned",
                "work_due": "None"
            },
            {
                "date": "January 22, 2025",
                "topic": "The New South",
                "readings": "Henry Grady, 'The New South' Speech (1886)\nAmerican Yawp, Chapter 16 excerpt",
                "work_due": "Short Essay #1 due"
            },
            {
                "date": "January 24, 2025",
                "topic": "TA Session #2",
                "readings": "All January 22 Readings",
                "work_due": "Quiz #1"
            },
            {
                "date": "January 27, 2025",
                "topic": "Gilded Age Politics",
                "readings": "American Yawp, Chapter 18\nSelections from Nast Cartoons",
                "work_due": "Reading Response #2"
            },
            {
                "date": "January 29, 2025",
                "topic": "Labor in the Gilded Age",
                "readings": "Jacob Riis, 'The Working Girls of New York'\nAmerican Yawp, Chapter 18 (cont.)",
                "work_due": "Short Essay #2 due"
            }
        ]
        # --- Add AMH2020 template ---
        amh2020_template = SyllabusTemplate(
            course_code="AMH2020",
            title="United States Since 1877",
            description=(
                "In this course, students will trace the history of the United States from the end of the Reconstruction era to the contemporary era. "
                "Topics will include but are not limited to the rise of Industrialization, the United States' emergence as an actor on the world stage, "
                "Constitutional amendments and their impact, the Progressive era, World War I, the Great Depression and New Deal, World War II, the Civil Rights era, "
                "the Cold War, and the United States since 1989.\n\n"
                "NOTE: All topics in this course will be taught objectively as objects of analysis, without endorsement of particular viewpoints, and will be observed from multiple perspectives. "
                "No lesson is intended to espouse, promote, advance, inculcate, or compel a particular feeling, perception, or belief. Students are encouraged to employ critical thinking and to rely on data and verifiable sources to explore readings and subject matter in this course. All perspectives will be respected in class discussions."
            ),
            objectives=[
                "Address how the Civil War and Reconstruction set the stage for the development of the modern United States.",
                "Explore how US involvement in the Spanish-American War, World War One, and World War Two reshaped US foreign policy and civil society.",
                "Present the origins of the Cold War, its implications for US international relations, and its influence on American political culture.",
                "Enable students to analyze and evaluate the origins and influences of the civil rights movement, the Vietnam War, the women's movement, and New Right conservatism.",
                "Teach students how to analyze historical documents and scholarship from a range of authors and time periods."
            ],
            outcomes=[
                "Describe the factual details of the substantive historical episodes under study.",
                "Identify and analyze foundational developments that shaped American history since 1877 using critical thinking skills.",
                "Demonstrate an understanding of the primary ideas, values, and perceptions that have shaped American history.",
                "Demonstrate competency in civic literacy."
            ]
        )
        amh2020_template.prerequisites = "None."
        amh2020_template.semester = "Spring 2025"
        amh2020_template.credits = "3"
        amh2020_template.class_days = "M, W"
        amh2020_template.class_times = "12:50p - 1:40p"
        amh2020_template.classroom = "MCCC 0100"
        # Instructor/TA fields left blank for user to fill in
        amh2020_template.tas = []
        amh2020_template.learning_objectives = {
            "Content": {
                "slo": "Identify, describe, and explain key themes, principles, and terminology; the history, theory and/or methodologies used; and social institutions, structures and processes.",
                "assignments": "Outcomes 1-4",
                "course_specific": "Students will demonstrate their knowledge of the details of the substantive historical episodes of US History since 1877 by analyzing primary and secondary sources in short papers, homework assignments, exams, and in-class discussion."
            },
            "Critical Thinking": {
                "slo": "Apply formal and informal qualitative or quantitative analysis effectively to examine the processes and means by which individuals make personal and group decisions. Assess and analyze ethical perspectives in individual and societal decisions.",
                "assignments": "Outcomes 1-4",
                "course_specific": "Students will demonstrate their ability in applying qualitative and quantitative methods by analyzing primary and secondary sources in short papers, homework assignments, and exams by using critical thinking skills."
            },
            "Communication": {
                "slo": "Communication is the development and expression of ideas in written and oral forms.",
                "assignments": "Outcomes 1-4",
                "course_specific": (
                    "Students will identify and explain key developments that shaped United States history since 1877 in written assignments and class discussion.\n\n"
                    "Students will demonstrate their understandings of the primary ideas, values, and perceptions that have shaped United States history and will describe them in written assignments, exams, and class discussion."
                )
            }
        }
        amh2020_template.late_policy = "Custom"
        amh2020_template.late_policy_text = "Late assignments will be penalized 10% per day late unless prior arrangements have been made with the instructor due to documented emergency or illness. Contact instructor as soon as possible if you anticipate being unable to meet a deadline."
        amh2020_template.extra_credit_policy = "Custom"
        amh2020_template.extra_credit_policy_text = "Extra credit opportunities may be available at the instructor's discretion. These will be announced in class and posted on Canvas when available."
        amh2020_template.grading_categories = []
        amh2020_template.schedule = []  # Leave schedule empty for user to fill in
        # Set policies
        amh2020_template.optional_policies = {
            "in_class_recording": True,
            "conflict_resolution": True,
            "campus_resources": True,
            "academic_resources": True,
            "late_work": True
        }
        # Add these policies to the AMH2020 template with slight modifications for testing
        amh2020_template.canvas_policy = canvas_policy_default + "\n\n[AMH2020 Template Loaded]"
        amh2020_template.technology_policy = technology_policy_default + "\n\n[AMH2020 Template Loaded]"
        amh2020_template.communication_policy = class_communication_policy_default + "\n\n[AMH2020 Template Loaded]"
        amh2020_template.support_policy = assignment_support_default + "\n\n[AMH2020 Template Loaded]"
        # Add to templates list
        

        # --- Add AMH2010 template ---
        amh2010_template = SyllabusTemplate(
            course_code="AMH2010",
            title="United States History to 1877",
            description=(
                "Examine United States history from before European contact to 1877. Topics include but are not limited to indigenous peoples, the European background, the colonial period, the American Revolution, the Articles of Confederation, the Constitution, issues within the new Republic, sectionalism, manifest destiny, slavery, the American Civil War, and Reconstruction.\n\n"
                "NOTE: All topics in this course will be taught objectively as objects of analysis, without endorsement of particular viewpoints, and will be observed from multiple perspectives. "
                "No lesson is intended to espouse, promote, advance, inculcate, or compel a particular feeling, perception, or belief. Students are encouraged to employ critical thinking and to rely on data and verifiable sources to explore readings and subject matter in this course. All perspectives will be respected in class discussions."
            ),
            objectives=[
                "Analyze primary and secondary sources to understand various historical interpretations and perspectives on significant events, individuals, and movements in early American history. ",
                "Develop critical thinking skills by evaluating evidence, making connections between historical events, and synthesizing information to form reasoned arguments and interpretations.",
                "Analyze historical patterns and trends, identify causes and consequences of historical developments, and assess their significance in shaping the course of American history.",
                "Explore experiences, perspectives, and identities of people in early America, including indigenous peoples, European settlers, enslaved Africans, and other marginalized groups.",
                "Examine the evolution of political institutions, ideologies, and movements in the United States, including the development of colonial governments, the American Revolution, the Constitution, and the Civil War.",
                "Investigate social and economic transformations in early America, including the impact of colonialism, westward expansion, industrialization, slavery, and the market revolution.",
                "Explore the role of religion, philosophy, and intellectual trends in shaping American society and culture, including the influence of religious beliefs on colonial settlements, Enlightenment ideas, and reform movements.",
                "Develop research and writing skills by conducting historical research, analyzing primary sources, and effectively communicating their findings through written assignments and presentations."
            ],
            outcomes=[
                "Students will describe the factual details of the substantive historical episodes under study.",
                "Students will identify and analyze foundational developments that shaped American history from before European contact to 1877 using critical thinking skills.",
                "Students will demonstrate an understanding of the primary ideas, values, and perceptions that have shaped united states history.",
                "Students will demonstrate competency in civic literacy."
            ]
        )
        amh2010_template.prerequisites = "None."
        amh2010_template.semester = "Spring 2025"
        amh2010_template.credits = "3"
        amh2010_template.class_days = "M, W"
        amh2010_template.class_times = "12:50p - 1:40p"
        amh2010_template.classroom = "MCCC 0100"
        # Instructor/TA fields left blank for user to fill in
        amh2010_template.tas = []
        amh2010_template.learning_objectives = {
            "Content": {
                "slo": "Identify, describe, and explain key themes, principles, and terminology; the history, theory and/or methodologies used; and social institutions, structures and processes.",
                "assignments": "Outcomes 1-4",
                "course_specific": "Students will demonstrate their understanding of foundational developments that shaped American history from before European contact to 1877 by analyzing primary and secondary sources in short papers, exams, and through in-class discussion."
            },
            "Critical Thinking": {
                "slo": "Apply formal and informal qualitative or quantitative analysis effectively to examine the processes and means by which individuals make personal and group decisions. Assess and analyze ethical perspectives in individual and societal decisions.",
                "assignments": "Outcomes 1-4",
                "course_specific": "Students will demonstrate their ability in qualitative and quantitative methods by examining primary and secondary sources in short writing assignments, in-class exams, and class discussions, students by using critical thinking skills."
            },
            "Communication": {
                "slo": "Communication is the development and expression of ideas in written and oral forms.",
                "assignments": "Outcomes 1-4",
                "course_specific": (
                    "Students will identify and analyze foundational developments that shaped American history from before European contact to 1877 in written assignments and class discussion.\n\n"
                    "Students will demonstrate an understanding of the primary ideas, values, and perceptions that have shaped United States history and will describe them in  written assignments, periodic exams and class discussion."
                )
            }
        }
        amh2010_template.late_policy = "Custom"
        amh2010_template.late_policy_text = "Late assignments will be penalized 10% per day late unless prior arrangements have been made with the instructor due to documented emergency or illness. Contact instructor as soon as possible if you anticipate being unable to meet a deadline."
        amh2010_template.extra_credit_policy = "Custom"  
        amh2010_template.extra_credit_policy_text = "Extra credit opportunities may be available at the instructor's discretion. These will be announced in class and posted on Canvas when available."
        amh2010_template.grading_categories = []
        amh2010_template.schedule = []  # Leave schedule empty for user to fill in
        # Set policies
        amh2010_template.optional_policies = {
            "in_class_recording": True,
            "conflict_resolution": True,
            "campus_resources": True,
            "academic_resources": True,
            "late_work": True
        }
        # Add these policies to the AMH2010 template with slight modifications for testing
        amh2010_template.canvas_policy = canvas_policy_default + "\n\n[AMH2010 Template Loaded]"
        amh2010_template.technology_policy = technology_policy_default + "\n\n[AMH2010 Template Loaded]"
        amh2010_template.communication_policy = class_communication_policy_default + "\n\n[AMH2010 Template Loaded]"
        amh2010_template.support_policy = assignment_support_default + "\n\n[AMH2010 Template Loaded]"
        # Add to templates list
        return [amh2010_template, amh2020_template]
    except Exception as e:
        print(f"Error loading default templates: {e}")
        return []
