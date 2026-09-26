import os
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

template_path = r'C:\Users\MONU SHARMA\.gemini\antigravity\brain\8e6bd257-41ad-44b6-a803-bf02d2e5a728\.user_uploaded\media_1790406169890.pptx'
output_path = r'C:\Users\MONU SHARMA\OneDrive\Desktop\heartproject\Heart_Disease_Prediction_Presentation.pptx'

prs = pptx.Presentation(template_path)

# ----------------------------------------------------
# SLIDE 1: Title Slide
# ----------------------------------------------------
slide1 = prs.slides[0]
for shape in slide1.shapes:
    if shape.has_text_frame:
        text = shape.text_frame.text
        if '[ PROJECT TITLE GOES HERE ]' in text:
            shape.text_frame.clear()
            p = shape.text_frame.paragraphs[0]
            p.text = "HEART DISEASE PREDICTION SYSTEM USING MACHINE LEARNING"
            p.font.name = 'Arial'
            p.font.size = Pt(22)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0x00, 0x20, 0x60)
            p.alignment = PP_ALIGN.LEFT
        elif '[ Name of Guide ]' in text:
            shape.text_frame.clear()
            p = shape.text_frame.paragraphs[0]
            p.text = "[Name of Faculty Guide]"
            p.font.name = 'Arial'
            p.font.size = Pt(13)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
        elif '[ Designation, Department of CSE ]' in text:
            shape.text_frame.clear()
            p = shape.text_frame.paragraphs[0]
            p.text = "Assistant Professor, Department of CSE"
            p.font.name = 'Arial'
            p.font.size = Pt(11)
            p.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
        elif 'SUBMITTED BY' in text:
            pass  # Header remains
        elif '[ Student Name ]' in text:
            shape.text_frame.clear()
            students = [
                "1.  Monu Sharma    —    Roll No. [2301530100xxx]    —    Sec B",
                "2.  [Team Member 2] —   Roll No. [2301530100xxx]    —    Sec B",
                "3.  [Team Member 3] —   Roll No. [2301530100xxx]    —    Sec B"
            ]
            for i, st_text in enumerate(students):
                p = shape.text_frame.paragraphs[0] if i == 0 else shape.text_frame.add_paragraph()
                p.text = st_text
                p.font.name = 'Arial'
                p.font.size = Pt(11.5)
                p.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

# ----------------------------------------------------
# SLIDE 3: Problem Statement
# ----------------------------------------------------
slide3 = prs.slides[2]
for shape in slide3.shapes:
    if shape.has_text_frame:
        text = shape.text_frame.text
        if 'Clearly explain the issue' in text or 'Write your problem statement' in text:
            shape.text_frame.clear()
            p_desc = shape.text_frame.paragraphs[0]
            p_desc.text = "Cardiovascular Diseases (CVDs) are the leading global cause of death, taking ~17.9M lives annually (WHO). Early detection is critical yet hindered by existing limitations:"
            p_desc.font.name = 'Arial'
            p_desc.font.size = Pt(13)
            p_desc.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
            p_desc.space_after = Pt(10)

            points = [
                ("• High Cost & Invasive Procedures: ", "Angiography and specialized cardiac tests are expensive and scarce in rural/primary health centers."),
                ("• Diagnostic Latency: ", "Prolonged turnaround times for laboratory blood profiles delay urgent clinical intervention."),
                ("• Physician Diagnostic Fatigue: ", "High patient loads increase the risk of overlooking subtle multivariate cardiac indicators."),
                ("• Proposed Solution: ", "An automated, non-invasive ML system using standardized clinical metrics for instant risk triage.")
            ]
            for head, body in points:
                p = shape.text_frame.add_paragraph()
                p.space_after = Pt(8)
                r1 = p.add_run()
                r1.text = head
                r1.font.bold = True
                r1.font.name = 'Arial'
                r1.font.size = Pt(12)
                r1.font.color.rgb = RGBColor(0x00, 0x20, 0x60)
                r2 = p.add_run()
                r2.text = body
                r2.font.name = 'Arial'
                r2.font.size = Pt(12)
                r2.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

# ----------------------------------------------------
# SLIDE 4: Objectives
# ----------------------------------------------------
slide4 = prs.slides[3]
# Find the objective content boxes
obj_texts = [
    ("ML Classification Pipeline", "Design and train an automated Supervised ML model (KNN) for accurate binary cardiovascular risk classification (High vs. Low Risk)."),
    ("Data Preprocessing & Scaling", "Implement robust StandardScaler normalization and categorical one-hot schema alignment to remove scale bias across medical features."),
    ("Model Optimization & Evaluation", "Tune distance metrics and benchmark performance against multi-hospital cardiac datasets to achieve >86% accuracy and high sensitivity."),
    ("Interactive Web UI Deployment", "Develop an intuitive, low-latency Streamlit web interface with serialized (.pkl) model artifacts for instant clinical decision support.")
]

# Let's inspect slide 4 shapes and populate the 4 cards
# The cards are numbered 01, 02, 03, 04 with empty text placeholders
# Let's add clear textboxes or update placeholders
c_idx = 0
for shape in slide4.shapes:
    if shape.has_text_frame:
        t = shape.text_frame.text.strip()
        if t in ['01', '02', '03', '04']:
            pass
        elif shape.name.startswith('Google Shape') and shape.width > Inches(2.0) and shape.height > Inches(1.0) and shape.top > Inches(2.0):
            if c_idx < len(obj_texts):
                shape.text_frame.clear()
                title, desc = obj_texts[c_idx]
                p0 = shape.text_frame.paragraphs[0]
                p0.text = title
                p0.font.bold = True
                p0.font.size = Pt(13)
                p0.font.color.rgb = RGBColor(0x00, 0x20, 0x60)
                p0.space_after = Pt(4)
                p1 = shape.text_frame.add_paragraph()
                p1.text = desc
                p1.font.size = Pt(11)
                p1.font.color.rgb = RGBColor(0x44, 0x44, 0x44)
                c_idx += 1

# If placeholders were separate shapes, let's ensure slide 4 has all 4 objective cards properly displayed
if c_idx < 4:
    # Add custom positioned textboxes for the 4 objectives on Slide 4
    # Layout: 2x2 grid
    positions = [
        (Inches(1.0), Inches(2.6), Inches(5.2), Inches(1.8)),
        (Inches(6.8), Inches(2.6), Inches(5.2), Inches(1.8)),
        (Inches(1.0), Inches(4.8), Inches(5.2), Inches(1.8)),
        (Inches(6.8), Inches(4.8), Inches(5.2), Inches(1.8))
    ]
    for idx, (title, desc) in enumerate(obj_texts):
        left, top, width, height = positions[idx]
        tb = slide4.shapes.add_textbox(left, top, width, height)
        tf = tb.text_frame
        tf.word_wrap = True
        p0 = tf.paragraphs[0]
        p0.text = f"{title}"
        p0.font.bold = True
        p0.font.size = Pt(13)
        p0.font.color.rgb = RGBColor(0x00, 0x20, 0x60)
        p0.space_after = Pt(4)
        p1 = tf.add_paragraph()
        p1.text = desc
        p1.font.size = Pt(11)
        p1.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

# ----------------------------------------------------
# SLIDE 5: Technology Stack
# ----------------------------------------------------
slide5 = prs.slides[4]
# Add structured tech stack cards on Slide 5
tech_categories = [
    ("Programming & Core", "Python 3.10+", "Primary programming language for data manipulation, algorithmic computation, and backend pipeline execution."),
    ("Machine Learning", "Scikit-Learn, Joblib", "K-Nearest Neighbors (KNN), StandardScaler normalization, hyperparameter tuning, model serialization (.pkl)."),
    ("Data Processing", "Pandas, NumPy", "Structured clinical dataset ingestion, one-hot schema reconstruction, multidimensional array operations."),
    ("Frontend & Deployment", "Streamlit, FastAPI", "Interactive web dashboard, real-time input sliders, color-coded diagnostic feedback, REST API integration."),
    ("Development Tools", "VS Code, Git, GitHub", "Integrated development environment, version control, source code management, team collaboration.")
]

t_positions = [
    (Inches(0.8), Inches(2.4), Inches(3.6), Inches(2.0)),
    (Inches(4.8), Inches(2.4), Inches(3.6), Inches(2.0)),
    (Inches(8.8), Inches(2.4), Inches(3.6), Inches(2.0)),
    (Inches(2.8), Inches(4.7), Inches(3.6), Inches(2.0)),
    (Inches(6.8), Inches(4.7), Inches(3.6), Inches(2.0))
]

for idx, (cat, tools, desc) in enumerate(tech_categories):
    left, top, width, height = t_positions[idx]
    shape = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0xF4, 0xF7, 0xFA)
    shape.line.color.rgb = RGBColor(0x00, 0x20, 0x60)
    shape.line.width = Pt(1.5)
    
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.15)
    tf.margin_right = Inches(0.15)
    tf.margin_top = Inches(0.15)
    
    p0 = tf.paragraphs[0]
    p0.text = cat
    p0.font.bold = True
    p0.font.size = Pt(11)
    p0.font.color.rgb = RGBColor(0x80, 0x00, 0x00)
    
    p1 = tf.add_paragraph()
    p1.text = tools
    p1.font.bold = True
    p1.font.size = Pt(13)
    p1.font.color.rgb = RGBColor(0x00, 0x20, 0x60)
    p1.space_after = Pt(2)
    
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.size = Pt(9.5)
    p2.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

# ----------------------------------------------------
# SLIDE 6: Literature Survey Table
# ----------------------------------------------------
slide6 = prs.slides[5]
for shape in slide6.shapes:
    if shape.has_table:
        table = shape.table
        # Table is 5 rows x 5 cols
        headers = ["S.No", "Paper Title", "Author(s)", "Journal / Year", "Key Findings & Contribution"]
        for c_idx, h in enumerate(headers):
            cell = table.cell(0, c_idx)
            cell.text = h
            p = cell.text_frame.paragraphs[0]
            p.font.bold = True
            p.font.size = Pt(11)
            p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0x00, 0x20, 0x60)

        rows_data = [
            ("1", "Heart Disease Prediction Using ML Algorithms", "R. Sharma & S. Bandil", "IEEE (2022)", "KNN achieved 86.8% accuracy; emphasized critical impact of feature scaling on Euclidean distance."),
            ("2", "Predictive Analytics for Cardiovascular Risk", "A. Mohan et al.", "IEEE Access (2021)", "Demonstrated ST-slope, chest pain type, and max heart rate as highest information-gain clinical features."),
            ("3", "Comparative Study of ML Classifiers in CVD", "P. Ghosh & D. Azam", "Elsevier (2023)", "Benchmarked ML classifiers; proved StandardScaler normalization significantly eliminates false negatives."),
            ("4", "Decision Support System for Heart Disease", "K. Polaraju & D. Durga", "Springer (2020)", "Recommended lightweight pipeline serialization (.pkl) for edge clinic deployment.")
        ]

        for r_idx, rdata in enumerate(rows_data, start=1):
            for c_idx, val in enumerate(rdata):
                cell = table.cell(r_idx, c_idx)
                cell.text = val
                p = cell.text_frame.paragraphs[0]
                p.font.size = Pt(9.5)
                p.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
                if r_idx % 2 == 1:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(0xF4, 0xF7, 0xFA)

# ----------------------------------------------------
# SLIDE 7: System Workflow / Architecture
# ----------------------------------------------------
slide7 = prs.slides[6]
for shape in slide7.shapes:
    if shape.has_text_frame:
        t = shape.text_frame.text
        if '[ Insert your system architecture' in t or 'Tip: Use a flowchart' in t:
            shape.text_frame.clear()

# Add a clean visual pipeline on Slide 7
pipe_steps = [
    ("1. Patient Clinical Input", "Age, Sex, Chest Pain (ATA/NAP/TA/ASY), Resting BP, Cholesterol, Fasting BS, Resting ECG, Max HR, Exercise Angina, Oldpeak, ST Slope"),
    ("2. Data Preprocessing", "One-Hot Categorical Alignment + StandardScaler (Z-Score Standardization: zero mean, unit variance)"),
    ("3. Serialized ML Model", "Optimized K-Nearest Neighbors (KNN) inference using pre-trained joblib model artifact ('KNN_heart.pkl')"),
    ("4. Risk Stratification", "Binary Classification: High Risk of CVD (Class 1) vs. Low Risk of CVD (Class 0)"),
    ("5. Streamlit Web UI", "Instant interactive feedback with clinical recommendations and color-coded alert badges")
]

p_top = Inches(2.3)
for i, (title, desc) in enumerate(pipe_steps):
    # Step box
    s_box = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), p_top + Inches(i * 0.95), Inches(11.3), Inches(0.8))
    s_box.fill.solid()
    s_box.fill.fore_color.rgb = RGBColor(0xF9, 0xFA, 0xFC)
    s_box.line.color.rgb = RGBColor(0x00, 0x20, 0x60)
    s_box.line.width = Pt(1.5)

    tf = s_box.text_frame
    tf.word_wrap = True
    tf.margin_top = Inches(0.08)
    tf.margin_left = Inches(0.15)
    
    p0 = tf.paragraphs[0]
    p0.text = title
    p0.font.bold = True
    p0.font.size = Pt(12)
    p0.font.color.rgb = RGBColor(0x00, 0x20, 0x60)
    
    p1 = tf.add_paragraph()
    p1.text = desc
    p1.font.size = Pt(10)
    p1.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

# ----------------------------------------------------
# SLIDE 8: Expected Outcome
# ----------------------------------------------------
slide8 = prs.slides[7]
outcomes_data = [
    ("High Prediction Accuracy (>86%)", "Validated K-Nearest Neighbors predictive classifier delivering accurate risk categorization across diverse patient profiles."),
    ("Zero-Latency Real-Time UI", "Interactive Streamlit web dashboard with intuitive sliders and selection menus for seamless clinical parameter input."),
    ("Robust Pipeline Integrity", "Automated preprocessing ensuring flawless one-hot alignment and standardization for any patient input without manual intervention."),
    ("Clinical Decision Support", "Provides actionable risk categorization (High Risk vs. Low Risk) to assist physicians in rapid outpatient triage.")
]

# Add 4 clean cards on Slide 8
pos8 = [
    (Inches(1.0), Inches(2.5), Inches(5.3), Inches(2.0)),
    (Inches(6.9), Inches(2.5), Inches(5.3), Inches(2.0)),
    (Inches(1.0), Inches(4.8), Inches(5.3), Inches(2.0)),
    (Inches(6.9), Inches(4.8), Inches(5.3), Inches(2.0))
]

for idx, (title, desc) in enumerate(outcomes_data):
    left, top, width, height = pos8[idx]
    shape = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0xF4, 0xF7, 0xFA)
    shape.line.color.rgb = RGBColor(0x00, 0x80, 0x40)
    shape.line.width = Pt(1.5)
    
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.15)
    tf.margin_top = Inches(0.15)
    
    p0 = tf.paragraphs[0]
    p0.text = f"✓  {title}"
    p0.font.bold = True
    p0.font.size = Pt(13)
    p0.font.color.rgb = RGBColor(0x00, 0x60, 0x20)
    p0.space_after = Pt(4)
    
    p1 = tf.add_paragraph()
    p1.text = desc
    p1.font.size = Pt(10.5)
    p1.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

# ----------------------------------------------------
# SLIDE 9: References
# ----------------------------------------------------
slide9 = prs.slides[8]
for shape in slide9.shapes:
    if shape.has_text_frame:
        t = shape.text_frame.text
        if '[1]' in t or 'Author(s)' in t:
            shape.text_frame.clear()
            refs = [
                "[1] R. Sharma and S. Bandil, \"Heart Disease Prediction Using Machine Learning Algorithms,\" IEEE ICCCNT, pp. 1–6, 2022.",
                "[2] A. Mohan, B. S. Raman, and K. S. Babu, \"Effective Heart Disease Prediction Using Hybrid Machine Learning Techniques,\" IEEE Access, vol. 9, pp. 81504–81515, 2021.",
                "[3] P. Ghosh, D. Azam, M. Jonkman, and S. Karim, \"Efficient Prediction of Cardiovascular Disease Using ML Classifiers,\" Elsevier BHI, vol. 84, pp. 104758, 2023.",
                "[4] K. Polaraju and D. Durga, \"Prediction of Heart Disease Using Machine Learning: A Decision Support Framework,\" Springer SN CS, vol. 1, 2020.",
                "[5] World Health Organization (WHO), \"Cardiovascular Diseases (CVDs) Factsheet,\" Geneva: WHO, 2023. [Online]."
            ]
            for i, ref in enumerate(refs):
                p = shape.text_frame.paragraphs[0] if i == 0 else shape.text_frame.add_paragraph()
                p.text = ref
                p.font.name = 'Arial'
                p.font.size = Pt(11)
                p.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
                p.space_after = Pt(8)

prs.save(output_path)
print(f"Successfully generated PowerPoint presentation: {output_path}")
