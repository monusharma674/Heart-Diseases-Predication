import os
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="6" w:space="0" w:color="002060"/>
            <w:bottom w:val="single" w:sz="6" w:space="0" w:color="002060"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>
            <w:insideV w:val="none"/>
            <w:left w:val="none"/>
            <w:right w:val="none"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def build_synopsis():
    doc = docx.Document()

    # Configure Margins: Left: 2.5 cm, Top: 2.5 cm, Right: 1.25 cm, Bottom: 1.25 cm
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(1.25)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(1.25)
        section.page_width = Inches(8.27)  # A4
        section.page_height = Inches(11.69)

    # Styles
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Times New Roman'
    style_normal.font.size = Pt(11)
    style_normal.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    style_normal.paragraph_format.line_spacing = 1.5
    style_normal.paragraph_format.space_after = Pt(6)

    def add_p(text="", align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, italic=False, size=11, color_rgb=(0x22,0x22,0x22), space_before=0, space_after=6, line_spacing=1.5):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        if text:
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(size)
            run.bold = bold
            run.italic = italic
            run.font.color.rgb = RGBColor(*color_rgb)
        return p

    def add_heading_15(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(15)
        run.bold = True
        run.font.color.rgb = RGBColor(0x00, 0x20, 0x60)
        return p

    def add_subheading_13(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(13)
        run.bold = True
        run.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
        return p

    aktu_logo_path = 'extracted_assets/word/media/image1.png'
    lloyd_logo_path = 'extracted_assets/word/media/image2.png'

    # ==========================================
    # COVER PAGE / TITLE PAGE
    # ==========================================
    add_p("HEART DISEASE PREDICTION SYSTEM USING MACHINE LEARNING TECHNIQUES", 
          align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=15, color_rgb=(0x00, 0x20, 0x60), space_before=10, space_after=10)
    
    add_p("MAJOR PROJECT SYNOPSIS", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, color_rgb=(0x80, 0x00, 0x00), space_after=8)
    
    add_p("Submitted in partial fulfilment of the requirement for the award of the degree of", 
          align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, size=10.5, space_after=2)
    add_p("Bachelor of Technology", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=12.5, space_after=2)
    add_p("in", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, size=10.5, space_after=2)
    add_p("Computer Science & Engineering – AIML", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=12, color_rgb=(0x80, 0x00, 0x00), space_after=10)
    
    add_p("Submitted to", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, size=10.5, space_after=2)
    add_p("DR. A.P.J. ABDUL KALAM TECHNICAL UNIVERSITY, UTTAR PRADESH, LUCKNOW", 
          align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=11, color_rgb=(0x00, 0x20, 0x60), space_after=10)

    # AKTU Logo
    if os.path.exists(aktu_logo_path):
        p_logo1 = doc.add_paragraph()
        p_logo1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo1.paragraph_format.space_after = Pt(10)
        run_logo1 = p_logo1.add_run()
        run_logo1.add_picture(aktu_logo_path, width=Inches(1.6))

    # Submission Table (Submitted by | Under Supervision of)
    sub_table = doc.add_table(rows=1, cols=2)
    sub_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sub_table.autofit = False

    cell_left = sub_table.rows[0].cells[0]
    cell_right = sub_table.rows[0].cells[1]
    cell_left.width = Inches(3.6)
    cell_right.width = Inches(3.6)

    # Left Cell: Submitted by
    p_sub_hdr = cell_left.paragraphs[0]
    p_sub_hdr.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p_sub_hdr.add_run("Submitted by:")
    r.bold = True
    r.font.size = Pt(10.5)
    
    p_sub_body = cell_left.add_paragraph()
    p_sub_body.paragraph_format.line_spacing = 1.2
    p_sub_body.paragraph_format.space_after = Pt(2)
    r1 = p_sub_body.add_run("1. Monu Sharma\n   Roll No: [2301530100xxx]  Sec: B\n")
    r1.font.size = Pt(10)
    r2 = p_sub_body.add_run("2. [Student Name 2]\n   Roll No: [2301530100xxx]  Sec: B\n")
    r2.font.size = Pt(10)
    r3 = p_sub_body.add_run("3. [Student Name 3]\n   Roll No: [2301530100xxx]  Sec: B")
    r3.font.size = Pt(10)

    # Right Cell: Under Supervision of
    p_sup_hdr = cell_right.paragraphs[0]
    p_sup_hdr.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p_sup_hdr.add_run("Under the Supervision of:")
    r.bold = True
    r.font.size = Pt(10.5)

    p_sup_body = cell_right.add_paragraph()
    p_sup_body.paragraph_format.line_spacing = 1.2
    p_sup_body.paragraph_format.space_after = Pt(2)
    r_g1 = p_sup_body.add_run("[Name of Faculty Guide]\n")
    r_g1.bold = True
    r_g1.font.size = Pt(10.5)
    r_g2 = p_sup_body.add_run("Assistant Professor\nDepartment of Computer Science & Engineering")
    r_g2.font.size = Pt(10)

    # Lloyd Logo
    if os.path.exists(lloyd_logo_path):
        p_logo2 = doc.add_paragraph()
        p_logo2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo2.paragraph_format.space_before = Pt(8)
        p_logo2.paragraph_format.space_after = Pt(4)
        run_logo2 = p_logo2.add_run()
        run_logo2.add_picture(lloyd_logo_path, width=Inches(1.4))

    add_p("Department of Computer Science & Engineering", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=11, color_rgb=(0x00, 0x20, 0x60), space_after=2)
    add_p("Lloyd Institute of Engineering & Technology", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=12, color_rgb=(0x80, 0x00, 0x00), space_after=2)
    add_p("Plot No. 3, Knowledge Park II, Greater Noida, Uttar Pradesh 201306", align=WD_ALIGN_PARAGRAPH.CENTER, size=9.5, space_after=2)
    add_p("Session: 2026–27", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=10.5, space_after=0)

    doc.add_page_break()

    # ==========================================
    # INDEX PAGE
    # ==========================================
    add_heading_15("INDEX")
    
    index_table = doc.add_table(rows=8, cols=3)
    index_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    index_table.autofit = False
    set_table_borders(index_table)

    headers = ["Sr. No.", "Topic", "Page No."]
    col_widths = [Inches(1.0), Inches(4.5), Inches(1.5)]

    for i, h in enumerate(headers):
        cell = index_table.rows[0].cells[i]
        cell.width = col_widths[i]
        set_cell_background(cell, "002060")
        set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i != 1 else WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(h)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(10.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    index_data = [
        ("01", "Introduction", "1 – 2"),
        ("02", "Statement of Problem and Objectives", "3"),
        ("03", "Literature Survey", "4"),
        ("04", "Expected Outcome / Scope of the Project", "5"),
        ("05", "Tentative Work Plan (Mapped to 10-Week Timeline)", "6"),
        ("06", "Software / Hardware Required for Project Development", "7"),
        ("07", "References (IEEE Standard Format)", "8")
    ]

    for row_idx, data in enumerate(index_data, start=1):
        for col_idx, text in enumerate(data):
            cell = index_table.rows[row_idx].cells[col_idx]
            cell.width = col_widths[col_idx]
            if row_idx % 2 == 1:
                set_cell_background(cell, "F2F4F7")
            set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx != 1 else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(10)

    doc.add_paragraph().paragraph_format.space_after = Pt(50)

    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_table.autofit = False
    
    cell_s1 = sig_table.rows[0].cells[0]
    cell_s2 = sig_table.rows[0].cells[1]
    cell_s1.width = Inches(3.6)
    cell_s2.width = Inches(3.6)
    
    p1 = cell_s1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p1.add_run("_________________________\nStudent’s Signature").bold = True
    
    p2 = cell_s2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.add_run("_________________________\nGuide’s Signature").bold = True

    doc.add_page_break()

    # ==========================================
    # SECTION 1: INTRODUCTION
    # ==========================================
    add_heading_15("1. INTRODUCTION")
    
    p = add_p()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("Cardiovascular Diseases (CVDs) represent the foremost cause of mortality and physical morbidity worldwide. As documented by the World Health Organization (WHO), approximately 17.9 million individuals lose their lives to cardiovascular conditions every year, representing nearly 32% of all global deaths. Among these fatalities, over 85% are precipitated by acute myocardial infarction (heart attacks) and cerebrovascular strokes. In developing nations like India, demographic and epidemiological transitions, sedentary employment patterns, adverse dietary habits, smoking, and delayed clinical evaluations have substantially accelerated the incidence of cardiovascular pathologies among the working-age population.")

    p = add_p()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("Clinical diagnosis in traditional healthcare settings heavily depends upon specialized, multi-stage diagnostic investigations such as coronary angiography, 12-lead electrocardiography (ECG), stress echocardiography, and complex biomarker blood profiling. While these clinical diagnostic modalities provide authoritative evaluations, they present significant limitations: high operational costs, invasive intervention risks, prolonged diagnostic turnaround times, and severe shortage of trained cardiologists in rural and primary healthcare environments. To combat these challenges, deploying automated, intelligent, and non-invasive clinical decision-support systems powered by Machine Learning is imperative for rapid, early-stage cardiovascular risk stratification.")

    add_subheading_13("1.1 Project Domain & Applied Technologies")
    p = add_p()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("This project resides in the intersection of ")
    p.add_run("Applied Machine Learning, Predictive Data Analytics, and Healthcare Informatics").bold = True
    p.add_run(". By utilizing structured multi-parametric patient records—encompassing demographic attributes (Age, Sex), hemodynamic indicators (Resting Blood Pressure, Maximum Heart Rate), metabolic factors (Serum Cholesterol, Fasting Blood Sugar), and cardiac stress markers (Chest Pain Type, Exercise-Induced Angina, ST-Segment Depression, ST Slope)—machine learning classification models can uncover complex multivariate patterns and generate real-time probabilistic risk assessments.")

    add_subheading_13("1.2 Important Technical Terms & Concepts")
    terms = [
        ("K-Nearest Neighbors (KNN): ", "A robust, non-parametric, instance-based supervised classification algorithm that assigns risk labels to unseen patient records based on the plurality vote of the 'k' closest historical clinical instances measured across multidimensional Euclidean space."),
        ("Feature Scaling (StandardScaler): ", "A critical statistical standardization technique that transforms continuous features (Cholesterol, Resting BP, Heart Rate) to exhibit zero mean and unit variance (Z-score normalization), preventing features with larger numerical magnitudes from distorting distance calculations."),
        ("One-Hot Encoding & Schema Preservation: ", "Systematic transformation of multi-class categorical medical variables (e.g., Chest Pain Types [ATA, NAP, TA, ASY], Resting ECG, ST Slope) into binary numerical vectors while strictly aligning input feature columns with training schema."),
        ("Joblib Artifact Serialization: ", "High-performance object serialization used to store trained ML models, pre-fitted scaler parameters, and column schema matrices into binary files (.pkl) for zero-latency execution in production."),
        ("Streamlit Web Framework: ", "A state-of-the-art Python framework leveraged to engineer a responsive, interactive, and clinician-friendly user interface for seamless parameter entry, real-time model inference, and clear visual risk alerts.")
    ]

    for title, desc in terms:
        p = add_p(space_after=4)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_t = p.add_run("• " + title)
        r_t.bold = True
        r_t.font.color.rgb = RGBColor(0x00, 0x20, 0x60)
        p.add_run(desc)

    doc.add_page_break()

    # ==========================================
    # SECTION 2: STATEMENT OF PROBLEM & OBJECTIVES
    # ==========================================
    add_heading_15("2. STATEMENT OF PROBLEM AND OBJECTIVES")
    
    add_subheading_13("2.1 Problem Statement")
    p = add_p()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("The accurate, early-stage detection of cardiovascular diseases poses a critical healthcare challenge. Traditional diagnostic protocols are predominantly reactive, cost-prohibitive, and dependent on extensive laboratory infrastructure. In high-density outpatient environments, physicians encounter heavy clinical workloads, where manual synthesis of multi-dimensional patient risk markers increases vulnerability to diagnostic fatigue and delayed interventions. Furthermore, remote and rural clinics lack specialist cardiologists and specialized equipment. Consequently, there is an urgent demand for a non-invasive, cost-effective, and automated machine learning-driven risk assessment system capable of delivering instantaneous and reliable heart disease evaluations.")

    add_subheading_13("2.2 Limitations of Existing Diagnostic Approaches")
    limitations = [
        ("Invasive & Costly Procedures: ", "Angiography and CT scans are expensive, physically invasive, and unavailable in primary rural healthcare dispensaries."),
        ("Diagnostic Latency: ", "Substantial time delays between blood test sample collection, laboratory processing, and physician review can compromise acute patient care."),
        ("Subjective Clinical Interpretation: ", "Inter-observer variability in manual symptom grading and ECG interpretation can result in inconsistent initial risk stratification."),
        ("Absence of Integrated Triage Tools: ", "Existing clinical software solutions are complex, proprietary, and lack lightweight, accessible web interfaces for rapid preliminary screening.")
    ]
    for title, desc in limitations:
        p = add_p(space_after=4)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_t = p.add_run("• " + title)
        r_t.bold = True
        p.add_run(desc)

    add_subheading_13("2.3 Project Objectives")
    p = add_p()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("The core objectives of the proposed Major Project are established as follows:")

    objectives = [
        ("Design and implement an automated Machine Learning classification pipeline ", "specifically optimized for binary cardiovascular disease risk assessment (High Risk vs. Low Risk) based on standardized clinical indicators."),
        ("Develop a rigorous data preprocessing and feature scaling module ", "incorporating StandardScaler normalization and dynamic one-hot categorical alignment to eliminate scale bias and maintain schema integrity."),
        ("Train, optimize, and benchmark the K-Nearest Neighbors (KNN) classifier ", "on standard multi-hospital cardiac datasets to maximize prediction accuracy, sensitivity, and clinical reliability."),
        ("Construct an interactive, web-based graphical user interface (GUI) using Streamlit ", "enabling clinicians, paramedics, and users to input clinical parameters effortlessly and receive instant, color-coded visual risk alerts."),
        ("Package and serialize pre-trained model artifacts (Joblib / PKL) ", "to guarantee lightweight, portable, and low-latency deployment suitable for local clinic workstations and cloud deployment.")
    ]

    for idx, (head, tail) in enumerate(objectives, start=1):
        p = add_p(space_after=4)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_num = p.add_run(f"Objective {idx}: ")
        r_num.bold = True
        r_num.font.color.rgb = RGBColor(0x00, 0x20, 0x60)
        p.add_run(head)
        p.add_run(tail)

    doc.add_page_break()

    # ==========================================
    # SECTION 3: LITERATURE SURVEY
    # ==========================================
    add_heading_15("3. LITERATURE SURVEY")
    
    p = add_p()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("Extensive academic literature has explored the integration of computational intelligence and supervised machine learning algorithms in cardiovascular diagnostics. A critical review of prominent publications is summarized below:")

    lit_table = doc.add_table(rows=5, cols=5)
    lit_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    lit_table.autofit = False
    set_table_borders(lit_table)

    lit_headers = ["S.No", "Paper Title", "Author(s)", "Journal / Year", "Key Findings & Limitations"]
    lit_widths = [Inches(0.6), Inches(1.8), Inches(1.4), Inches(1.2), Inches(2.2)]

    for i, h in enumerate(lit_headers):
        cell = lit_table.rows[0].cells[i]
        cell.width = lit_widths[i]
        set_cell_background(cell, "002060")
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    lit_data = [
        ("1", "Heart Disease Prediction Using Machine Learning Algorithms", "R. Sharma & S. Bandil", "IEEE Xplore (2022)", "Evaluated KNN, Naive Bayes, and SVM on UCI dataset. KNN achieved 86.8% accuracy; noted high dependency on proper feature scaling."),
        ("2", "Predictive Analytics for Cardiovascular Disease Risk", "A. Mohan et al.", "IEEE Access (2021)", "Explored hybrid ML architectures; demonstrated that ST-slope, chest pain type, and max heart rate are the highest-gain predictors."),
        ("3", "Comparative Study of ML Classifiers in Heart Disease", "P. Ghosh & D. Azam", "Elsevier BHI (2023)", "Benchmarked distance-based classifiers; demonstrated that StandardScaler normalization significantly reduces false-negative diagnostic rates."),
        ("4", "Intelligent Decision Support System for Heart Disease", "K. Polaraju & D. Durga", "Springer SN CS (2020)", "Developed an automated risk prediction system; advocated for lightweight serialization and standalone web GUI for practical clinic deployment.")
    ]

    for row_idx, data in enumerate(lit_data, start=1):
        for col_idx, text in enumerate(data):
            cell = lit_table.rows[row_idx].cells[col_idx]
            cell.width = lit_widths[col_idx]
            if row_idx % 2 == 1:
                set_cell_background(cell, "F9FAFC")
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx in [0, 3] else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(9)

    add_subheading_13("3.1 Research Gap & Proposed Innovation")
    p = add_p()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("Although existing academic studies confirm the theoretical efficacy of classification algorithms, existing systems often neglect real-world deployment challenges: rigid input schema enforcement, graceful handling of unobserved categorical combinations during live inference, and streamlined, low-latency web interfaces. The proposed system directly resolves these deficiencies by uniting an optimized K-Nearest Neighbors classifier with an automated one-hot schema reconstructor and a lightweight Streamlit web application.")

    doc.add_page_break()

    # ==========================================
    # SECTION 4: EXPECTED OUTCOME / SCOPE
    # ==========================================
    add_heading_15("4. EXPECTED OUTCOME AND SCOPE OF THE PROJECT")
    
    add_subheading_13("4.1 Expected Outcomes")
    outcomes = [
        ("High-Precision Risk Stratification: ", "A fully validated K-Nearest Neighbors ML model delivering robust classification accuracy (>86%) in categorizing patient cardiovascular risk."),
        ("Interactive Web Application: ", "A clean, responsive Streamlit dashboard featuring intuitive sliders, numeric inputs, and selection menus for frictionless patient data input."),
        ("Zero-Latency Real-Time Inference: ", "Instantaneous risk evaluation (High Risk / Low Risk) powered by serialized (.pkl) pipeline artifacts with zero re-training latency."),
        ("Clinical Decision-Support UI: ", "Clear visual feedback with contextual warnings (Red for High Risk, Green for Low Risk) to assist preliminary triage.")
    ]
    for title, desc in outcomes:
        p = add_p(space_after=4)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_t = p.add_run("✓ " + title)
        r_t.bold = True
        r_t.font.color.rgb = RGBColor(0x00, 0x60, 0x20)
        p.add_run(desc)

    add_subheading_13("4.2 Key Features & Clinical Parameters Ingested")
    p = add_p()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("The application evaluates 11 critical clinical indicators: Age (years), Sex (M/F), Chest Pain Type (ATA: Atypical Angina, NAP: Non-Anginal Pain, TA: Typical Angina, ASY: Asymptomatic), Resting Blood Pressure (mm Hg), Serum Cholesterol (mg/dL), Fasting Blood Sugar (>120 mg/dL), Resting ECG (Normal, ST, LVH), Maximum Heart Rate (bpm), Exercise-Induced Angina (Y/N), Oldpeak (ST depression in mm), and ST Slope (Up, Flat, Down).")

    add_subheading_13("4.3 Intended Beneficiaries")
    p = add_p()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("• ")
    p.add_run("Primary Healthcare Centers & Rural Dispensaries: ").bold = True
    p.add_run("Empowers healthcare staff to perform initial risk stratification before referring patients to tertiary cardiologists.\n• ")
    p.add_run("General Medical Practitioners: ").bold = True
    p.add_run("Functions as an instant computational second opinion during routine checkups.\n• ")
    p.add_run("General Public / Patients: ").bold = True
    p.add_run("Facilitates proactive self-screening to encourage early clinical consultations and lifestyle modifications.")

    add_subheading_13("4.4 Scope for Future Enhancement")
    p = add_p()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("1. ")
    p.add_run("Explainable AI (XAI) Integration: ").bold = True
    p.add_run("Incorporation of SHAP / LIME explainability frameworks to highlight the exact mathematical contribution of each patient feature to the final risk score.\n2. ")
    p.add_run("IoT Wearable Sensor Sync: ").bold = True
    p.add_run("Direct synchronization with consumer smartwatch telemetry (real-time PPG heart rate, single-lead ECG data, SpO2) for dynamic risk tracking.\n3. ")
    p.add_run("EHR & Cloud Integration: ").bold = True
    p.add_run("Integration with hospital Electronic Health Records (FHIR standard) and cloud deployment (AWS/GCP) for hospital-wide accessibility.")

    doc.add_page_break()

    # ==========================================
    # SECTION 5: TENTATIVE WORK PLAN
    # ==========================================
    add_heading_15("5. TENTATIVE WORK PLAN (10-WEEK TIMELINE)")
    
    p = add_p()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("The development lifecycle of the Major Project is planned across a 10-week systematic timeline mapped to academic review milestones:")

    plan_table = doc.add_table(rows=11, cols=4)
    plan_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    plan_table.autofit = False
    set_table_borders(plan_table)

    plan_headers = ["Week", "Activity / Major Task", "Deliverable / Outcome", "Status"]
    plan_widths = [Inches(0.9), Inches(2.8), Inches(2.5), Inches(1.0)]

    for i, h in enumerate(plan_headers):
        cell = plan_table.rows[0].cells[i]
        cell.width = plan_widths[i]
        set_cell_background(cell, "002060")
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    plan_data = [
        ("Week 1", "Problem Identification & Domain Research", "Project Scope & Feasibility Report", "Completed"),
        ("Week 2", "Literature Review & Dataset Acquisition", "Literature Survey & UCI Dataset Curation", "Completed"),
        ("Week 3", "Data Cleaning, Imputation & Exploratory Data Analysis", "EDA Visualizations & Clean Dataset", "Completed"),
        ("Week 4", "Feature Engineering, Categorical Encoding & Scaling", "Preprocessed Matrix & Scaler Object", "Completed"),
        ("Week 5", "Model Selection, KNN Architecture & Parameter Tuning", "Baseline Trained KNN Classifier", "Completed"),
        ("Week 6", "Model Evaluation (Cross-Validation, Confusion Matrix)", "Accuracy & Performance Metric Benchmark", "Completed"),
        ("Week 7", "Model Serialization & Streamlit Web UI Development", "Interactive Streamlit Web Dashboard", "Completed"),
        ("Week 8", "System Integration, Exception Handling & Edge Testing", "Integrated Production Application", "In Progress"),
        ("Week 9", "Synopsis & Major Project Documentation Preparation", "Draft Synopsis & Presentation Deck", "In Progress"),
        ("Week 10", "Final Demonstration, Review-01 Presentation Defense", "Final Presentation & Code Repository", "Planned")
    ]

    for row_idx, data in enumerate(plan_data, start=1):
        for col_idx, text in enumerate(data):
            cell = plan_table.rows[row_idx].cells[col_idx]
            cell.width = plan_widths[col_idx]
            if row_idx % 2 == 1:
                set_cell_background(cell, "F9FAFC")
            set_cell_margins(cell, top=70, bottom=70, left=100, right=100)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx in [0, 3] else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(9)
            if col_idx == 3:
                if text == "Completed":
                    run.font.color.rgb = RGBColor(0x00, 0x80, 0x00)
                    run.bold = True
                elif text == "In Progress":
                    run.font.color.rgb = RGBColor(0xB8, 0x60, 0x00)
                    run.bold = True

    doc.add_page_break()

    # ==========================================
    # SECTION 6: HARDWARE & SOFTWARE REQUIREMENTS
    # ==========================================
    add_heading_15("6. SOFTWARE AND HARDWARE REQUIREMENTS")
    
    add_subheading_13("6.1 Software Requirements")
    
    sw_table = doc.add_table(rows=7, cols=3)
    sw_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sw_table.autofit = False
    set_table_borders(sw_table)

    sw_headers = ["Category", "Software / Tool Name", "Version & Purpose"]
    sw_widths = [Inches(1.8), Inches(2.2), Inches(3.2)]

    for i, h in enumerate(sw_headers):
        cell = sw_table.rows[0].cells[i]
        cell.width = sw_widths[i]
        set_cell_background(cell, "002060")
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    sw_data = [
        ("Operating System", "Microsoft Windows 10 / 11 / Linux (Ubuntu)", "Primary development and deployment OS environment"),
        ("Programming Language", "Python", "Version 3.10+ (Core language for ML and backend scripting)"),
        ("Machine Learning", "Scikit-Learn, Joblib", "Algorithm implementation (KNN), StandardScaler, serialization"),
        ("Data Manipulation", "Pandas, NumPy", "Dataset ingestion, cleaning, one-hot encoding, array handling"),
        ("Web Framework / UI", "Streamlit", "Version 1.30+ (Interactive medical web interface)"),
        ("Development Tools", "VS Code / Jupyter Notebook / Git", "Source code editing, prototyping, and GitHub version control")
    ]

    for row_idx, data in enumerate(sw_data, start=1):
        for col_idx, text in enumerate(data):
            cell = sw_table.rows[row_idx].cells[col_idx]
            cell.width = sw_widths[col_idx]
            if row_idx % 2 == 1:
                set_cell_background(cell, "F9FAFC")
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(9)

    add_subheading_13("6.2 Hardware Requirements")
    
    hw_table = doc.add_table(rows=6, cols=3)
    hw_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hw_table.autofit = False
    set_table_borders(hw_table)

    hw_headers = ["Component", "Minimum Specification", "Recommended Specification"]
    hw_widths = [Inches(2.0), Inches(2.5), Inches(2.7)]

    for i, h in enumerate(hw_headers):
        cell = hw_table.rows[0].cells[i]
        cell.width = hw_widths[i]
        set_cell_background(cell, "002060")
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    hw_data = [
        ("Processor (CPU)", "Intel Core i3 / AMD Ryzen 3 (2.4 GHz)", "Intel Core i5 / i7 or AMD Ryzen 5+ (3.0 GHz+)"),
        ("System Memory (RAM)", "4 GB DDR4", "8 GB / 16 GB DDR4/DDR5"),
        ("Storage Drive", "128 GB HDD", "256 GB / 512 GB NVMe SSD"),
        ("Display Monitor", "1366 x 768 Resolution", "1920 x 1080 (Full HD IPS Display)"),
        ("Network Interface", "Standard Wi-Fi / LAN Card", "High-Speed Broadband Internet Connection")
    ]

    for row_idx, data in enumerate(hw_data, start=1):
        for col_idx, text in enumerate(data):
            cell = hw_table.rows[row_idx].cells[col_idx]
            cell.width = hw_widths[col_idx]
            if row_idx % 2 == 1:
                set_cell_background(cell, "F9FAFC")
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(9)

    doc.add_page_break()

    # ==========================================
    # SECTION 7: REFERENCES
    # ==========================================
    add_heading_15("7. REFERENCES")
    
    p = add_p()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("All scholarly publications, journals, and authoritative technical resources referenced during the conceptualization and development of this project are cited below according to the IEEE standard format:")

    references = [
        "[1] R. Sharma and S. Bandil, \"Heart Disease Prediction Using Machine Learning Algorithms and Feature Selection Techniques,\" in Proceedings of the IEEE International Conference on Computing, Communication and Networking Technologies (ICCCNT), pp. 1–6, 2022.",
        "[2] A. Mohan, B. S. Raman, and K. S. Babu, \"Effective Heart Disease Prediction Using Hybrid Machine Learning Techniques,\" IEEE Access, vol. 9, pp. 81504–81515, 2021.",
        "[3] P. Ghosh, D. Azam, M. Jonkman, and S. Karim, \"Efficient Prediction of Cardiovascular Disease Using Machine Learning Classifiers with Distance Metric Tuning,\" Elsevier Biomedical Signal Processing and Control, vol. 84, pp. 104758, 2023.",
        "[4] K. Polaraju and D. Durga, \"Prediction of Heart Disease Using Machine Learning Techniques: A Decision Support Framework,\" Springer SN Computer Science, vol. 1, no. 4, pp. 210–220, 2020.",
        "[5] World Health Organization (WHO), \"Cardiovascular Diseases (CVDs) Factsheet,\" Geneva, Switzerland: World Health Organization, 2023. [Online]. Available: https://www.who.int/news-room/fact-sheets/detail/cardiovascular-diseases-(cvds)",
        "[6] F. Pedregosa et al., \"Scikit-learn: Machine Learning in Python,\" Journal of Machine Learning Research (JMLR), vol. 12, pp. 2825–2830, 2011.",
        "[7] A. Trevisani, Streamlit: The Comprehensive Guide to Rapid Data Web Application Development in Python, New York: O'Reilly Media, 2023.",
        "[8] UC Irvine Machine Learning Repository, \"Heart Disease Dataset,\" Center for Machine Learning and Intelligent Systems, University of California, Irvine, 2020. [Online]. Available: https://archive.uci.edu/ml/datasets/heart+disease"
    ]

    for ref in references:
        p = add_p(space_after=8)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run = p.add_run(ref)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(10)

    output_path = r'C:\Users\MONU SHARMA\OneDrive\Desktop\heartproject\Heart_Disease_Prediction_Synopsis.docx'
    doc.save(output_path)
    print(f'Successfully generated Synopsis document: {output_path}')

if __name__ == '__main__':
    build_synopsis()
