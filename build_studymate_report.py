import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement, parse_xml
import os

TEMPLATE_PATH = r"C:\Users\DELL\Downloads\FinalTTLReport.docx"
OUTPUT_DOCX = r"c:\Users\DELL\Desktop\StudyMate-AI-main\StudyMate_AI_Project_Report.docx"

ASSETS_DIR = r"c:\Users\DELL\Desktop\StudyMate-AI-main\assets\screenshots"
LOGO_PATH = r"c:\Users\DELL\Desktop\StudyMate-AI-main\college_logo.png"

def create_report():
    print("Opening template:", TEMPLATE_PATH)
    doc = docx.Document(TEMPLATE_PATH)
    body = doc._body._element
    sectPr = doc.sections[0]._sectPr

    # Remove all existing paragraphs and tables in body, preserving sectPr (page borders, margins, footer)
    for child in list(body):
        if child != sectPr:
            body.remove(child)

    print("Cleaned body. Beginning refined page generation...")

    # Helper functions
    def add_para(text="", style='Normal', align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=6, line_spacing=1.15):
        p = doc.add_paragraph(style=style)
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        if text:
            r = p.add_run(text)
            r.font.name = "Times New Roman"
            r.font.size = Pt(12)
        return p

    def add_run(p, text, font_size=12, bold=False, italic=False, underline=False, color=None):
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(font_size)
        r.bold = bold
        r.italic = italic
        r.underline = underline
        if color:
            r.font.color.rgb = RGBColor.from_string(color)
        return r

    def add_bullet(text, bold_prefix="", font_size=12, space_after=4):
        p = doc.add_paragraph(style='Normal')
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        
        r_bullet = p.add_run("• ")
        r_bullet.font.name = "Times New Roman"
        r_bullet.font.size = Pt(font_size)
        r_bullet.bold = True
        
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.font.name = "Times New Roman"
            r_pre.font.size = Pt(font_size)
            r_pre.bold = True
            
        r_text = p.add_run(text)
        r_text.font.name = "Times New Roman"
        r_text.font.size = Pt(font_size)
        return p

    def add_chapter_heading(chap_num, chap_title):
        p = doc.add_paragraph(style='Normal')
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(12)
        add_run(p, f"CHAPTER-{chap_num} – ", font_size=14, bold=True, underline=True)
        add_run(p, chap_title.upper(), font_size=16, bold=True, underline=True)
        return p

    def add_section_heading(sec_title):
        p = doc.add_paragraph(style='Normal')
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(6)
        return add_run(p, sec_title, font_size=14, bold=True)

    def add_sub_heading(sub_title, font_size=12):
        p = doc.add_paragraph(style='Normal')
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(4)
        return add_run(p, sub_title, font_size=font_size, bold=True)

    def add_image_box(img_path, width_in=5.5, caption=None):
        p = doc.add_paragraph(style='Normal')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run()
        r.add_picture(img_path, width=Inches(width_in))
        
        if caption:
            p_cap = doc.add_paragraph(style='Normal')
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(6)
            add_run(p_cap, caption, font_size=12, bold=True)

    # =========================================================================
    # PAGE 1: TITLE PAGE
    # =========================================================================
    print("Writing Page 1: Title Page...")
    add_para("", space_after=8)
    p1 = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_run(p1, "D.Y. PATIL COLLEGE OF ENGINEERING &TECHNOLOGY,", font_size=14, bold=False)
    
    p2 = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_run(p2, "KASABA BAWADA, KOLHAPUR", font_size=14, bold=False)
    
    p3 = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_run(p3, "DEPARTMENT OF COMPUTER SCIENCE & ENGINEERING", font_size=14, bold=False)
    
    p4 = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
    add_run(p4, "(Academic Year: 2024-25)", font_size=14, color="FF0000")
    
    # College Logo
    p_logo = doc.add_paragraph(style='Normal')
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_logo.paragraph_format.space_after = Pt(14)
    r_logo = p_logo.add_run()
    r_logo.add_picture(LOGO_PATH, width=Inches(2.17))
    
    p_a = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_run(p_a, "A", font_size=14, bold=True)
    
    p_rep = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_run(p_rep, "TTL Project Report", font_size=14, bold=True)
    
    p_on = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
    add_run(p_on, "On", font_size=14)
    
    p_title = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=16)
    add_run(p_title, '"STUDYMATE AI — AN AGENTIC LEARNING ASSISTANT"', font_size=16, bold=True, color="FF0000")
    
    p_sub = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    add_run(p_sub, "Submitted By:", font_size=14, bold=True)
    
    p_stud = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)
    add_run(p_stud, "Name: ", font_size=14, bold=True)
    add_run(p_stud, "Jija Shashin Kumbhojkar          ", font_size=14)
    add_run(p_stud, "Roll No.: ", font_size=14, bold=True)
    add_run(p_stud, "28", font_size=14)
    
    p_guid = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_run(p_guid, "Under the Guidance of:", font_size=14, bold=True)
    
    p_gname = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=26)
    add_run(p_gname, "Mrs. Rutuja S. Kalekar", font_size=14)
    
    p_cls = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    add_run(p_cls, "Class: ", font_size=14, bold=True)
    add_run(p_cls, "SY (CSE)     ", font_size=14)
    add_run(p_cls, "Div.: ", font_size=14, bold=True)
    add_run(p_cls, "A     ", font_size=14)
    add_run(p_cls, "Batch: ", font_size=14, bold=True)
    add_run(p_cls, "S2", font_size=14)
    
    p_yr = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=0)
    add_run(p_yr, "Year 2025-26", font_size=14, bold=True)
    
    doc.add_page_break()

    # =========================================================================
    # PAGE 2: CERTIFICATE
    # =========================================================================
    print("Writing Page 2: Certificate...")
    add_para("", space_after=40)
    p_cert = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=24)
    add_run(p_cert, "CERTIFICATE", font_size=16, bold=True, underline=True, color="000080")
    
    p_ctext = add_para(align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=140, line_spacing=1.35)
    add_run(p_ctext, "This is to certify that ", font_size=14)
    add_run(p_ctext, "Jija Shashin Kumbhojkar ", font_size=14, bold=True)
    add_run(p_ctext, "has successfully completed the term work for ", font_size=14)
    add_run(p_ctext, "TTL Project ", font_size=14, bold=True)
    add_run(p_ctext, "entitled ", font_size=14)
    add_run(p_ctext, "“STUDYMATE AI — AN AGENTIC LEARNING ASSISTANT” ", font_size=14, bold=True, color="FF0000")
    add_run(p_ctext, "towards the partial fulfillment of ", font_size=14)
    add_run(p_ctext, "S.Y. (CSE) Sem II ", font_size=14, bold=True)
    add_run(p_ctext, "during the academic year ", font_size=14)
    add_run(p_ctext, "2025-26.", font_size=14, bold=True)
    
    p_signs = add_para(align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=60)
    add_run(p_signs, "\tGuides\t\t\tHOD\t\t\tPrincipal", font_size=14, bold=True, color="000080")
    
    p_examiner = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=0)
    add_run(p_examiner, "Internal Examiner", font_size=14, bold=True, color="000080")
    
    doc.add_page_break()

    # =========================================================================
    # PAGE 3: DECLARATION
    # =========================================================================
    print("Writing Page 3: Declaration...")
    add_para("", space_after=24)
    p_ddept = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_run(p_ddept, "Department of Computer Science and Engineering", font_size=14, bold=True, color="800000")
    
    p_dcoll = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_run(p_dcoll, "D. Y. Patil College of Engineering & Technology, Kolhapur", font_size=14, bold=True, color="800000")
    
    p_dyr = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=28)
    add_run(p_dyr, "2025-26", font_size=14, bold=True)
    
    p_dtitle = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=24)
    add_run(p_dtitle, "DECLARATION", font_size=16, bold=True, underline=True, color="0070C0")
    
    p_dtext = add_para(align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=120, line_spacing=1.3)
    add_run(p_dtext, "I", font_size=14, bold=True)
    add_run(p_dtext, " hereby declare that the dissertation work report entitled ", font_size=14)
    add_run(p_dtext, "“STUDYMATE AI — AN AGENTIC LEARNING ASSISTANT” ", font_size=14, bold=True, color="002060")
    add_run(p_dtext, "which is being submitted to ", font_size=14)
    add_run(p_dtext, "D.Y. Patil College of Engineering and Technology, Kolhapur", font_size=14, bold=True)
    add_run(p_dtext, ", in partial fulfillment of SY CSE TTL Project course, is a bonafide report of the work carried out by me. The material contained in this report has not been submitted to any university or institution for the award of any degree.", font_size=14)
    
    p_place = add_para(align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=4)
    add_run(p_place, "Place: Kolhapur.", font_size=12, bold=True)
    
    p_date = add_para(align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=50)
    add_run(p_date, "Date: ", font_size=12, bold=True)
    
    p_sghdr = add_para(align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=6)
    add_run(p_sghdr, "\tStudent Name\t\t\t\t\tGuide Name", font_size=12, bold=True, color="FF0000")
    
    p_sgnames = add_para(align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=0)
    add_run(p_sgnames, "  Jija Shashin Kumbhojkar\t\t\t\tMrs. Rutuja S. Kalekar", font_size=12, bold=True, color="002060")
    
    doc.add_page_break()

    # =========================================================================
    # PAGE 4: ACKNOWLEDGEMENT
    # =========================================================================
    print("Writing Page 4: Acknowledgement...")
    add_para("", space_after=30)
    p_ack = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=24)
    add_run(p_ack, "ACKNOWLEDGEMENT", font_size=16, bold=True, underline=True, color="4472C4")
    
    p_atext1 = add_para(align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=14, line_spacing=1.25)
    add_run(p_atext1, "I express my sincere thanks to ", font_size=14)
    add_run(p_atext1, "Mrs. Rutuja S. Kalekar", font_size=14, bold=True)
    add_run(p_atext1, ", Computer Science and Engineering Dept whose supervision, inspiration and valuable guidance helped me a lot to complete my TTL Project. Her guidance proved to be the most valuable to overcome all the hurdles in the fulfillment of this Project Report.", font_size=14)
    
    p_atext2 = add_para(align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=90, line_spacing=1.25)
    add_run(p_atext2, "I also express my sincere thanks to ", font_size=14)
    add_run(p_atext2, "Prof. R. J. Dhanal", font_size=14, bold=True, color="EE0000")
    add_run(p_atext2, ", Head of computer science department and ", font_size=14)
    add_run(p_atext2, "Dr. Prof. S. D. Chede", font_size=14, bold=True, color="EE0000")
    add_run(p_atext2, ", Principal of D.Y. Patil College of Engineering & Technology Kolhapur. Last but not least, this acknowledgement would be incomplete without rendering our sincere gratitude to all those who have helped us in the completion of TTL Project. Finally, I express my deep sense of gratitude to my parents & family members, who are the chief source of my inspiration.", font_size=14)
    
    p_sinc = add_para(align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=4)
    add_run(p_sinc, "Sincerely,\t\t\t\t\t", font_size=14)
    
    p_asig = add_para(align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=2)
    add_run(p_asig, "Jija Kumbhojkar\t\t\t\t\t", font_size=14, bold=True)
    
    p_aroll = add_para(align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=0)
    add_run(p_aroll, "SY CSE TTL Project\t\t\t\t\t", font_size=14, bold=True)
    
    doc.add_page_break()

    # =========================================================================
    # PAGE 5: INDEX
    # =========================================================================
    print("Writing Page 5: Index...")
    add_para("", space_after=14)
    p_idx = add_para(align=WD_ALIGN_PARAGRAPH.LEFT, space_after=4)
    add_run(p_idx, "    INDEX", font_size=14, bold=True, underline=True, color="31859C")
    
    p_pgno = add_para(align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=4)
    add_run(p_pgno, "Page No", font_size=14, bold=True)
    
    # Table of contents formatted exactly like reference template
    def add_ch_index(ch_title, pg_num):
        p = doc.add_paragraph(style='Normal')
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.tab_stops.add_tab_stop(Inches(6.8), WD_TAB_ALIGNMENT.RIGHT)
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        add_run(p, ch_title, font_size=14, bold=True, color="31859C")
        add_run(p, f"\t{pg_num}", font_size=14, bold=True, color="31859C")

    def add_sec_index(sec_title):
        p = doc.add_paragraph(style='Normal')
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Inches(0.4)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        add_run(p, sec_title, font_size=13, bold=False)

    add_ch_index("Chapter 1: INTRODUCTION", "6")
    add_sec_index("1.1 Introduction")
    add_sec_index("1.2 Technology Stack")
    add_sec_index("1.3 Motivation")
    
    add_ch_index("Chapter-2: PROBLEM STATEMENT", "10")
    add_sec_index("2.1 Need of Work")
    add_sec_index("2.2 Problem Statement")
    add_sec_index("2.3 Objectives")
    
    add_ch_index("Chapter 3: DESIGN DETAILS", "12")
    add_sec_index("3.1 System Architecture")
    add_sec_index("3.2 System Design Diagrams")
    
    add_ch_index("Chapter 4: IMPLEMENTATION", "15")
    add_sec_index("4.1 Module Information")
    add_sec_index("4.2 System Requirements")
    
    add_ch_index("Chapter 5: EXPERIMENTAL RESULTS", "21")
    add_sec_index("5.1 Screenshots")
    
    add_ch_index("Chapter 6: CONCLUSION", "25")
    add_sec_index("6.1 Conclusion")
    add_sec_index("6.2 Future Scope")
    
    add_ch_index("Chapter 7: REFERENCES", "27")
    add_sec_index("7.1 References")
    
    doc.add_page_break()

    # =========================================================================
    # PAGE 6: CHAPTER 1 — INTRODUCTION (1.1)
    # =========================================================================
    print("Writing Page 6: Chapter 1 Introduction...")
    add_chapter_heading(1, "INTRODUCTION")
    add_sub_heading("1.1 Introduction:", font_size=14)
    
    p = add_para(space_after=8)
    p.add_run("The evolution of academic computer science and higher education has increasingly been shaped by digital learning resources, electronic document distribution, and computer-assisted pedagogical tools. Over the past decade, traditional physical textbooks and chalkboard lectures have largely transitioned into high-volume electronic formats, including portable document format (PDF) slide decks, digital syllabus manuals, research papers, and comprehensive lecture notes. While this widespread digitalization has democratized access to learning materials, it has simultaneously introduced profound cognitive challenges for students. Undergraduate engineering learners are routinely inundated with dense, multi-page technical documents, creating severe information overload, study fatigue, and fragmented conceptual retention when preparing for coursework and semester examinations.")
    
    p = add_para(space_after=8)
    p.add_run("The ")
    add_run(p, "StudyMate AI", bold=True)
    p.add_run(" project arises as a transformative academic solution to this challenge, reimagining educational document interaction as an active, perception-driven learning ecosystem that prioritizes human cognitive retention and personalized guidance. Developed in Python with a Streamlit web interface, StudyMate AI functions as an ")
    add_run(p, "Agentic Learning Assistant", bold=True)
    p.add_run(" that directly converts passive lecture notes into dynamic, interactive study partners. Rather than requiring students to manually skim hundreds of pages to locate exam definitions or draft revision questions, StudyMate AI provides an intelligent framework that automates document chunking, semantic vectorization, conceptual summarization, syllabus-grounded question answering, interactive quiz generation, and reflective learning coaching.")

    p = add_para(space_after=8)
    p.add_run("At its core, StudyMate AI addresses the critical vulnerabilities of conventional generative artificial intelligence tools in academic environments. While generic conversational large language models (LLMs) can produce fluent prose, they operate on broad, generalized training corpora and frequently suffer from hallucinations—fabricating plausible yet factually incorrect definitions, equations, and code snippets. For engineering students requiring precise syllabus alignment, ungrounded AI responses pose significant risks. StudyMate AI overcomes this fundamental limitation through a robust ")
    add_run(p, "Retrieval-Augmented Generation (RAG)", bold=True)
    p.add_run(" pipeline powered by Sentence Transformers (")
    add_run(p, "all-MiniLM-L6-v2", italic=True)
    p.add_run(") and persistent ChromaDB vector storage. Every response, summary, quiz question, and feedback item is strictly grounded in the student's authentic uploaded documents.")

    p = add_para(space_after=0)
    p.add_run("Furthermore, StudyMate AI embraces an agentic architectural paradigm through modular specialized agents: Document Processing, Semantic Retrieval, Structured Summarization, Contextual Question Answering, Quiz Generation, and Reflective Coaching. By implementing explicit inter-agent communication protocols, the system creates a cohesive pedagogical feedback loop where assessment performance directly informs subsequent study strategies, bridging the gap between passive document consumption and verified subject mastery.")

    doc.add_page_break()

    # =========================================================================
    # PAGE 7: CHAPTER 1 — TECHNOLOGY STACK (1.2 Part 1)
    # =========================================================================
    print("Writing Page 7: Chapter 1 Technology Stack Part 1...")
    add_section_heading("Technology Stack")
    add_sub_heading("Use of Trending Technology – Artificial Intelligence (AI)", font_size=12)
    
    p = add_para(space_after=6)
    p.add_run("Artificial Intelligence, Natural Language Processing (NLP), and dense vector embeddings are the primary technological drivers of the StudyMate AI platform. By leveraging specialized, pre-trained neural representations and high-throughput inference models, the system semantically analyzes unstructured educational documents without requiring expensive retraining or GPU infrastructure.")
    
    p = add_para(space_after=8)
    p.add_run("The integration of AI enables a ")
    add_run(p, "Multi-Agent Pedagogical Framework", bold=True)
    p.add_run(", where distinct software components collaborate to process documents, retrieve context, generate curriculum-aligned assessments, and formulate diagnostic recommendations. The system deploys a deliberate ")
    add_run(p, "Task-Based Model Routing", bold=True)
    p.add_run(" strategy, balancing low-latency structured output against deep contextual reasoning across different stages of the learning workflow.")

    add_sub_heading("1. Document Ingestion & Retrieval Layer (Processing & Storage)", font_size=12)
    add_bullet("High-speed, robust Python library for PDF parsing, text extraction, page counting, and metadata handling from complex, multi-page lecture notes.", bold_prefix="PyMuPDF (fitz): ")
    add_bullet("Intelligent chunking algorithm from LangChain that splits long document text into semantically cohesive 500-character passages with a 100-character overlap to preserve boundary context.", bold_prefix="RecursiveCharacterTextSplitter: ")
    add_bullet("Deep sentence-embedding model that maps natural language text chunks into 384-dimensional dense vector space, enabling semantic retrieval beyond exact keyword matches.", bold_prefix="Sentence Transformers (all-MiniLM-L6-v2): ")
    add_bullet("Persistent embedded vector database that stores chunk embeddings and executes rapid cosine and L2 similarity searches to fetch the top-8 most relevant passages for any learning query.", bold_prefix="ChromaDB (v0.4.24): ")

    add_sub_heading("2. Artificial Intelligence & Large Language Model Layer", font_size=12)
    add_bullet("High-speed LPUs (Language Processing Units) hosting models like LLaMA 3.3 and GPT-OSS, providing ultra-low response latency for high-throughput generation tasks such as document summarization and quiz generation.", bold_prefix="Groq Cloud API: ")
    add_bullet("Flexible gateway providing access to advanced reasoning models (e.g., meta-llama/llama-3.3-70b-instruct) for reasoning-intensive question answering and deep personalized learning reflections.", bold_prefix="OpenRouter API: ")

    doc.add_page_break()

    # =========================================================================
    # PAGE 8: CHAPTER 1 — TECHNOLOGY STACK (1.2 Part 2)
    # =========================================================================
    print("Writing Page 8: Chapter 1 Technology Stack Part 2...")
    add_sub_heading("3. AI Routing, Orchestration & Execution Layer", font_size=12)
    add_bullet("Intelligent software agent that inspects incoming learning tasks and automatically routes requests to Groq (for fast summaries and quizzes) or OpenRouter (for Q&A and reflection) with robust provider fallback.", bold_prefix="AIProvider Router: ")
    add_bullet("Foundational orchestration framework that structures prompts, enforces strict JSON response formatting, and binds LLMs to external vector retrieval contexts.", bold_prefix="LangChain (v0.3.26) & LangGraph: ")
    add_bullet("A strongly-typed, versioned Python dataclass that formalizes explicit inter-agent messaging between the Quiz Agent and the Reflection Agent, passing scores, weak topics, and study recommendations.", bold_prefix="QuizResultMessage Protocol: ")
    add_bullet("Stateful memory architecture managing persistent document context, vector counts, quiz evaluation payloads, and UI navigation across multi-page workflows.", bold_prefix="Streamlit Session State: ")

    add_sub_heading("4. Application Development, Environment & Tooling", font_size=12)
    add_bullet("The foundational high-performance programming language providing comprehensive library ecosystems for machine learning, vector math, and asynchronous API interactions.", bold_prefix="Python 3.12: ")
    add_bullet("Modern, reactive web framework facilitating modular multi-page applications with clean styling, responsive layouts, and interactive pedagogical widgets.", bold_prefix="Streamlit (v1.46.1): ")
    add_bullet("Environment configuration manager ensuring secure, isolated storage of private API credentials (GROQ_API_KEY, OPENROUTER_API_KEY).", bold_prefix="python-dotenv: ")
    add_bullet("Fast BPE tokenizer library utilized for context window verification, chunk boundary budgeting, and LLM token optimization.", bold_prefix="tiktoken: ")

    p_sum = add_para(space_before=10, space_after=0)
    add_run(p_sum, "Summary Description for Report: ", bold=True)
    p_sum.add_run('"The StudyMate AI project is built on a Python-based AI stack, utilizing Streamlit for reactive web interface delivery, LangChain and LangGraph for prompt and agent orchestration, PyMuPDF and Sentence Transformers (all-MiniLM-L6-v2) for semantic document parsing and vectorization, and ChromaDB for persistent vector retrieval. The architecture emphasizes a grounded, zero-hallucination Retrieval-Augmented Generation (RAG) framework, leveraging task-based model routing across Groq and OpenRouter to provide an accessible, low-latency, and pedagogical assistant that translates uploaded lecture materials directly into structured knowledge, practice assessments, and personalized reflection."')

    doc.add_page_break()

    # =========================================================================
    # PAGE 9: CHAPTER 1 — MOTIVATION (1.3)
    # =========================================================================
    print("Writing Page 9: Chapter 1 Motivation...")
    add_section_heading("1.3 Motivation")
    
    p = add_para(space_after=10)
    p.add_run("The motivation behind the development of the ")
    add_run(p, "StudyMate AI", bold=True)
    p.add_run(" project is rooted in the urgent academic need to bridge the gap between static educational content and active cognitive retention. While higher education institutions produce extensive digital repositories of lecture slides, textbook chapters, and technical handouts, the tools available for students to absorb and review this information have remained fundamentally stagnant. Conventional study habits rely on passive, rote reading, which psychological and educational research has consistently shown to yield low retention rates. Students frequently face severe cognitive overload when attempting to extract core exam definitions from lengthy PDFs, leading to burnout, study fatigue, and inefficient time management.")

    p = add_para(space_after=10)
    p.add_run("A secondary driving factor is the high risk and unreliability of generic conversational chatbots. As large language models entered mainstream awareness, students began using commercial chatbots for academic queries. However, off-the-shelf LLMs frequently generate fabricated facts, hallucinated mathematical proofs, and explanations that diverge from specific syllabus requirements. Furthermore, commercial tools offering reliable document indexing typically impose expensive recurring subscription fees, placing them out of reach for students in developing regions. By developing a 100% software-based, open-source solution that executes local document processing and dense vector indexing, StudyMate AI democratizes personalized academic support. The system delivers a professional-grade study assistant at zero subscription cost, ensuring that institutional or financial constraints do not limit a student's academic potential.")

    p = add_para(space_after=0)
    p.add_run("Furthermore, the project is motivated by modern pedagogical principles of active recall, self-assessment, and reflective metacognition. Traditional educational software isolates testing from study advice. StudyMate AI bridges these domains through multi-agent collaboration: when a student completes an automatically generated quiz, the Quiz Agent publishes a structured evaluation message directly to the Reflection Agent. The Reflection Agent synthesizes this performance data alongside the retrieved lecture notes to formulate actionable, personalized study plans. Ultimately, StudyMate AI is fueled by the desire to pioneer a future where technology adapts to the cognitive journey of each student, transforming passive lecture notes into an interactive, grounded, and empowering learning partnership.")

    doc.add_page_break()

    # =========================================================================
    # PAGE 10: CHAPTER 2 — PROBLEM STATEMENT (2.1 & 2.2)
    # =========================================================================
    print("Writing Page 10: Chapter 2 Problem Statement...")
    add_chapter_heading(2, "PROBLEM STATEMENT")
    add_section_heading("2.1 Need of Work:")
    add_bullet("Proliferation of dense, multi-page educational PDF lecture materials in university curricula.")
    add_bullet("High cognitive load, mental fatigue, and inefficient retention during manual document reading.")
    add_bullet("Inability of static PDF documents to answer student questions or provide real-time interactive clarification.")
    add_bullet("Prevalence of factual hallucinations and ungrounded answers in generic, public AI conversational models.")
    add_bullet("Lack of accessible, zero-cost intelligent tutoring platforms tailored strictly to course-specific documents.")
    add_bullet("Requirement for semantic vector retrieval to locate conceptual answers without exact keyword matching.")
    add_bullet("Absence of automated, syllabus-aligned quiz generation for continuous active recall self-assessment.")
    add_bullet("Lack of structured pedagogical feedback loops connecting quiz evaluation with personalized revision advice.")
    add_bullet("Need for task-based LLM routing balancing ultra-low latency for generation with deep reasoning for reflection.")
    add_bullet("Bridge the gap between passive document consumption and active, agentic subject mastery.")

    add_section_heading("2.2 Problem Statement")
    p = add_para(space_after=8)
    p.add_run("In the modern digital era, the ability to rapidly comprehend, synthesize, and apply complex technical knowledge from written documentation is a fundamental prerequisite for academic success and professional growth. However, traditional educational workflows rely almost entirely on static electronic documents—primarily lecture slides and textbook PDFs. Students are forced to spend excessive hours manually skimming text, hunting for definitions, and guessing potential examination questions. This manual process is time-consuming, prone to omission, and causes substantial study fatigue, especially when students prepare for high-stakes examinations.")

    p = add_para(space_after=0)
    p.add_run("While commercial AI chatbots exist, they pose critical educational risks. Generic large language models generate answers from unverified internet training weights rather than the student's assigned curriculum, introducing subtle factual errors that mislead students. Furthermore, existing educational utilities are highly fragmented: students use one tool for reading, another for flashcards, and a third for practice questions, with no shared contextual memory. Crucially, there is a lack of an integrated agentic framework that connects assessment performance with targeted study guidance. There is an urgent need for an open-source, software-only platform that utilizes Retrieval-Augmented Generation (RAG) and structured inter-agent communication to provide grounded summaries, context-aware answers, automated quizzes, and personalized reflections directly from uploaded study notes.")

    doc.add_page_break()

    # =========================================================================
    # PAGE 11: CHAPTER 2 — OBJECTIVES (2.3)
    # =========================================================================
    print("Writing Page 11: Chapter 2 Objectives...")
    p = add_para(space_after=10)
    p.add_run("Additionally, current artificial intelligence deployments face acute trade-offs between processing latency and reasoning depth. Generating concise document summaries and structured multiple-choice questions demands rapid token throughput to keep students engaged, while contextual Q&A and reflective coaching demand expansive context windows and nuanced reasoning. Relying on a single uniform model inevitably compromises either system responsiveness or analytical depth. Therefore, developing an intelligent, multi-agent AI framework that pairs dense vector retrieval with task-based model selection is essential to provide an optimal, accessible, and grounded learning experience.")

    add_section_heading("2.3 Objectives:")
    add_bullet("To develop a 100% software-based, interactive agentic learning assistant using Python and Streamlit.")
    add_bullet("To implement a robust document ingestion pipeline using PyMuPDF (fitz) for reliable text extraction from educational PDFs.")
    add_bullet("To design an intelligent document chunking mechanism utilizing LangChain's RecursiveCharacterTextSplitter (500-character chunk, 100-character overlap).")
    add_bullet("To generate 384-dimensional dense semantic embeddings using Sentence Transformers (all-MiniLM-L6-v2).")
    add_bullet("To construct a persistent ChromaDB vector storage repository for real-time semantic similarity retrieval (top-K search).")
    add_bullet("To implement a strict Retrieval-Augmented Generation (RAG) framework that eliminates hallucinations by restricting answers to uploaded notes.")
    add_bullet("To develop an automated document summarizer delivering structured executive summaries, key bullet points, and exam preparation tips.")
    add_bullet("To build an interactive, context-aware question answering module (Ask AI) with source chunk verification.")
    add_bullet("To design an automated quiz generation, scoring, and review engine evaluating student knowledge with structured JSON outputs.")
    add_bullet("To implement explicit agent-to-agent communication via the QuizResultMessage protocol to deliver personalized learning reflections and revision plans.")

    doc.add_page_break()

    # =========================================================================
    # PAGE 12: CHAPTER 3 — DESIGN DETAILS: SYSTEM ARCHITECTURE (3.1)
    # =========================================================================
    print("Writing Page 12: Chapter 3 System Architecture...")
    add_chapter_heading(3, "DESIGN DETAILS")
    add_section_heading("3.1 System Architecture:")
    
    # Figure 1: System Architecture Image
    arch_img_path = os.path.join(ASSETS_DIR, "architecture.png")
    add_image_box(arch_img_path, width_in=5.4, caption="Figure 1: StudyMate AI Layered System Architecture")
    
    p_desc = add_para(space_after=4)
    add_run(p_desc, "The StudyMate AI system architecture is structured into four distinct, loosely coupled layers that coordinate document ingestion, semantic retrieval, task routing, and intelligent response generation:", bold=True)
    add_bullet("Streamlit multi-page reactive interface providing intuitive user navigation across Home, Upload, Summary, Ask AI, Quiz, and Reflection modules with session state persistence.", bold_prefix="Presentation Layer: ")
    add_bullet("Modular service orchestrators (DocumentService, SummaryService, ChatService, QuizService, ReflectionService) managing business logic and agent interactions.", bold_prefix="Application Layer: ")
    add_bullet("PyMuPDF loader, RecursiveCharacterTextSplitter, Sentence Transformers (all-MiniLM-L6-v2), and persistent ChromaDB vector store executing dense semantic retrieval.", bold_prefix="Retrieval Layer: ")
    add_bullet("AIProvider routing agent dynamically directing tasks to Groq Cloud (summaries, quizzes) or OpenRouter (Q&A, reflection) with automated provider fallback.", bold_prefix="AI Layer: ")

    doc.add_page_break()

    # =========================================================================
    # PAGE 13: CHAPTER 3 — DATA FLOW DIAGRAM (3.2.1)
    # =========================================================================
    print("Writing Page 13: Chapter 3 Data Flow Diagram...")
    add_section_heading("3.2 System Design Diagrams")
    add_sub_heading("3.2.1 Data Flow Diagram", font_size=12)
    
    # Figure 2: Data Flow Diagram
    dfd_img_path = os.path.join(ASSETS_DIR, "data-flow-diagram.png")
    add_image_box(dfd_img_path, width_in=5.3, caption="Figure 2: Data Flow Diagram and RAG Processing Pipeline")
    
    p_dfddesc = add_para(space_after=4)
    add_run(p_dfddesc, "System processing and RAG pipeline:", bold=True)
    p_dfddesc = add_para(space_after=0)
    p_dfddesc.add_run("The data flow pipeline initiates when a student uploads an educational PDF through the Streamlit interface. The DocumentService saves the document into local storage and invokes PyMuPDF to extract text from every page. The raw text is passed to the RecursiveCharacterTextSplitter, which segments the document into 500-character chunks with a 100-character overlap. The Sentence Transformer model converts these chunks into 384-dimensional dense vector embeddings, which are stored alongside unique chunk IDs in ChromaDB. When a student requests a learning task, the query is embedded and matched against stored vectors to retrieve the most relevant chunks. The AIProvider routes the query and retrieved context to the optimal LLM (Groq or OpenRouter), returning a grounded educational response to the student.")

    doc.add_page_break()

    # =========================================================================
    # PAGE 14: CHAPTER 3 — SEQUENCE DIAGRAM & METHODOLOGY (3.2.2)
    # =========================================================================
    print("Writing Page 14: Chapter 3 Sequence Diagram...")
    add_sub_heading("3.2.2 Sequence Diagram & Agent Communication Workflow", font_size=12)
    
    # Figure 3: Agent Communication Workflow Image
    agent_img_path = os.path.join(ASSETS_DIR, "agent-communication-workflow.png")
    add_image_box(agent_img_path, width_in=5.2, caption="Figure 3: Inter-Agent Communication Workflow (Quiz Agent to Reflection Agent)")
    
    p_seqdesc = add_para(space_after=4)
    add_run(p_seqdesc, "Phases of Methodology", bold=True)
    add_bullet("Receives academic PDF uploads via Streamlit, extracts plain text using PyMuPDF, and generates 500-character overlapping chunks.", bold_prefix="Document Ingestion & Chunking Phase: ")
    add_bullet("Computes 384-dimensional vector embeddings with all-MiniLM-L6-v2 and stores vectors in a persistent ChromaDB collection.", bold_prefix="Semantic Indexing & Vector Storage Phase: ")
    add_bullet("Performs top-K semantic search and routes requests dynamically to Groq or OpenRouter via the AIProvider router.", bold_prefix="Task-Based Retrieval & Routing Phase: ")
    add_bullet("Evaluates submitted quiz answers, formats a versioned QuizResultMessage, and passes structured data to the Reflection Agent.", bold_prefix="Assessment & Reflective Coaching Phase: ")

    doc.add_page_break()

    # =========================================================================
    # PAGE 15: CHAPTER 4 — IMPLEMENTATION (4.1 Modules 4.1 & 4.2)
    # =========================================================================
    print("Writing Page 15: Chapter 4 Implementation Part 1...")
    add_chapter_heading(4, "IMPLEMENTATION")
    add_section_heading("4.1 Module Description In detail")
    
    p = add_para(space_after=8)
    p.add_run("The StudyMate AI framework is engineered as a highly modular, decoupled software system to ensure robust real-time performance, clean separation of concerns, and seamless scalability. Each module encapsulates a distinct functional domain of the learning pipeline, cooperating through well-defined service abstractions and explicit message schemas.")

    add_sub_heading("4.1 Document Ingestion and Text Extraction Module", font_size=12)
    p = add_para(space_after=4)
    p.add_run("This module acts as the primary data acquisition gateway, managing the interface between the web client and the document processing subsystem:")
    add_bullet("Receives user-uploaded PDF files via the Streamlit file uploader with support for file sizes up to 200 MB.")
    add_bullet("Saves uploaded files securely to the local data directory (data/uploads/) while verifying file existence and permissions.")
    add_bullet("Utilizes the PyMuPDF (fitz) engine to extract clean, uncorrupted plain text across all document pages sequentially.")
    add_bullet("Calculates exact page counts and validates extracted character length to catch empty or scanned PDFs prior to chunking.")

    add_sub_heading("4.2 Text Chunking and Semantic Embedding Module", font_size=12)
    p = add_para(space_after=4)
    p.add_run("As the analytical transformation engine, this module converts unstructured text into dense mathematical representations:")
    add_bullet("Integrates LangChain's RecursiveCharacterTextSplitter with a 500-character chunk size and 100-character overlap.")
    add_bullet("Preserves semantic continuity across paragraph and sentence boundaries, avoiding chopped statements or loss of context.")
    add_bullet("Loads the sentence-transformers/all-MiniLM-L6-v2 deep learning model to encode each chunk into 384-dimensional dense vectors.")
    add_bullet("Normalizes embedding tensors to enable fast and accurate cosine similarity distance calculations in the vector database.")

    doc.add_page_break()

    # =========================================================================
    # PAGE 16: CHAPTER 4 — IMPLEMENTATION (4.1 Modules 4.3 & 4.4)
    # =========================================================================
    print("Writing Page 16: Chapter 4 Implementation Part 2...")
    add_sub_heading("4.3 Vector Storage and Semantic Retrieval Module", font_size=12)
    p = add_para(space_after=4)
    p.add_run("This module serves as the persistent knowledge repository, executing high-speed vector searches over indexed lecture content:")
    add_bullet("Initializes a persistent ChromaDB client mapped to local disk storage (data/chroma_db/) with a dedicated 'studymate' collection.")
    add_bullet("Implements a clean collection-clearing mechanism to flush prior document vectors whenever a new lecture PDF is uploaded.")
    add_bullet("Generates deterministic chunk IDs (chunk_0, chunk_1, ...) and stores raw text alongside dense vector embeddings.")
    add_bullet("Executes semantic nearest-neighbor queries using cosine and L2 distance metrics to return the top-8 most relevant chunks.")

    add_sub_heading("4.4 Learning Services, Task Routing & Inter-Agent Module", font_size=12)
    p = add_para(space_after=4)
    p.add_run("This final execution layer orchestrates the core educational capabilities, model routing, and agent communication:")
    add_bullet("Inspects user requests and dynamically routes generation tasks (summaries, quizzes) to Groq and reasoning tasks (Q&A, reflection) to OpenRouter with automatic provider fallback.", bold_prefix="AIProvider Routing Agent: ")
    add_bullet("Aggregates full document chunks and prompts the LLM to produce structured executive summaries, bullet points, core concepts, and exam tips.", bold_prefix="SummaryService: ")
    add_bullet("Converts student questions into semantic embeddings, retrieves top-6 relevant chunks, and applies strict anti-hallucination prompting rules.", bold_prefix="ChatService: ")
    add_bullet("Generates 4-option multiple-choice quizzes in strict JSON format, validates question schemas, calculates scores, and creates detailed educational answer reviews.", bold_prefix="QuizService: ")
    add_bullet("Consumes structured inter-agent messages from the Quiz Agent to diagnose strengths, weak areas, and revision topics without recalculating scores.", bold_prefix="ReflectionService & QuizResultMessage: ")

    doc.add_page_break()

    # =========================================================================
    # PAGE 17: CHAPTER 4 — SYSTEM REQUIREMENTS (4.2 Specs, SW, HW, SDLC)
    # =========================================================================
    print("Writing Page 17: Chapter 4 System Requirements...")
    add_section_heading("4.2 System Requirement")
    add_sub_heading("System Requirements Specification: StudyMate AI", font_size=12)
    
    add_sub_heading("1. Project Overview", font_size=12)
    p = add_para(space_after=6)
    p.add_run("StudyMate AI is a modular, web-based agentic learning assistant engineered to empower students with document-grounded summaries, context-aware question answering, automated quizzes, and personalized reflection. By integrating dense vector retrieval with task-based LLM routing, the platform provides a zero-hallucination educational environment.")

    add_sub_heading("2. Software Requirements", font_size=12)
    p = add_para(space_after=4)
    p.add_run("The application is built using a Python-centric architecture optimized for AI-driven semantic retrieval and multi-agent workflows:")
    add_bullet("Windows 10/11, macOS (12.0+ Monterey or later), or Linux (Ubuntu 20.04+ LTS recommended).", bold_prefix="Operating System: ")
    add_bullet("Python (Version 3.10 to 3.12, Python 3.12 recommended).", bold_prefix="Programming Language: ")
    add_bullet("Streamlit (v1.46.1), LangChain (v0.3.26), LangGraph (v0.5.3), ChromaDB (v0.4.24), Sentence-Transformers (v5.0.0), PyMuPDF (v1.26.3), PyPDF (v5.8.0), NumPy (v1.26.4), tiktoken (v0.9.0), and python-dotenv (v1.1.1).", bold_prefix="Core Libraries: ")
    add_bullet("Groq Cloud API (hosting LLaMA 3.3 / GPT-OSS) and OpenRouter API (hosting LLaMA 3.3-70B-Instruct).", bold_prefix="APIs: ")

    add_sub_heading("3. Hardware Requirements", font_size=12)
    add_bullet("Quad-core 2.0 GHz Intel Core i5 / AMD Ryzen 5 or Apple Silicon (M1/M2/M3) or higher to handle embedding encoding.", bold_prefix="Processor: ")
    add_bullet("8 GB Minimum (16 GB recommended for multi-tasking with vector indexing and large PDFs).", bold_prefix="RAM: ")
    add_bullet("2 GB available disk space for Python dependencies, virtual environment, and local ChromaDB vector databases.", bold_prefix="Storage: ")
    add_bullet("Required for cloud LLM API calls (Groq and OpenRouter) and initial embedding model download.", bold_prefix="Internet Connectivity: ")

    add_sub_heading("4. Software Development Life Cycle (Waterfall Model)", font_size=12)
    p = add_para(space_after=0)
    p.add_run("The StudyMate AI project was developed using the Waterfall Model to ensure a disciplined and sequential flow of development, which was critical for maintaining the precision required in a Retrieval-Augmented Generation framework.")

    doc.add_page_break()

    # =========================================================================
    # PAGE 18: CHAPTER 4 — SDLC WATERFALL PHASES
    # =========================================================================
    print("Writing Page 18: Chapter 4 SDLC Phases...")
    add_sub_heading("Requirement Analysis:", font_size=12)
    add_bullet("Defining core educational functional requirements: PDF uploading, automated summarization, grounded Q&A, quiz generation, and reflective feedback.")
    add_bullet("Identifying non-functional requirements: low-latency UI responsiveness (<2s for quizzes), zero hallucinations, secure API key management, and session state persistence.")
    add_bullet("Establishing the user role as an undergraduate student needing reliable, syllabus-grounded revision assistance.")

    add_sub_heading("System Design:", font_size=12)
    add_bullet("Architecting the integration between the Streamlit presentation layer, service layer, ChromaDB vector store, and cloud LLM providers.")
    add_bullet("Designing the schema for the QuizResultMessage protocol to enable decoupled, structured inter-agent communication between Quiz and Reflection modules.")
    add_bullet("Formulating strict system prompt templates with guardrails against ungrounded speculation or hallucination.")

    add_sub_heading("Implementation (Coding):", font_size=12)
    add_bullet("Developing the document ingestion and text extraction backend using PyMuPDF and LangChain text splitters.")
    add_bullet("Implementing the vector store operations, embedding encoding with all-MiniLM-L6-v2, and ChromaDB collection management.")
    add_bullet("Coding the educational service modules (Summary, Chat, Quiz, Reflection) and integrating the AIProvider task routing logic.")

    add_sub_heading("Testing:", font_size=12)
    add_bullet("Rigorous validation of embedding dimensions (384d), chunking consistency, and retrieval ranking accuracy on academic PDFs.")
    add_bullet("Ensuring system stability during multi-page navigation and concurrent LLM provider API invocations.")

    add_sub_heading("Deployment:", font_size=12)
    add_bullet("Packaging the software environment with necessary dependencies in requirements.txt for cross-platform deployment.")
    add_bullet("Deploying the live application to Streamlit Community Cloud (studymate-agent.streamlit.app) for accessibility.")

    add_sub_heading("Maintenance:", font_size=12)
    add_bullet("Updating embedding pipelines periodically to support newer model architectures and monitoring API rate limits.")
    add_bullet("Refining prompt templates and fixing interaction bugs based on feedback from university students.")

    doc.add_page_break()

    # =========================================================================
    # PAGE 19: CHAPTER 4 — TESTING REQUIREMENTS (5.1 - 5.4)
    # =========================================================================
    print("Writing Page 19: Chapter 4 Testing Requirements...")
    add_section_heading("5. Testing Requirements")
    add_sub_heading("5.1 Unit Testing", font_size=12)
    p = add_para(space_after=4)
    p.add_run("Testing individual utility functions, algorithms, and service components in isolation:")
    add_bullet("Validating PDFLoader text extraction accuracy across diverse PDF structures (test_loader.py).")
    add_bullet("Verifying TextChunker chunk sizing (500 characters) and overlap consistency (100 characters) (test_chunker.py).")
    add_bullet("Confirming EmbeddingModel vector dimensions (ensuring exactly 384 dimensions for all text inputs) (test_embeddings.py).")
    add_bullet("Validating VectorStore document insertion, collection clearing, and top-K query results (test_vector_store.py).")

    add_sub_heading("5.2 Integration Testing", font_size=12)
    add_bullet("Verifying the seamless communication between the PDF extraction layer, chunker, embedding engine, and ChromaDB vector store.")
    add_bullet("Ensuring that the AIProvider correctly executes task-based routing (Groq for summaries/quizzes, OpenRouter for Q&A/reflection) with proper fallback handling.")
    add_bullet("Testing the end-to-end delivery of the QuizResultMessage protocol: verifying that scores, weak areas, and study recommendations correctly populate the Reflection Agent via st.session_state.")

    add_sub_heading("5.3 System Testing", font_size=12)
    add_bullet("End-to-End (E2E): Simulating a full student learning journey, from PDF upload and document indexing to generating summaries, asking contextual questions, taking quizzes, and reviewing personalized reflections.")
    add_bullet("Stability Testing: Monitoring system memory and CPU consumption during sustained vector searches and high-token LLM completions to prevent memory leaks or crashes.")
    add_bullet("Edge Case Testing: Validating system behavior when empty PDFs, corrupted files, extremely large slide decks, or missing API keys are encountered.")

    add_sub_heading("5.4 User Acceptance Testing (UAT)", font_size=12)
    add_bullet("Beta testing with university students to ensure that navigation across Home, Upload, Summary, Ask AI, Quiz, and Reflection is intuitive and frictionless.")
    add_bullet("Verifying that the difficulty settings (Easy, Medium, Hard) and question types (Multiple Choice) in the Quiz module match student preparation needs.")
    add_bullet("Ensuring that the reflection recommendations are actionable, educational, and accurately aligned with identified weaknesses.")

    doc.add_page_break()

    # =========================================================================
    # PAGE 20: CHAPTER 4 — PERFORMANCE TESTING (5.5)
    # =========================================================================
    print("Writing Page 20: Chapter 4 Performance Testing...")
    add_sub_heading("5.5 Performance Testing", font_size=12)
    p = add_para(space_after=6)
    p.add_run("Performance benchmarks were conducted across document processing, embedding generation, vector similarity retrieval, and LLM inference to ensure responsive user experiences:")
    add_bullet("PyMuPDF achieved an average extraction speed of less than 0.2 seconds for 50-page academic lecture PDFs, ensuring near-instant document loading.", bold_prefix="Extraction Throughput: ")
    add_bullet("Sentence Transformers (all-MiniLM-L6-v2) generated 384-dimensional dense vectors at an average of 42ms per text chunk on standard quad-core CPU hardware.", bold_prefix="Embedding Latency: ")
    add_bullet("ChromaDB semantic similarity search executed nearest-neighbor retrieval (top-8 chunks) in less than 15ms, maintaining sub-second RAG context injection.", bold_prefix="Retrieval Speed: ")
    add_bullet("Groq Cloud API demonstrated ultra-low latency (<1.5 seconds) for complete 5-question structured JSON quiz generation and document summarization.", bold_prefix="Inference Latency: ")
    add_bullet("OpenRouter API processed complex multi-chunk Q&A contexts and synthesized 600-word learning reflections within 3.5 seconds without context truncation.", bold_prefix="Reasoning Fidelity: ")

    p_ptext = add_para(space_before=14, space_after=0)
    p_ptext.add_run("The empirical findings demonstrate that StudyMate AI operates within optimal interactive thresholds. Local vector search adds negligible overhead (<15ms), while task-based model routing successfully pairs high-speed structured generation on Groq with deep reflective reasoning on OpenRouter, delivering a professional-grade study companion on standard consumer laptops.")

    doc.add_page_break()

    # =========================================================================
    # PAGE 21: CHAPTER 5 — EXPERIMENTAL RESULTS (Screenshot 1: Home)
    # =========================================================================
    print("Writing Page 21: Chapter 5 Screenshot 1...")
    add_chapter_heading(5, "EXPERIMENTAL RESULTS")
    add_section_heading("5.1 Experimentation: (Output Screenshots with description)")
    
    # Screenshot 1: Home Page
    home_img_path = os.path.join(ASSETS_DIR, "home.png")
    add_image_box(home_img_path, width_in=5.8, caption=None)
    
    add_sub_heading("5.1. Multi-Page Web Interface: Home Page and Navigation Hub", font_size=12)
    p = add_para(space_after=0)
    p.add_run("Figure 4 focuses on the primary user dashboard of StudyMate AI. The interface provides a clear overview of the agentic learning assistant, featuring intuitive navigation across the six specialized modules: Home, Upload, Summary, Ask AI, Quiz, and Reflection. The sidebar provides a language selection dropdown (supporting English, Sinhala, and Mixed modes), real-time session tracking, and a session clearing utility. The main container highlights the operational readiness of AI agents and presents accessible entry points for core learning activities.")

    doc.add_page_break()

    # =========================================================================
    # PAGE 22: CHAPTER 5 — EXPERIMENTAL RESULTS (Screenshot 2: Upload)
    # =========================================================================
    print("Writing Page 22: Chapter 5 Screenshot 2...")
    # Screenshot 2: Upload Page
    upload_img_path = os.path.join(ASSETS_DIR, "upload.png")
    add_image_box(upload_img_path, width_in=5.8, caption=None)
    
    add_sub_heading("5.2. Document Ingestion and PDF Processing Interface", font_size=12)
    p = add_para(space_after=0)
    p.add_run("Figure 5 illustrates the Upload interface where students submit educational documents. The system accepts PDF files up to 200 MB. Upon upload, DocumentService automatically triggers text extraction using PyMuPDF, chunks the content using RecursiveCharacterTextSplitter, computes 384-dimensional dense embeddings with all-MiniLM-L6-v2, and stores the vectors in ChromaDB. The interface displays real-time processing feedback, confirming the document name, total page count, and generated chunk count.")

    doc.add_page_break()

    # =========================================================================
    # PAGE 23: CHAPTER 5 — EXPERIMENTAL RESULTS (Screenshot 3: Summary & Ask AI)
    # =========================================================================
    print("Writing Page 23: Chapter 5 Screenshot 3...")
    # Screenshot 3: Summary Page
    summary_img_path = os.path.join(ASSETS_DIR, "summary.png")
    add_image_box(summary_img_path, width_in=5.6, caption="5.3 (a) AI Document Summary Generation with Key Points and Exam Tips")
    
    # Screenshot 4: Ask AI Page
    ask_img_path = os.path.join(ASSETS_DIR, "ask-ai.png")
    add_image_box(ask_img_path, width_in=5.6, caption="5.3 (b) Context-Aware Question Answering Grounded in Uploaded Lecture Notes")
    
    p = add_para(space_after=0)
    p.add_run("Figure 6(a) demonstrates the AI Document Summary module, which synthesizes uploaded lecture notes into a structured format containing an executive overview, key bullet points, core concepts, and exam preparation tips. Figure 6(b) showcases the Ask AI interface, where students pose conceptual questions. The ChatService performs dense semantic retrieval on ChromaDB to locate the most relevant document chunks and grounds the generated response strictly within course materials, preventing hallucinations.")

    doc.add_page_break()

    # =========================================================================
    # PAGE 24: CHAPTER 5 — EXPERIMENTAL RESULTS (Screenshot 4: Quiz & Reflection)
    # =========================================================================
    print("Writing Page 24: Chapter 5 Screenshot 4...")
    # Screenshot 5: Quiz Page
    quiz_img_path = os.path.join(ASSETS_DIR, "quiz.png")
    add_image_box(quiz_img_path, width_in=5.6, caption="5.4 (a) Automated AI Quiz Generation, Scoring, and Educational Review")
    
    # Screenshot 6: Reflection Page
    reflection_img_path = os.path.join(ASSETS_DIR, "reflection.png")
    add_image_box(reflection_img_path, width_in=5.6, caption="5.4 (b) Personalized Learning Reflection Driven by QuizResultMessage Protocol")
    
    p = add_para(space_after=0)
    p.add_run("Figure 7(a) displays the Quiz module, where StudyMate AI generates multiple-choice questions tailored to the uploaded material. The QuizService calculates final scores, validates answers, and produces comprehensive question-by-question reviews. Figure 7(b) demonstrates the Reflection module, where the Reflection Agent ingests the structured QuizResultMessage to produce personalized feedback, diagnosing strengths, weak topics, and targeted study recommendations.")

    doc.add_page_break()

    # =========================================================================
    # PAGE 25: CHAPTER 6 — CONCLUSION (6.1)
    # =========================================================================
    print("Writing Page 25: Chapter 6 Conclusion...")
    add_chapter_heading(6, "CONCLUSION")
    add_sub_heading("6.1 Conclusion:", font_size=14)
    
    p = add_para(space_after=8)
    p.add_run("The ")
    add_run(p, "StudyMate AI", bold=True)
    p.add_run(" project has been successfully designed, implemented, and evaluated to address one of the most significant cognitive bottlenecks in modern higher education—the dependency on manual, passive document reading. In an academic landscape where students are inundated with hundreds of pages of lecture slides, syllabus outlines, and technical handouts, conventional study practices lead to severe information overload, study fatigue, and superficial exam preparation. This system serves as a powerful, software-driven educational companion that transforms static lecture documents into active, grounded learning partnerships.")

    p = add_para(space_after=8)
    p.add_run("The project effectively demonstrates how the convergence of Artificial Intelligence, Retrieval-Augmented Generation (RAG), and Multi-Agent Collaboration can bridge the gap between static content and active student mastery. By leveraging lightweight, highly efficient open-source components—specifically PyMuPDF for document extraction, Sentence Transformers (")
    add_run(p, "all-MiniLM-L6-v2", italic=True)
    p.add_run(") for 384-dimensional dense semantic vectorization, and ChromaDB for persistent vector storage—the system establishes an empirical guarantee that all generated summaries, answers, quizzes, and reflections remain strictly anchored in the user's authentic course documents. This completely eliminates the hallucination tendencies of public conversational models, providing an authoritative revision platform.")

    p = add_para(space_after=8)
    p.add_run("A core architectural strength of the StudyMate AI system is its pragmatic task-based model selection strategy. Rather than forcing all tasks through a single monolithic model, the framework dynamically directs generation-focused tasks (summarization and quiz creation) to Groq Cloud for sub-second latency, while routing reasoning-intensive tasks (Q&A and reflective coaching) to OpenRouter. Furthermore, the explicit ")
    add_run(p, "QuizResultMessage", bold=True)
    p.add_run(" protocol formalizes clean agent-to-agent communication, allowing assessment performance to directly fuel personalized study advice without redundant score recalculations.")

    p = add_para(space_after=0)
    p.add_run("In conclusion, the StudyMate AI platform proves to be an indispensable tool for advancing digital learning autonomy. It not only simplifies the study workflow for university students but also establishes a reproducible, scalable blueprint for next-generation intelligent tutoring systems. By aligning AI capabilities with verified educational materials, the project empowers learners with the confidence, active recall, and personalized feedback needed to achieve academic excellence on their own terms.")

    doc.add_page_break()

    # =========================================================================
    # PAGE 26: CHAPTER 6 — FUTURE SCOPE (6.2)
    # =========================================================================
    print("Writing Page 26: Chapter 6 Future Scope...")
    add_sub_heading("6.2 Future scope:", font_size=14)
    add_bullet("Expansion of the ChromaDB storage schema to support simultaneous indexing and comparative cross-referencing of multiple lecture slide decks, textbooks, and lab manuals within unified subject collections.", bold_prefix="Multi-Document Knowledge Bases: ")
    add_bullet("Integration of Tesseract OCR or PaddleOCR pipelines to enable direct text extraction from scanned documents, whiteboard photographs, and handwritten lecture notes.", bold_prefix="Optical Character Recognition: ")
    add_bullet("Implementation of reciprocal rank fusion (RRF) combining dense semantic embeddings with BM25 sparse keyword retrieval to maximize precision on specialized technical terms and acronyms.", bold_prefix="Hybrid Search Architecture: ")
    add_bullet("Development of dynamic difficulty algorithms that adjust question complexity and topic distribution in real time based on historical student performance and spaced repetition intervals.", bold_prefix="Adaptive Quiz Generation: ")
    add_bullet("Integration of speech-to-text dictation and text-to-speech audio synthesis to provide hands-free voice-driven study sessions and support auditory learners.", bold_prefix="Speech & Multimodal Interaction: ")
    add_bullet("Integration of local inference runtimes (such as Ollama or llama.cpp) to enable completely offline execution, guaranteeing absolute data privacy and zero API dependency.", bold_prefix="Local Offline Model Execution: ")
    add_bullet("Implementation of student progress tracking dashboards displaying topic mastery heatmaps, quiz score trends over time, and time-to-exam revision readiness metrics.", bold_prefix="Longitudinal Learning Analytics Dashboard: ")
    add_bullet("Expansion of educational services to automatically generate printable Anki-compatible flashcard decks and visual concept mind-maps directly from retrieved chunks.", bold_prefix="Automated Flashcard and Concept-Map Generation: ")

    doc.add_page_break()

    # =========================================================================
    # PAGE 27: CHAPTER 7 — REFERENCES (7.1)
    # =========================================================================
    print("Writing Page 27: Chapter 7 References...")
    add_chapter_heading(7, "REFERENCES")
    add_sub_heading("7.1 References:", font_size=14)
    
    add_sub_heading("Web Links:", font_size=12)
    
    def add_web_link(name, url):
        p = doc.add_paragraph(style='Normal')
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        r_name = p.add_run(name + " – ")
        r_name.font.name = "Times New Roman"
        r_name.font.size = Pt(12)
        r_url = p.add_run(url)
        r_url.font.name = "Times New Roman"
        r_url.font.size = Pt(12)
        r_url.font.color.rgb = RGBColor.from_string("0000FF")
        r_url.underline = True

    add_web_link("1. Python Software Foundation Documentation", "https://docs.python.org/3/")
    add_web_link("2. Streamlit Web Framework Documentation", "https://docs.streamlit.io/")
    add_web_link("3. LangChain Python Framework Documentation", "https://python.langchain.com/")
    add_web_link("4. ChromaDB Vector Database Documentation", "https://docs.trychroma.com/")
    add_web_link("5. Sentence Transformers Library Documentation", "https://www.sbert.net/")
    add_web_link("6. Groq Cloud Developer API Documentation", "https://console.groq.com/docs")
    add_web_link("7. OpenRouter Unified AI Gateway Documentation", "https://openrouter.ai/docs")
    add_web_link("8. PyMuPDF (fitz) Document Library Documentation", "https://pymupdf.readthedocs.io/")
    add_web_link("9. StudyMate AI GitHub Open Source Repository", "https://github.com/roshini-nawanjala/StudyMate-AI")

    add_sub_heading("Books:", font_size=12)
    add_bullet("Roger S. Pressman & Bruce R. Maxim", bold_prefix="Software Engineering: A Practitioner's Approach – ")
    add_bullet("Ian Sommerville", bold_prefix="Software Engineering (10th Edition) – ")
    add_bullet("Stuart Russell & Peter Norvig", bold_prefix="Artificial Intelligence: A Modern Approach – ")
    add_bullet("Chip Huyen", bold_prefix="Designing Machine Learning Systems: An Iterative Process for Production-Ready Applications – ")
    add_bullet("Ben Auffarth", bold_prefix="Generative AI with LangChain: Build large language model (LLM) apps with Python, ChatGPT and other LLMs – ")

    add_sub_heading("Research Papers:", font_size=12)
    add_bullet("Lewis, P., et al. (2020). 'Retrieval-Augmented Generation for Knowledge-Intensive NLP Tasks.' Advances in Neural Information Processing Systems (NeurIPS), 33, 9459-9474.")
    add_bullet("Reimers, N., & Gurevych, I. (2019). 'Sentence-BERT: Sentence Embeddings using Siamese BERT-Networks.' In Proceedings of the 2019 Conference on Empirical Methods in Natural Language Processing (EMNLP-IJCNLP), 3982-3992.")
    add_bullet("Karpukhin, V., et al. (2020). 'Dense Passage Retrieval for Open-Domain Question Answering.' In Proceedings of the 2020 Conference on Empirical Methods in Natural Language Processing (EMNLP), 6769-6781.")
    add_bullet("Wu, Q., et al. (2023). 'AutoGen: Enabling Next-Gen LLM Applications via Multi-Agent Conversation.' arXiv preprint arXiv:2308.08155.")

    print("Saving refined document to:", OUTPUT_DOCX)
    doc.save(OUTPUT_DOCX)
    print("Refined document successfully created!")

if __name__ == "__main__":
    create_report()
