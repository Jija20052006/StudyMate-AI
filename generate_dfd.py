import matplotlib.pyplot as plt
import matplotlib.patches as patches

fig, ax = plt.subplots(figsize=(10.5, 7.5), dpi=300)
ax.set_xlim(0, 100)
ax.set_ylim(0, 100)
ax.axis('off')

fig.patch.set_facecolor('#FFFFFF')
ax.set_facecolor('#FFFFFF')

# Color palette
c_blue = '#1E3A8A'
c_light_blue = '#EFF6FF'
c_border_blue = '#3B82F6'
c_green = '#065F46'
c_light_green = '#ECFDF5'
c_border_green = '#10B981'
c_purple = '#581C87'
c_light_purple = '#F5F3FF'
c_border_purple = '#8B5CF6'
c_amber = '#78350F'
c_light_amber = '#FFFBEB'
c_border_amber = '#F59E0B'
c_text_dark = '#111827'
c_text_muted = '#4B5563'

def draw_box(x, y, w, h, title, subtitle, bg_color, border_color, title_color=c_text_dark):
    rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.8,rounding_size=1.5",
                                  facecolor=bg_color, edgecolor=border_color, linewidth=2, zorder=2)
    ax.add_patch(rect)
    ax.text(x + w/2, y + h*0.65, title, ha='center', va='center', fontsize=10, fontweight='bold', color=title_color, zorder=3)
    if subtitle:
        ax.text(x + w/2, y + h*0.28, subtitle, ha='center', va='center', fontsize=8, color=c_text_muted, zorder=3)

def draw_db(x, y, w, h, title, subtitle, bg_color, border_color):
    rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.8,rounding_size=2.0",
                                  facecolor=bg_color, edgecolor=border_color, linewidth=2.2, zorder=2)
    ax.add_patch(rect)
    ax.text(x + w/2, y + h*0.65, title, ha='center', va='center', fontsize=10, fontweight='bold', color='#4C1D95', zorder=3)
    ax.text(x + w/2, y + h*0.28, subtitle, ha='center', va='center', fontsize=8, color=c_text_muted, zorder=3)

def draw_arrow(x1, y1, x2, y2, label="", label_pos=(0, 0), rad=0.0):
    connectionstyle = f"arc3,rad={rad}" if rad != 0.0 else "arc3,rad=0"
    ax.annotate("", xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle="-|>", color="#2563EB", lw=1.8, mutation_scale=13, connectionstyle=connectionstyle),
                zorder=4)
    if label:
        lx = (x1 + x2)/2 + label_pos[0]
        ly = (y1 + y2)/2 + label_pos[1]
        ax.text(lx, ly, label, ha='center', va='center', fontsize=7.8, fontweight='bold', color='#1E40AF',
                bbox=dict(boxstyle='round,pad=0.2', facecolor='#FFFFFF', edgecolor='#93C5FD', lw=0.8, alpha=0.98), zorder=5)

# Title Header
ax.text(50, 96.5, "STUDYMATE AI — DATA FLOW & RAG PIPELINE DIAGRAM", ha='center', va='center',
        fontsize=12, fontweight='bold', color='#1E3A8A')
ax.plot([8, 92], [93.5, 93.5], color='#CBD5E1', lw=1.2)

# Row 1: Ingestion
draw_box(4, 76, 23, 12, "Student / User", "Web Browser Client", c_light_blue, c_border_blue, c_blue)
draw_box(38, 76, 24, 12, "PDF Ingestion", "PyMuPDF (fitz) Loader", c_light_blue, c_border_blue, c_blue)
draw_box(73, 76, 23, 12, "Text Chunking", "Recursive Splitter (500/100)", c_light_blue, c_border_blue, c_blue)

draw_arrow(27, 82, 38, 82, "Uploads PDF", (0, 3.2))
draw_arrow(62, 82, 73, 82, "Raw Text", (0, 3.2))

# Downward to embedding
draw_arrow(84.5, 76, 84.5, 62, "Text Chunks", (7.5, 0))

# Row 2: Embedding & Storage
draw_box(73, 50, 23, 12, "Embedding Model", "all-MiniLM-L6-v2 (384d)", c_light_green, c_border_green, c_green)
draw_db(38, 50, 24, 12, "ChromaDB Storage", "Vector DB ('studymate')", c_light_purple, c_border_purple)
draw_box(4, 50, 23, 12, "Retrieval Engine", "Semantic Top-K Search (k=8)", c_light_purple, c_border_purple, c_purple)

draw_arrow(73, 56, 62, 56, "384d Vectors", (0, 3.2))
draw_arrow(38, 56, 27, 56, "Ranked Chunks", (0, 3.2))

# Downward context
draw_arrow(11, 50, 11, 36, "Grounded Context", (-8.5, 0))

# User initiates request
draw_arrow(20, 76, 20, 50, "Queries & Tasks", (9, 0))

# Row 3: Services & LLM Layer
draw_box(4, 22, 23, 14, "Learning Services", "Summary | Ask AI | Quiz\nReview | Reflection", c_light_amber, c_border_amber, c_amber)
draw_box(38, 22, 24, 14, "AIProvider Router", "Task-Based Model Routing\n(preferred vs. fallback)", c_light_blue, c_border_blue, c_blue)
draw_box(73, 22, 23, 14, "Dual LLM Inference", "Groq: LLaMA 3.3 / GPT-OSS\nOpenRouter: LLaMA 3.3-70B", c_light_green, c_border_green, c_green)

draw_arrow(27, 30, 38, 30, "Prompt + Task", (0, 3.2))
draw_arrow(62, 30, 73, 30, "API Call", (0, 3.2))
draw_arrow(73, 25, 27, 25, "Grounded AI Response", (0, -3.2), rad=0.1)

# Return learning results to student
draw_arrow(6, 36, 6, 76, "Learning Results", (-7, 0))

# Row 4: Inter-Agent Protocol
draw_box(28, 3, 44, 11, "QuizResultMessage Protocol", "Structured Inter-Agent Contract (Quiz Agent → Reflection Agent)", '#FEF3C7', '#F59E0B', '#92400E')
draw_arrow(16, 22, 28, 8.5, "Quiz Evaluation Data", (-7, -3.5), rad=-0.12)
draw_arrow(72, 8.5, 25, 22, "Personalized Feedback Loop", (16, -3.5), rad=-0.12)

plt.tight_layout()
output_path = r"c:\Users\DELL\Desktop\StudyMate-AI-main\assets\screenshots\data-flow-diagram.png"
plt.savefig(output_path, dpi=300, bbox_inches='tight', pad_inches=0.15)
plt.close()
print("Saved clean, un-overlapping DFD to:", output_path)
