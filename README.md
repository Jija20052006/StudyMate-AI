# 📚 StudyMate AI — Your Intelligent Learning Assistant

**Learn smarter. Revise faster. Grow with AI.**

StudyMate AI is an AI-powered learning assistant designed to make studying more interactive, personalized, and efficient. It helps students transform their study materials into structured summaries, interactive question-answering experiences, quizzes, and reflection activities.

## ✨ Key Features

### 📄 1. Upload Study Materials
- Upload PDF study documents.
- Extract and process document content.
- Divide content into smaller chunks for efficient retrieval.

### 📝 2. AI-Powered Summarization
- Generate concise summaries from uploaded study materials.
- Understand important concepts more efficiently.
- Reduce the time spent revising lengthy documents.

### 💬 3. Ask AI — Document-Based Q&A
- Ask questions about uploaded study materials.
- Retrieve relevant information using semantic search.
- Generate contextual answers using Large Language Models (LLMs).

### 🧠 4. Interactive Quiz Generation
- Generate quizzes based on study materials.
- Test your understanding of important concepts.
- Practice and strengthen your knowledge.

### 💡 5. Reflection and Learning Improvement
- Reflect on your learning progress.
- Identify areas that need improvement.
- Encourage continuous learning and self-assessment.

## 🛠️ Technology Stack

| Component | Technology |
|---|---|
| Frontend | Streamlit |
| Programming Language | Python |
| Language Models | Groq / OpenRouter |
| Embeddings | Sentence Transformers |
| Embedding Model | all-MiniLM-L6-v2 |
| Vector Database | ChromaDB |
| Document Processing | PDF extraction and text chunking |
| AI Architecture | Retrieval-Augmented Generation (RAG) |

## 🏗️ How It Works

StudyMate AI uses a Retrieval-Augmented Generation (RAG) workflow to provide answers grounded in uploaded learning materials.

1. **Document Upload:** The student uploads a PDF.
2. **Text Extraction:** Text is extracted from the document.
3. **Text Chunking:** Extracted content is divided into manageable chunks.
4. **Embedding Generation:** Sentence Transformers converts text chunks into vector embeddings.
5. **Vector Storage:** Embeddings are stored in ChromaDB.
6. **Semantic Retrieval:** Relevant chunks are retrieved based on the student's question.
7. **AI Response Generation:** The retrieved context is provided to an LLM to generate a relevant response.

## 📂 Project Modules

- **Home:** Main dashboard and application navigation.
- **Upload:** Upload and process study materials.
- **Summary:** Generate AI-powered summaries.
- **Ask AI:** Ask questions about uploaded documents.
- **Quiz:** Practice through AI-generated quizzes.
- **Reflection:** Review learning and identify improvement areas.

## ⚙️ Installation and Setup

### Prerequisites

- Python 3.10 or compatible version
- pip
- Git
- API key for the configured LLM provider

### 1. Clone the Repository

```bash
git clone https://github.com/Jija20052006/StudyMate-AI.git
cd StudyMate-AI
```

### 2. Create a Virtual Environment

```bash
python -m venv .venv
```

Activate it on Windows:

```powershell
.venv\Scripts\activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Configure Environment Variables

Create a `.env` file in the project root and add the API keys required by your configured providers.

Example:

```env
GROQ_API_KEY=your_groq_api_key
OPENROUTER_API_KEY=your_openrouter_api_key
```

Use only the keys required by your application. Never upload your `.env` file to GitHub.

### 5. Run the Application

```bash
streamlit run app.py
```

Replace `app.py` with your actual Streamlit entry-point filename if it is different.

## 🔐 Security

- API keys should be stored in environment variables.
- The `.env` file should remain excluded through `.gitignore`.
- Never commit private credentials or sensitive configuration files.

## 🚀 Future Enhancements

- Personalized learning recommendations.
- Learning progress analytics.
- More advanced quiz customization.
- Support for additional document formats.
- Improved adaptive learning experiences.

## 👩‍💻 Project Information

**Project:** StudyMate AI  
**Category:** Artificial Intelligence / EdTech  
**Architecture:** Retrieval-Augmented Generation (RAG)  
**Interface:** Streamlit

---

### StudyMate AI

**Learn Smarter. Understand Better. Grow Faster.**