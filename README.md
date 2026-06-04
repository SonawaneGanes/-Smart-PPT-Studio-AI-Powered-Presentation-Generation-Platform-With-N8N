# 🚀 AI-Powered PowerPoint Presentation Generator with n8n Workflow Automation

Transform natural language prompts into professional PowerPoint presentations using AI, Python, Streamlit, and n8n automation.

---

## 📌 Project Overview

The AI-Powered PowerPoint Presentation Generator is an intelligent automation platform that creates professional PowerPoint presentations from user prompts. The application combines AI-generated content, workflow automation, and PowerPoint generation to significantly reduce the time required to create presentation-ready slides.

Users simply enter a topic or detailed prompt, and the system automatically generates a visually appealing PowerPoint presentation with structured content, modern layouts, and professional styling.

---

## 🎯 Features

### Core Features

* AI-powered presentation generation
* Dynamic topic-based content creation
* Automatic PowerPoint (.pptx) generation
* Modern and professional slide layouts
* Interactive Streamlit web interface
* One-click PPT download

### Advanced Features

* n8n workflow automation integration
* Dynamic theme selection based on topic
* Automated content structuring
* Error handling and validation
* Modular architecture
* Scalable design for future AI integrations

---

## 🏗️ System Architecture

```text
User Prompt
      │
      ▼
 Streamlit Frontend
      │
      ▼
 n8n Webhook Workflow
      │
      ▼
 AI Content Generation
      │
      ▼
 Python PPT Generator
 (python-pptx)
      │
      ▼
 PowerPoint File (.pptx)
      │
      ▼
 Download to User
```

---

## 🛠️ Tech Stack

### Frontend

* Streamlit
* HTML
* CSS

### Backend

* Python

### Presentation Generation

* python-pptx

### Automation

* n8n

### Communication

* REST APIs
* Requests

### Utilities

* Subprocess
* OS
* Sys
* Traceback

---

## 📂 Project Structure

```text
AI-Powered-PPT-Generator/
│
├── app.py
├── app1.py
├── generated_presentation.py
├── requirements.txt
├── assets/
├── output/
├── screenshots/
├── README.md
│
└── generated_presentation.pptx
```

---

## ⚙️ Installation

### Clone Repository

```bash
git clone https://github.com/yourusername/AI-Powered-PPT-Generator.git

cd AI-Powered-PPT-Generator
```

### Create Virtual Environment

```bash
python -m venv venv
```

### Activate Environment

Windows:

```bash
venv\Scripts\activate
```

Linux/Mac:

```bash
source venv/bin/activate
```

### Install Dependencies

```bash
pip install -r requirements.txt
```

---

## ▶️ Running the Application

### Start n8n Workflow

Ensure your n8n workflow is running and webhook endpoint is active.

Example:

```text
http://localhost:5678/webhook-test/Agentic-Ai-Power-Point-Generator.-main
```

### Run Streamlit App

```bash
streamlit run app.py
```

---

## 📝 Usage

### Step 1

Open the application in your browser.

### Step 2

Enter a prompt.

Example:

```text
Create a professional presentation on Artificial Intelligence in Healthcare
```

### Step 3

Click:

```text
Generate Presentation
```

### Step 4

The system will:

* Send prompt to n8n
* Generate content
* Create PowerPoint slides
* Export PPTX file

### Step 5

Download the generated presentation.

---

## 📊 Sample Topics

* Artificial Intelligence
* Machine Learning
* Cricket Analytics
* Data Science
* Blockchain Technology
* Healthcare Innovation
* Renewable Energy
* Cyber Security
* Cloud Computing
* Business Intelligence

---

## 📸 Screenshots

### Application Dashboard

```text
screenshots/dashboard.png
```

### Generated Presentation

```text
screenshots/presentation.png
```

### Download Page

```text
screenshots/download.png
```

---

## 🚧 Challenges Faced

* Dynamic content generation for different domains
* PowerPoint formatting automation
* Handling AI-generated code safely
* Managing topic-specific slide creation
* Error handling during PPT generation
* Workflow integration with n8n

---

## 🔮 Future Improvements

* OpenAI Integration
* Gemini Integration
* Multi-language presentations
* AI-generated images
* PPT themes marketplace
* PDF export support
* Voice-to-Presentation generation
* Cloud deployment
* User authentication
* Presentation templates library

---

## 📈 Project Impact

This project demonstrates:

* Python Development
* API Integration
* Workflow Automation
* AI Application Development
* Streamlit Development
* PowerPoint Automation
* Software Engineering Best Practices

The solution reduces manual presentation creation effort by automating content generation, formatting, and slide design.

---

## 📦 Requirements

```text
streamlit
python-pptx
requests
```

Install using:

```bash
pip install -r requirements.txt
```

---

## 👨‍💻 Author

Ganesh Sonawane

Software Developer | AI Enthusiast | Data Science & Automation

---

## ⭐ Acknowledgements

* Python
* Streamlit
* python-pptx
* n8n
* Open Source Community

---

### If you like this project, consider giving it a ⭐ on GitHub.
