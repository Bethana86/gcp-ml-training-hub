# GCP Professional ML Engineer Certification Training Hub

Welcome to the training repository for the **Google Cloud Professional Machine Learning Engineer Certification**. This hub contains an interactive training application, study guides, and a high-fidelity presentation detailing the **Vertex AI Feature Store**.

## 🚀 Key Features

* **Interactive Web Presentation Deck**: A premium 12-slide web application built with clean HTML/CSS/JS, featuring:
  * Glassmorphic card styling, responsive 16:9 widescreen layout, and smooth animations.
  * Interactive SVG diagrams for **Active Data Ingestion**, **Online Serving Loops**, and **Point-in-Time Joins** (showing future data leakage protection).
  * Professional syntax-highlighted IDE panels for Python SDK queries and `gcloud` CLI commands.
  * Structured real-world MLOps examples on every slide (e.g. credit card fraud prediction skew, ride-sharing matches, and model monitoring drift loops).
* **PowerPoint Slide Deck (`.pptx`)**: A standalone, natively styled presentation built using a Python-based generator to match the web presentation design system.
* **Curriculum App**: An MLOps certification study portal mapping out 6 main training courses, 4 practice simulators, and 25 practice questions.

---

## 🎨 Theme Design System

The presentation is fully compiled using a premium **Light-Grey (Zinc/Platinum)** theme:
* **Background**: Soft clean zinc-grey (`#e4e4e7` / `RGBColor(228, 228, 230)`)
* **Slide Cards**: Pure white containers with subtle shadows and borders
* **Typography**: Deep charcoal (`#18181b`) Segoe UI and Inter fonts
* **Code highlighting**: Custom GitHub-Light color token system
* **Visual contrast**: High-visibility blue, green, and pink highlights for diagrams and alerts

---

## 📁 Repository Structure

```directory
├── index.html                           # Presentation Main Portal
├── style.css                            # CSS Design Tokens & Animation Keyframes
├── app.js                               # Navigation, Transitions & Page Counter Logic
├── featurestore-presentation/
│   ├── index.html                       # Standalone Feature Store Web slides
│   ├── style.css                        # Slide specific light-grey styles
│   ├── app.js                           # Slide navigation controller
│   ├── generate_pptx.py                 # python-pptx Slide Compiler Script
│   └── Vertex_AI_Feature_Store.pptx     # Compiled Widescreen PowerPoint Deck
├── generative-ai-presentation/
│   ├── index.html                       # Standalone Generative AI Web slides
│   ├── style.css                        # Slide specific light-grey styles
│   ├── app.js                           # Slide navigation controller
│   ├── generate_pptx.py                 # python-pptx Slide Compiler Script
│   └── Introduction_to_Generative_AI.pptx # Compiled Widescreen PowerPoint Deck
├── llm-presentation/
│   ├── index.html                       # Standalone LLM Web slides
│   ├── style.css                        # Slide specific light-grey styles
│   ├── app.js                           # Slide navigation controller
│   ├── generate_pptx.py                 # python-pptx Slide Compiler Script
│   └── Introduction_to_LLM.pptx         # Compiled Widescreen PowerPoint Deck
└── .gitignore                           # Excludes logs and python cache files
```

---

## 🛠️ How to Run Locally

### 1. View the Interactive Web Slides
Start a local web server in the root of this project and navigate to either slide deck:
```bash
# Using Python
python -m http.server 8089
```
Then open your browser and navigate to:
* Vertex AI Feature Store Presentation: 👉 [http://localhost:8089/featurestore-presentation/](http://localhost:8089/featurestore-presentation/)
* Introduction to Generative AI Presentation: 👉 [http://localhost:8089/generative-ai-presentation/](http://localhost:8089/generative-ai-presentation/)
* Introduction to LLM Presentation: 👉 [http://localhost:8089/llm-presentation/](http://localhost:8089/llm-presentation/)

### 2. Generate/Recompile the PowerPoint Slide Decks
To modify slide contents or colors and re-compile either PowerPoint file, ensure you have Python and `python-pptx` installed:
```bash
# Install requirements
pip install python-pptx

# Run compiler for Feature Store
python featurestore-presentation/generate_pptx.py

# Run compiler for Generative AI
python generative-ai-presentation/generate_pptx.py

# Run compiler for Introduction to LLM
python llm-presentation/generate_pptx.py
```
This will compile and write/overwrite the `.pptx` file in their respective directories.

---

## 🎓 Presenter Details

* **Presenter**: Bethanasamy Rajamani
* **Audience**: Professional ML Engineer Candidates
