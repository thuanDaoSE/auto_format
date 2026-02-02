PROJECT: VIETDOC ASSISTANT - AI DEVELOPER GUIDE
TARGET AUDIENCE: AI Agents (DeepSeek, ChatGPT, Claude, Gemini). GOAL: Build a robust, hybrid document processing system using Java & Apache POI + LLM APIs. CURRENT STATUS: MVP / Refactoring Phase.

1. PROJECT PHILOSOPHY (CRITICAL)
This project follows a HYBRID ARCHITECTURE. You must strictly adhere to the separation of concerns:

LAYER 1: HARD LOGIC (The Muscle) - Apache POI

Responsibility: Precision formatting (Margins, Fonts, Line Spacing, Table Widths, Page Numbers).

Rule: NEVER use AI to guess margins or positions. Use hard-coded values based on Vietnam Academic Standards (e.g., Margins: 3.5cm Left, 2cm others).

Tech: Java poi-ooxml.

LAYER 2: SOFT LOGIC (The Brain) - AI API (DeepSeek/Gemma)

Responsibility: Content generation, context understanding.

Tasks:

Detect missing sections (e.g., "Lời cam đoan") and write them.

Generate captions for images/tables based on surrounding text context.

Paraphrase content (Future).

Rule: AI is expensive and slow. Only call AI when Logic Layer detects a need. Always implement Fallback Mechanisms (e.g., if API fails, insert placeholder text).

2. TECHNICAL STACK & CONFIGURATION
Language: Java (JDK 11+).

Core Library: Apache POI (poi-ooxml version 5.x).

JSON Handling: Native Java String parsing (to minimize dependencies) or Minimal JSON lib.

Network: java.net.http.HttpClient (Standard Java 11).

AI Providers:

DeepSeek V3: (OpenAI-compatible format).

Google Gemma 2/3 (Free Tier): (Google AI Studio format).

Requirement: Code must support Hot-Swapping between providers via a config flag.

3. CODING STANDARDS & REQUIREMENTS
A. Formatting Rules (Hard-Code)
Any code written must apply these standards by default:

Font: Times New Roman, Size 13 (Body), Size 14 + Bold (Heading 1).

Paragraph:

Alignment: Justified (Both).

Spacing: Before = 0pt, After = 6pt (120 twips).

Line Spacing: 1.5 lines.

Margins: Top/Bottom/Right = 2.0cm, Left = 3.0cm or 3.5cm.

Tables: Auto-fit to window (Width ~16cm).

B. AI Integration Rules
Context Window: When sending text to AI for captioning, only send 150 chars above and 100 chars below the image. Do not send the whole document.

System Prompt: Always enforce the persona: "You are a strict academic assistant. Output JSON only."

Error Handling: Wrap all API calls in try-catch. If AI fails, log error and continue formatting the rest of the document. Do not crash the app.

C. Debugging & Testing Mode
The code must have a DEBUG_MODE flag:

true: Do not call real AI API. Insert hard-coded strings (e.g., "Hình [TEST]: ...") to verify positioning.

false: Call real API.

4. IMPLEMENTATION CHECKLIST (DEFINITION OF DONE)
An AI agent working on this project must ensure the following features are working:

[ ] Core Format: Document opens in Word with correct margins (check Ruler) and Font (Times New Roman).

[ ] Table Fix: Tables do not overflow the page width.

[ ] Structure Check: Code detects if "Lời cam đoan" is missing.

[ ] API Switch: Changing CURRENT_PROVIDER variable switches between DeepSeek and Google URLs correctly.

[ ] JSON Parsing: Code correctly extracts content from both Google's candidates.content.parts.text and DeepSeek's choices.message.content.

[ ] Safety: The output file is saved successfully even if the input file has weird formatting (MathType, Charts).

5. REQUIRED PROJECT STRUCTURE
src/
└── vn/
    └── vietdoc/
        └── vietdoc_assistant/
            ├── VietDocCore.java       # Main Entry & Logic Orchestrator
            ├── utils/
            │   ├── POIHelper.java     # Helper for Margins, Fonts, Tables
            │   ├── AIClient.java      # HTTP Client for DeepSeek/Google
            │   └── TextUtils.java     # Context extraction, String parsing
            └── config/
                └── AppConfig.java     # API Keys, Constants (Margins, Fonts)
6. INSTRUCTION FOR AI AGENTS (HOW TO CONTRIBUTE)
Read: Analyze VietDocCore.java to understand the current state.

Refactor: If the code is monolithic (all in one file), suggest splitting it into utils as per structure above.

Test: Write a main method that processes a sample input.docx.

Verify:

Does the code compile without external dependencies (except POI)?

Does it handle Rate Limits (for Google Free Tier)?

Does it skip MathML/Charts to avoid corruption?

COMMAND TO AI: "Start by analyzing the provided Java file. Refactor it to meet the 'Hybrid Architecture' defined above. Implement the 'Debug Mode' first."