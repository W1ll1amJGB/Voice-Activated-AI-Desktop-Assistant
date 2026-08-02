Key Features:

🗣️ Voice Interaction: Speak to the agent (microphone) and hear responses (text-to-speech) using speech_recognition and pyttsx3.

🧠 Local LLM: Powered by Ollama (Llama 3.2) – no API costs, full privacy, works offline for general queries.

🖥️ System Tools: Create folders, delete/move files, open applications (WhatsApp, VS Code, Chrome, etc.), and empty the recycle bin.

🌐 Web Search: Integrated with DuckDuckGo for real-time information when internet is available.

📄 Document Creation: Generate .txt, .docx, .xlsx, and .pdf files on the fly.

🛡️ Safety First: Dangerous actions (like deleting files) require user confirmation via terminal.


### Requeriments

langchain>=0.3.0
langchain-ollama>=0.2.0
langchain-community>=0.3.0
langgraph>=0.3.0
python-dotenv>=1.0.0
duckduckgo-search>=6.0.0
speechrecognition>=3.10.0
pyttsx3>=2.90
pyaudio>=0.2.11
python-docx>=1.1.0
openpyxl>=3.1.0
fpdf>=1.7.2


# 🗣️ Voice-Activated AI Desktop Assistant

[![Python](https://img.shields.io/badge/Python-3.11+-blue.svg)](https://www.python.org/)
[![Ollama](https://img.shields.io/badge/Ollama-Llama3.2-green.svg)](https://ollama.com/)
[![LangChain](https://img.shields.io/badge/LangChain-0.3.x-orange.svg)](https://www.langchain.com/)

This is a fully local, voice-controlled AI assistant capable of managing your files, opening applications, searching the web, and answering questions using **Llama 3.2** via **Ollama**, orchestrated by **LangChain** and **LangGraph**.

## ✨ Features

- **Voice Interface**: Speak your commands and listen to responses (hands-free).
- **Local Privacy**: Runs entirely offline (except for online search and Google Speech API, which are optional).
- **System Control**: Create/delete folders, open installed applications (WhatsApp, Chrome, VS Code), and empty the recycle bin.
- **Document Generation**: Create Word (DOCX), Excel (XLSX), PDF, and TXT files instantly.
- **Web Search**: Fetches real-time information using DuckDuckGo when an internet connection is available.
- **Interactive Mode**: Choose between **Voice** or **Keyboard** input at startup.

## 🚀 Getting Started

### Prerequisites
1. **Python 3.11+** installed.
2. **Ollama** installed and running on your system (download from [ollama.com](https://ollama.com)).
3. A working **microphone** and **speakers** for voice mode.

