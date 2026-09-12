\# Hackathon Needle Agent



A small on-device AI model (Needle 2, 45M params, 13.7MB) fine-tuned to turn natural language instructions into tool calls, with real execution for common tasks (Word, YouTube, browser, email drafts, notes).



\## Setup



1\. Clone this repo

2\. Install Python 3.10+ if not already installed

3\. Install dependencies:
4. Run the agent:

5\. Type an instruction when prompted, e.g.:
	write hi my name is archishman into notes.docx

&#x09;open youtube and play fifa theme song

&#x09;mail xyz@gmail.com saying hello




\## What it can do



The model understands 26 tools covering: opening apps, browsing, YouTube search/play, Excel actions (write cell, apply formula, etc.), Word document writing, email drafting, file operations (open/find/delete/rename), notes, reminders, and generic UI actions (type/click).



Some tools are fully wired to real execution (Word, YouTube, browser, Gmail draft, notes, open app/file/folder). Others return a structured JSON action for the calling code to execute (useful for integrating into a larger orchestrator).



\## Model files



\- `my\_agent.cact` — the fine-tuned Needle 2 model (13.7MB)

\- `agent.py` — tool definitions + execution logic

\- `requirements.txt` — Python dependencies



\## Notes



\- Built with \[Cactus Needle 2](https://cactuscompute.com/) — a 45M-param on-device tool-calling model

\- Fine-tuned via LoRA on \~440 hand-verified + LLM-generated examples across 26 real-world tasks

