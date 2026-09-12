import needle
import win32com.client
import webbrowser
import urllib.parse
import os

DOWNLOADS_PATH = os.path.expanduser("~\\Downloads")

# ---------- Real, working tools ----------

@needle.tool
def write_text(filename: str, content: str):
    "Write text content into a Word document."
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    doc = word.Documents.Add()
    doc.Content.Text = content
    doc.SaveAs(os.path.join(DOWNLOADS_PATH, filename))
    return {"action": "write_text", "filename": filename, "status": "done"}

@needle.tool
def open_browser():
    "Open a web browser."
    webbrowser.open("https://www.google.com")
    return {"action": "open_browser", "status": "done"}

@needle.tool
def search_youtube(query: str):
    "Search YouTube for a query."
    url = "https://www.youtube.com/results?search_query=" + urllib.parse.quote(query)
    webbrowser.open(url)
    return {"action": "search_youtube", "query": query, "status": "done"}

@needle.tool
def play_video(title: str):
    "Play a video by title on YouTube."
    url = "https://www.youtube.com/results?search_query=" + urllib.parse.quote(title)
    webbrowser.open(url)
    return {"action": "play_video", "title": title, "status": "done"}

@needle.tool
def compose_email(to: str, subject: str, body: str):
    "Compose a draft email without sending it."
    params = {"view": "cm", "fs": "1", "to": to, "su": subject, "body": body}
    url = "https://mail.google.com/mail/?" + urllib.parse.urlencode(params)
    webbrowser.open(url)
    return {"action": "compose_email", "to": to, "status": "draft opened"}

@needle.tool
def open_website(url: str):
    "Open a specific website by URL."
    if not url.startswith("http"):
        url = "https://" + url
    webbrowser.open(url)
    return {"action": "open_website", "url": url, "status": "done"}

@needle.tool
def search_web(query: str):
    "Search the web for a query."
    url = "https://www.google.com/search?q=" + urllib.parse.quote(query)
    webbrowser.open(url)
    return {"action": "search_web", "query": query, "status": "done"}

@needle.tool
def open_app(app_name: str):
    "Open an application by name."
    os.system(f"start {app_name}")
    return {"action": "open_app", "app_name": app_name, "status": "done"}

@needle.tool
def create_note(content: str):
    "Create a note with given content (saved as a text file)."
    path = os.path.join(DOWNLOADS_PATH, "note.txt")
    with open(path, "a", encoding="utf-8") as f:
        f.write(content + "\n")
    return {"action": "create_note", "content": content, "status": "saved"}

# ---------- Simulated tools (return structured action, not yet wired to real execution) ----------

@needle.tool
def create_file(filename: str, filetype: str):
    "Create a new file with a given filename and type."
    return {"action": "create_file", "filename": filename, "filetype": filetype}

@needle.tool
def write_cell(sheet: str, cell: str, value: str):
    "Write a value into a specific cell of a spreadsheet."
    return {"action": "write_cell", "sheet": sheet, "cell": cell, "value": value}

@needle.tool
def write_range(sheet: str, start_cell: str, values: str):
    "Write multiple values starting at a cell in a spreadsheet."
    return {"action": "write_range", "sheet": sheet, "start_cell": start_cell, "values": values}

@needle.tool
def insert_row(sheet: str, row_index: int):
    "Insert a row into a spreadsheet at a given index."
    return {"action": "insert_row", "sheet": sheet, "row_index": row_index}

@needle.tool
def apply_formula(sheet: str, cell: str, formula: str):
    "Apply a formula to a cell in a spreadsheet."
    return {"action": "apply_formula", "sheet": sheet, "cell": cell, "formula": formula}

@needle.tool
def save_file(filename: str):
    "Save the current file."
    return {"action": "save_file", "filename": filename}

@needle.tool
def insert_heading(filename: str, text: str, level: int = 1):
    "Insert a heading into a document."
    return {"action": "insert_heading", "filename": filename, "text": text, "level": level}

@needle.tool
def generate_document(topic: str, filename: str):
    "Generate a full document draft on a given topic."
    return {"action": "generate_document", "topic": topic, "filename": filename}

@needle.tool
def open_email_client():
    "Open the email client."
    webbrowser.open("https://mail.google.com")
    return {"action": "open_email_client", "status": "done"}

@needle.tool
def open_file(filepath: str):
    "Open an existing file by path."
    os.startfile(filepath)
    return {"action": "open_file", "filepath": filepath}

@needle.tool
def find_file(filename: str):
    "Find a file by name."
    return {"action": "find_file", "filename": filename}

@needle.tool
def open_folder(path: str):
    "Open a folder by path."
    os.startfile(path)
    return {"action": "open_folder", "path": path}

@needle.tool
def delete_file(filepath: str):
    "Delete a file by path."
    return {"action": "delete_file", "filepath": filepath}

@needle.tool
def rename_file(old_name: str, new_name: str):
    "Rename a file."
    return {"action": "rename_file", "old_name": old_name, "new_name": new_name}

@needle.tool
def close_window(app_name: str):
    "Close an application window."
    return {"action": "close_window", "app_name": app_name}

@needle.tool
def minimize_window(app_name: str):
    "Minimize an application window."
    return {"action": "minimize_window", "app_name": app_name}

@needle.tool
def set_reminder(text: str, time: str):
    "Set a reminder with text and time."
    return {"action": "set_reminder", "text": text, "time": time}

@needle.tool
def type_text(target: str, text: str):
    "Type text into a target field or app."
    return {"action": "type_text", "target": target, "text": text}

@needle.tool
def click_element(target: str):
    "Click a UI element identified by its visible label and type."
    return {"action": "click_element", "target": target}


ALL_TOOLS = [
    write_text, open_browser, search_youtube, play_video, compose_email,
    open_website, search_web, open_app, create_note,
    create_file, write_cell, write_range, insert_row, apply_formula, save_file,
    insert_heading, generate_document, open_email_client, open_file, find_file,
    open_folder, delete_file, rename_file, close_window, minimize_window,
    set_reminder, type_text, click_element,
]

agent = needle.Needle(weights="my_agent.cact", tools=ALL_TOOLS)

if __name__ == "__main__":
    query = input("What should I do? ")
    result = agent.run(query)
    print(result["results"])