# Command-Center

A Google Sheets + Apps Script toolkit for managing Google Classroom student rosters, rubrics, and assignments—with a built-in AI Copilot powered by Gemini for natural language commands.

---

## ✨ Features

- **AI Copilot Sidebar**  
  Use natural language (like “create 3 assignments for this week”) to automate Google Classroom via Gemini API.
- **Command Center Sheet**  
  Central dashboard for tracking student scores, updating rubrics, and managing rosters.
- **Student Sheet Generator**  
  Quickly create and update individual student sheets using live Classroom data.
- **Flexible Rubric Builder**  
  Customizable rubric/score templates for week-by-week assessment.
- **Google Classroom Integration**  
  Interacts directly with Classroom API for rosters, assignments, and more.
- **Intuitive UI**  
  All features accessible from the Google Sheets menu or sidebar.

---

## 🚀 Getting Started

1. **Add the Script to Your Google Sheet**
   - Open your Google Sheet.
   - Go to `Extensions > Apps Script`.
   - Copy all `.gs` and `.html` files from this repo (e.g., `code.gs`, `commandcenter.gs`, `aicopliot.gs`, `sidebar.html`) into the Apps Script editor.

2. **Enable Google Classroom API**
   - In the Apps Script editor, click the “+” next to “Services”.
   - Add “Google Classroom API”.

3. **Set Script Properties**
   - Go to `Project Settings > Script Properties`.
   - Add your `course_id` and `gemini_api_key` (see below).

4. **Reload the Sheet**
   - Refresh your Google Sheet and use the new “Classroom Tools” menu.

5. **First-Time Setup**
   - Use the menu to select your active course, generate student sheets, and set up your Command Center.

---

## 🤖 AI Copilot

- Open the sidebar via “Classroom Tools > Commander Assistant”.
- Enter a natural language command (e.g., “list all assignments”, “delete assignment ‘Quiz 1’”).
- The AI will interpret your command, call the Gemini API, and execute actions via Google Classroom.

---

## 🛠️ Script Properties

You must set these Script Properties for full functionality:

| Property          | Description                                         |
|-------------------|-----------------------------------------------------|
| `course_id`       | Your Google Classroom course ID                     |
| `gemini_api_key`  | Your Gemini API key for AI Copilot                  |

---

## 📂 File Overview

- `code.gs`  
  Global config, UI menus, and sidebar setup.
- `commandcenter.gs`  
  Logic for student sheet creation, scores, rubrics, and command center.
- `aicopliot.gs`  
  Core AI Copilot—Gemini API integration and natural language interpreter.
- `sidebar.html`  
  Sidebar UI for entering and running AI commands.

---

## 🙌 Contributing

Pull requests are welcome! Please open an issue first to discuss major changes.

---

## 📄 License

No license specified yet.

---

## 📝 Notes

- This project is written almost entirely in JavaScript (94%) with some HTML (6%).
- Designed for educators and Google Workspace users who want to supercharge Google Classroom with automation and AI.

---

**Questions, suggestions, or want to say hi?**  
Open an issue or reach out on [GitHub](https://github.com/ejspttt).
