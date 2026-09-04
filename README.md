# Padhaya Bahut, Ab Jawab Do! - Teachers' Day Quiz Challenge 2026

An interactive Kaun Banega Crorepati (KBC) style quiz web application designed specifically for **Teachers' Day** celebrations, live auditorium hosting, TV broadcasting, and classroom entertainment.

---

## 🌟 Key Features

1. **Festive Theme & Design Aesthetics**:
   - Deep blackboard dark emerald green background (`#0d1f18`), festive marigold/saffron (`#f8cf8a`, `#eda23d`) toran garlands, glassmorphism cards, and KBC-style metallic option buttons.
2. **KBC Mechanics & Safe Checkpoints (Padav)**:
   - **15 Questions Ladder**: Ranging from ₹1,000 up to ₹1 Crore.
   - **Safe Checkpoints**:
     - **1st Padav (Q5)**: ₹10,000
     - **2nd Padav (Q10)**: ₹3,20,000
3. **4 Authentic Lifelines**:
   - ⚡ **50:50**: Eliminates 2 incorrect options.
   - 📊 **Audience Poll**: Animated percentage breakdown graph modal.
   - 👨‍🏫 **Ask Expert / Teacher**: Expert advice modal with percentage confidence.
   - 🔄 **Question Swap**: Replaces current question with a fresh unused question from the difficulty pool.
4. **Web Audio API Sound Engine**:
   - Zero external audio files required! Synthesizes authentic game audio in real time: timer ticks, option lock suspense, correct chime, wrong buzzer, lifeline sparkle, and victory fanfare.
5. **Auditorium TV Screen Mode**:
   - Synchronized live broadcast view for TV screens and projectors. Allows the host to push active games live.
6. **Question Lock & Race-Condition Guard**:
   - Uses Google Apps Script `LockService` and stable `Question ID` assignments so no two contestants receive duplicate questions during an event.
7. **Dual-Engine Bridge**:
   - Automatically detects Google Apps Script backend (`google.script.run`) or runs locally with built-in mock logic.

---

## 🚀 How to Run Locally

You can preview and test the web app directly on your machine without Google Sheets:

1. Open a terminal in `C:\Users\New\.gemini\antigravity\scratch\teachers-day-quiz`:
   ```bash
   npm install
   npm run dev
   ```
2. Open the URL shown in your terminal (e.g. `http://localhost:5173`) in your web browser.

---

## 🌐 How to Deploy on Google Apps Script (Web App)

1. Open [Google Apps Script](https://script.google.com/).
2. Create a **New Project** and name it `Teachers Day Quiz 2026`.
3. Paste the contents of `Code.gs` into `Code.gs` in your Apps Script editor.
4. Update `SPREADSHEET_ID` in `Code.gs` with your Google Spreadsheet ID (or keep the default provided).
5. Create a new HTML file in Apps Script named `index` (`index.html`) and paste the complete contents of `index.html`.
6. Click **Deploy** -> **New deployment**.
7. Select Type: **Web app**.
8. Set Configuration:
   - **Execute as**: `Me`
   - **Who has access**: `Anyone`
9. Click **Deploy** and authorize the script permissions.
10. Copy the Web App URL and open it in any browser or display on the auditorium screen!

---

## 🔑 Admin PIN

- Default Admin PIN: `1234`
- You can change the Admin PIN via the Admin Panel tab or directly in the `Admin` sheet cell `B2`.
