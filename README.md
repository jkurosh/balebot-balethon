# 🤖 Bale Appointment & Daily Quiz Bot

A fully-featured Python chatbot backend built with **Balethon** for the Bale messenger platform.  
This bot provides **appointment scheduling**, **user management**, and a **daily question/quiz system**, all powered by an SQLite database.

---

## 🚀 Features

- ✅ User registration and tracking  
- 📅 Appointment booking system  
- 🧹 Automatic cleanup of expired appointments  
- ❓ Daily Question / Quiz system  
- 🗂 SQLite database integration  
- 📆 Jalali (Persian) date support using `khayyam`  
- 🔒 Persistent storage for user answers and appointments  
- 🛠 Automatic database initialization on startup  

---

## 🏗 Project Structure

```
.
├── main.py        # Core bot logic and database handling
├── database.db    # SQLite database (auto-generated)
└── screen/        # Directory for stored screenshots/files
```

> The database and required tables are automatically created when the bot runs for the first time.

---

## 🛠 Requirements

Make sure you have Python 3.9+ installed.

Install dependencies:

```bash
pip install balethon khayyam
```

Or create a `requirements.txt` file:

```
balethon
khayyam
```

Then run:

```bash
pip install -r requirements.txt
```

---

## ⚙️ Configuration

Before running the bot:

1. Set your Bale bot token.
2. Adjust file paths if needed (avoid hardcoded absolute paths).
3. Make sure the `screen` directory exists or update the path in the code.

### 🔐 Security Recommendation

Do **NOT** hardcode your bot token inside `main.py`.

Instead, use environment variables:

```python
import os

TOKEN = os.getenv("BOT_TOKEN")
```

Then set it in your system:

**Windows:**
```bash
set BOT_TOKEN=your_token_here
```

**Linux / Mac:**
```bash
export BOT_TOKEN=your_token_here
```

---

## ▶️ Running the Bot

```bash
python main.py
```

If everything is configured correctly, the bot will start and connect to Bale.

---

## 🗄 Database Schema Overview

The bot uses SQLite and includes tables for:

- `users`
- `appointments`
- `questions`
- `user_answers`
- `question_of_day_state`

All tables are automatically created if they do not exist.

---

## 🧠 How It Works

- Users interact with the bot.
- The bot checks registration status.
- Users can book appointment slots.
- Expired slots are cleaned automatically.
- The system manages a "Question of the Day" and tracks user responses.
- All data is persisted in SQLite.

---

## 📜 License

This project is licensed under the MIT License.  
Feel free to use, modify, and distribute it.

---

## 💡 Future Improvements

- Admin dashboard
- Better scheduling UI
- Docker support
- Web-based control panel
- Multi-language support
