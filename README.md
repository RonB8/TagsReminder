# TagsReminder
Software designed for front desk computers at the reception desk sends a WhatsApp reminder to employees to return a badge they received.



https://github.com/user-attachments/assets/162da63d-dca8-4709-be6b-fd5aef336924


---

# 🛡️ TagsReminder

**Automated WhatsApp Reminders for Workplace Badge Returns**

`TagsReminder` is a Python-based utility designed for reception and front desk operations. It automatically scans daily badge tracking logs and sends WhatsApp reminders to employees who forgot to return their temporary badges.

---

## 🚀 Features

* **Automated Scanning**: Automatically reads Excel/CSV logs to identify unreturned badges.
* **Smart Date Matching**: Automatically looks for the sheet corresponding to "yesterday's" date in various formats (e.g., `DD.MM.YY`, `DD/MM/YYYY`).
* **WhatsApp Integration**: Uses `pywhatkit` to send instant, personalized messages via WhatsApp Web.
* **Scheduling**: Built-in scheduler to run the check at a specific time every day.
* **Easy Setup**: GUI-based configuration for first-time users to select the data file and set the run time.

---

## 🛠️ Installation

### Prerequisites

* Python 3.x
* A default web browser logged into **WhatsApp Web**.

### Setup

1. **Clone the repository**:
```bash
git clone https://github.com/ronb8/tagsreminder.git
cd tagsreminder

```

2.  **Create and activate a virtual environment (Recommended):**

    ```bash
    # Windows
    python -m venv venv
    venv\Scripts\activate

    # macOS/Linux
    python3 -m venv venv
    source venv/bin/activate
    ```


3. **Install dependencies**:
```bash
pip install -r requirements.txt

```

*Dependencies include: `pandas`, `openpyxl`, `schedule`, and `pywhatkit*`.

---

## 📖 How to Use

1. **Run the application**:
```bash
python main.py

```


2. **Configuration**:
* On the first run, a window will prompt you to select your **Badge Tracking Excel file**.
* Set the **Daily Run Time** (e.g., `08:00`) when the script should send reminders.


3. **Operation**:
* The script will remain active in the background.
* Every day at the scheduled time, it scans the sheet from the previous day.
* If an employee hasn't marked their badge as "Returned" (`הוחזר` or `כן`), they will receive a message.



---

## 📝 Data Format Requirement

The Excel/CSV file should contain columns such as:

* **Name**: `שם מלא` or `שם`
* **Phone**: `פלאפון` or `מס טלפון`
* **Status**: `הוחזר/לא הוחזר` (The system checks if this is empty or not "כן"/"הוחזר")
* **Badge ID**: `תג`

*Example:*
| שעה | שם מלא | פלאפון | תג | הוחזר/לא הוחזר |
| :--- | :--- | :--- | :--- | :--- |
| 08:01 | ישראל ישראלי | 051-1111111 | 53 | |

---

## 🤖 Automated Message Example

> "שלום [Name], נראה כי שכחת להחזיר את תג מספר [Badge Number] שלקחת אתמול. אנא דאג/י להחזיר אותו בהקדם. תודה ויום טוב!"

---


*Developed by Ron Boaron*
