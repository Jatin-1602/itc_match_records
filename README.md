# Compare Invoices Script

## Overview
- Reads the excel file containing invoices records.
- Generates new sheets in the same excel file with the matched, not_in_books, next_fy and prev_fy records.

---

## Setup

### 1. Clone the repo
```sh
git clone https://github.com/Jatin-1602/itc_match_records.git
cd itc_match_records/
```

### 2. Create virtual env
```sh
python -m venv venv
.\venv\Scripts\activate
```

### 3. Install the packages
```sh
pip install -r requirements.txt
```

### 4. Run the script
```sh
python main.py
```

---

### Note
- Make sure the excel file exists in the `data/` directory