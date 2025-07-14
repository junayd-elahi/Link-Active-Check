# 🔗 Excel Link Activity Checker

A Python tool to check the status of URLs inside Excel files. Useful for verifying hundreds of hyperlinks in marketing sheets, stock registers, or SEO audits.

---

## 📌 Features

- Reads Excel files and scans URLs in column B
- Writes "Link is Active" or "Link is Inactive" in column C
- Handles redirects (301/302) and broken links
- Skips empty cells automatically
- Multithreaded for fast checking
- Supports unlimited files via config list

---

## 🛠 Tech Stack

- **Python 3.8+**
- `openpyxl` – Excel file handling  
- `requests` – HTTP/HTTPS checking  
- `tqdm` – Progress bars  
- `concurrent.futures` – Threading

---

## 📂 Folder Structure

```
Link-Active-Check/
├── link_checker_commented.py   # Cleaned & commented script
├── sample_input.xlsx           # Excel file with test URLs
└── README.md
```

---

## ▶️ How to Use

1. Clone the repository:
   ```bash
   git clone https://github.com/junayd-elahi/Link-Active-Check.git
   cd Link-Active-Check
   ```

2. Install required libraries:
   ```bash
   pip install openpyxl requests tqdm
   ```

3. Replace file paths in `configurations = [...]`:
   ```python
   configurations = [
       {'file_path': r"path\\to\\your\\excel_file1.xlsx"},
       {'file_path': r"path\\to\\your\\excel_file2.xlsx"},
   ]
   ```

4. Run the script:
   ```bash
   python link_checker_commented.py
   ```

---

## 🧪 Excel File Format

- **Column B (2)** → URL to be checked
- **Column C (3)** → Script writes status result
- Script scans **Sheet 2 and 3** only (index 1 and 2)

---

## 🚀 Example Output

| Link                        | Status          |
|-----------------------------|------------------|
| https://www.google.com      | Link is Active   |
| https://badlink.fakeurl     | Link is Inactive |

---

## 📫 Contact

📧 junayd.elahi124@gmail.com  
🔗 [GitHub](https://github.com/junayd-elahi)  
🔗 [LinkedIn](https://www.linkedin.com/in/junayd-elahi-2029b9213/)
