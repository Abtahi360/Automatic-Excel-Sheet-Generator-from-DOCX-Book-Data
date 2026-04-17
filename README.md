## 🎯 Objective

The goal of this project is to automate the extraction of structured information from a Bangla document and store it into clean Excel files. The system identifies chapters, bold sub-sections, and numbered hadith entries from the `.docx` file and organizes them into three separate `.xlsx` files with sequential IDs. It keeps the original text unchanged while removing blank lines and unnecessary empty cells. This reduces manual work, prevents duplicate effort, and makes the data easier to use for further analysis or processing.

---

## ⚙️ Features

* ✅ Automatic **Chapter extraction**
* ✅ Smart detection of **Bold Sub-sections**
* ✅ Accurate **Hadith identification** (`[১], [২], ...`)
* ✅ Removes **blank lines & noise**
* ✅ Keeps **original text unchanged**
* ✅ Generates **3 Excel files instantly**
* ✅ Clean structure with **auto ID generation**

---

## 🧠 How It Works

1. 📂 Load `.docx` file
2. 🔍 Detect patterns:

   * `অধ্যায়:` → Chapter
   * Bold + spacing → Sub-section
   * `[number]` → Hadith
3. 🧹 Clean data (remove blank lines)
4. 📊 Export to Excel

---

## 📁 Output Files

### 📄 Chapters.xlsx

| id | name                                |
| -- | ----------------------------------- |
| 1  | অধ্যায়: পিতা-মাতার সাথে সদ্ব্যবহার |

### 📄 Subsections.xlsx

| id | name                                                              |
| -- | ----------------------------------------------------------------- |
| 1  | আমি মানুষকে তার পিতা-মাতার সাথে সদ্ব্যবহারের নির্দেশ প্রদান করেছি |

### 📄 Hadith.xlsx

| id | hadith              |
| -- | ------------------- |
| 1  | Full hadith text... |

---

## 🛠️ Tech Stack

* `Python`
* `python-docx`
* `pandas`
* `openpyxl`

---

## 🚀 How to Run

### 1️⃣ Install dependencies

```bash
pip install python-docx pandas openpyxl
```

### 2️⃣ Run notebook

Open and run:

```
automatic excel sheet generat.ipynb
```

### 3️⃣ Get output

You will get:

* `chapters.xlsx`
* `subsections.xlsx`
* `hadith.xlsx`

---

## 🌍 Real-Life Use Cases & Impact

This project is not just academic, it solves real problems 👇

### 📚 For Students & Researchers

* Quickly convert books into structured datasets
* Save hours of manual typing
* Prepare data for research or ML models

### 🕌 For Islamic Content Management

* Extract hadith collections into searchable format
* Build apps/websites using structured religious data
* Organize large texts easily

### 📊 For Data Entry & Office Work

* Replace repetitive manual Excel work
* Avoid human errors and duplication
* Handle large documents efficiently

### 🤖 For Developers & ML Engineers

* Use as preprocessing step for NLP tasks
* Convert unstructured text → structured dataset
* Build training data easily

### 💼 Real Problem It Solves

👉 Manual data entry from books to Excel is slow, boring, and error-prone.
👉 This tool automates the whole process in seconds.

---

## 💡 Why This Project Matters

* ⏳ Saves time
* 🎯 Improves accuracy
* 📈 Makes data usable
* 🤝 Reduces repetitive work

---

## 🤝 Contribution

Feel free to fork this repo and improve it 🚀

---

## ⭐ Support

If you find this useful, give it a ⭐ on GitHub!
