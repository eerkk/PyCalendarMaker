# PyCalendarMaker 🗓️

PyCalendarMaker is a **Python-based tool** that generates **A4 and A0 printable calendars** with customized events like **public holidays, birthdays, and special days**. It is designed for professional and personal use, making it easier to visualize important dates.

---

## 🚀 Features

✔️ Generates **A4 and A0** calendar PDFs  
✔️ Supports **public holidays, special days, birthdays, and leave days**  
✔️ Uses **Garamond font** for a professional look  
✔️ Easily customizable **Excel-based data input**  

---

## 🔧 Installation

### 1️⃣ Clone the repository:
```bash
git clone https://github.com/eerkk/PyCalendarMaker.git
cd PyCalendarMaker
```

### 2️⃣ Install dependencies:
```bash
pip install -r requirements.txt
```

---

## 📅 Generate the Excel file:
Before running the calendar generator, you need to create the **pycalendar_data.xlsx** file that contains public holidays, birthdays, and special events. This will create a file called **pycalendar_data.xlsx** in your project folder. You can modify this file to add custom events, birthdays, and holidays.

Run the following script to generate the Excel file:
```bash
python pycalendar_data.py
```

### 4️⃣ Run the calendar generator script:
```bash
python pycalendar_generator.py
```

### 5️⃣ Enter the desired calendar year when prompted.

---

## 📦 Dependencies
This project requires the following Python packages:

`matplotlib` `pandas` `openpyxl`

You can install all dependencies using:
```bash
pip install -r requirements.txt
```

---

## 📖 How to Use
1️⃣ Generate the Excel file first using **pycalendar_data.py** (Step 3).  
2️⃣ Modify **pycalendar_data.xlsx** to include your own events.  
3️⃣ Run **pycalendar_generator.py** to generate **A4 and A0 calendar PDFs**.  
4️⃣ Open the generated **calendar_A4_YEAR.pdf** and **calendar_A0_YEAR.pdf** files.  
5️⃣ Print or share your custom calendar!  

---

## 🤝 Contributing
Feel free to fork this repository and submit pull requests with improvements!  
Found a bug? Open an issue.  
Want a new feature? Let's discuss it!  

---

## 📜 License
This project is licensed under the **MIT License**.
