import pandas as pd

# File path for the Excel file
file_path = "pycalendar_data.xlsx"

COLORS = {
    "Weekend": "#D3D3D3",  # Light Gray (neutral but visible)
    "Public_Holiday": "#FF5C5C",  # Soft Red (attention-grabbing but not aggressive)
    "Special_Day": "#FFD700",  # Gold (celebratory & warm)
    "Birthday": "#87CEFA",  # Sky Blue (cheerful and friendly)
    "Leave_Day": "#98FB98"  # Pastel Green (relaxed & vacation-like)
}

# Public Holidays
public_holidays = pd.DataFrame({
    "Date": ["01.01.2025", "23.04.2025", "19.05.2025", "30.08.2025", "29.10.2025"],
    "Description": ["New Year's Day", "National Sovereignty and Children's Day", 
                    "Atatürk Memorial, Youth and Sports Day", "Victory Day", "Republic Day"],
    "Grid Color": [COLORS["Public_Holiday"]] * 5,
    "Text Color": ["#000000"] * 5  # Black text
})

# Special Days
special_days = pd.DataFrame({
    "Date": ["14.02.2025", "08.03.2025", "05.06.2025", "01.09.2025"],
    "Description": ["Valentine's Day", "International Women's Day", "Environment Day", "World Peace Day"],
    "Grid Color": [COLORS["Special_Day"]] * 4,
    "Text Color": ["#000000"] * 4
})

# Birthdays
birthdays = pd.DataFrame({
    "Date": ["10.04.2025", "20.07.2025", "15.09.2025"],
    "Name": ["Ahmet", "Buse", "Mehmet"],
    "Grid Color": [COLORS["Birthday"]] * 3,
    "Text Color": ["#000000"] * 3
})

# Add "Birthday" to the names
birthdays["Description"] = birthdays["Name"] + "'s Birthday"
birthdays = birthdays.drop(columns=["Name"])  # Remove the "Name" column

# Leave Days
leave_days = pd.DataFrame({
    "Date": ["12.06.2025", "25.12.2025"],
    "Description": ["Annual Leave", "Christmas Holiday"],
    "Grid Color": [COLORS["Leave_Day"]] * 2,
    "Text Color": ["#000000"] * 2
})

# Save all sheets to an Excel file
with pd.ExcelWriter(file_path) as writer:
    public_holidays.to_excel(writer, sheet_name="Public_Holidays", index=False)
    special_days.to_excel(writer, sheet_name="Special_Days", index=False)
    birthdays.to_excel(writer, sheet_name="Birthdays", index=False)
    leave_days.to_excel(writer, sheet_name="Leave_Days", index=False)

print(f"Excel file successfully created: {file_path}")
