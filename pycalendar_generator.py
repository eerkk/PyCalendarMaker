import calendar
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import pandas as pd
import textwrap  # Now it will be actively used!

# Use Garamond font
plt.rcParams["font.family"] = "Garamond"

def load_data(year, file_path="pycalendar_data.xlsx"):
    """Loads public holidays, special days, birthdays, and leave days from an Excel file."""
    try:
        xls = pd.ExcelFile(file_path)
        
        def read_sheet(sheet_name):
            """Reads a sheet if it exists, otherwise returns an empty DataFrame."""
            if sheet_name in xls.sheet_names:
                return pd.read_excel(xls, sheet_name)
            else:
                return pd.DataFrame()

        holidays_df = read_sheet("Public_Holidays")
        special_days_df = read_sheet("Special_Days")
        birthdays_df = read_sheet("Birthdays")
        leaves_df = read_sheet("Leave_Days")
        
        holidays = {row['Date']: (row['Description'], row['Grid Color'], row['Text Color']) for _, row in holidays_df.iterrows()} if not holidays_df.empty else {}
        special_days = {row['Date']: (row['Description'], row['Grid Color'], row['Text Color']) for _, row in special_days_df.iterrows()} if not special_days_df.empty else {}
        leaves = {row['Date']: (row['Description'], row['Grid Color'], row['Text Color']) for _, row in leaves_df.iterrows()} if not leaves_df.empty else {}
        birthdays = {row['Date'][:5]: (f"{row['Description']}", row['Grid Color'], row['Text Color']) for _, row in birthdays_df.iterrows()} if not birthdays_df.empty else {}
        
        return holidays, special_days, birthdays, leaves
    except FileNotFoundError:
        print("Excel file not found! Please check the file location.")
        return {}, {}, {}, {}


def create_calendar(year, holidays, special_days, birthdays, leaves, filename="calendar_A4.pdf", page_size="A4"):
    """Creates a calendar in the specified format (A4 or horizontal A0)."""
    if page_size == "A4":
        figsize = (11.7, 8.3)  # A4 Landscape
    elif page_size == "A0":
        figsize = (60, 33.1)  # A0 Landscape
    else:
        print("Invalid page size!")
        return
    
    with PdfPages(filename) as pdf:
        if page_size == "A4":
            for month in range(1, 13):
                fig, ax = plt.subplots(figsize=figsize)
                create_monthly_calendar(ax, month, year, holidays, special_days, birthdays, leaves, scale=1.0)
                pdf.savefig(fig, bbox_inches='tight')
                plt.close()
        else:
            fig, axes = plt.subplots(3, 4, figsize=figsize)  # 3 rows, 4 columns
            plt.subplots_adjust(wspace=0.05, hspace=0.2)  # Reduce empty spaces
            
            for month in range(1, 13):
                row, col = (month - 1) // 4, (month - 1) % 4
                create_monthly_calendar(axes[row, col], month, year, holidays, special_days, birthdays, leaves, scale=1.3)
            
            pdf.savefig(fig, bbox_inches='tight')
            plt.close()

    print(f"PDF successfully created: {filename}")

def create_monthly_calendar(ax, month, year, holidays, special_days, birthdays, leaves, scale=1.0):
    """Generates a monthly calendar layout."""
    ax.set_xlim(0, 7)
    ax.set_ylim(6, 0)
    ax.axis("off")
    
    month_name = calendar.month_name[month].capitalize()
    ax.text(3.5, -0.5, f"{month_name} {year}", fontsize=20 * scale, ha="center")
    
    days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    for i, day in enumerate(days):
        ax.text(i + 0.5, 0.5, day, fontsize=14 * scale, ha="center", fontweight="bold", 
                color="gray" if day in ["Sat", "Sun"] else "black")
    
    month_calendar = calendar.monthcalendar(year, month)
    for row_idx, week in enumerate(month_calendar):
        for col_idx, day in enumerate(week):
            if day != 0:
                date_str = f"{day:02d}.{month:02d}.{year}"
                birthday_key = f"{day:02d}.{month:02d}"
                grid_color = "white"
                text_color = "black"
                text = ""

                # Assign text and colors based on event type
                if date_str in holidays:
                    text, grid_color, text_color = holidays[date_str]
                elif date_str in special_days:
                    text, grid_color, text_color = special_days[date_str]
                elif birthday_key in birthdays:
                    text, grid_color, text_color = birthdays[birthday_key]
                elif date_str in leaves:
                    text, grid_color, text_color = leaves[date_str]
                elif col_idx >= 5:
                    grid_color = "#808080"

                # Draw grid cell
                ax.add_patch(plt.Rectangle((col_idx, row_idx + 1), 1, 1, color=grid_color, alpha=0.5))
                
                # Add the date number
                ax.text(col_idx + 0.9, row_idx + 1.1, str(day), fontsize=10 * scale, ha="right", va="top", 
                        color=text_color, fontweight="bold")

                # Use textwrap to fit long descriptions
                if text:
                    wrapped_text = "\n".join(textwrap.wrap(text, width=12))  # Wrap text inside cells
                    ax.text(col_idx + 0.5, row_idx + 1.5, wrapped_text, fontsize=8 * scale, ha="center", va="center", 
                            color=text_color)

    # Draw grid lines
    for x in range(8):
        ax.plot([x, x], [0, 6], color="gray", linewidth=0.2)
    for y in range(7):
        ax.plot([0, 7], [y, y], color="gray", linewidth=0.2)


if __name__ == "__main__":
    year = int(input("Enter the calendar year: "))
    holidays, special_days, birthdays, leaves = load_data(year)
    
    create_calendar(year, holidays, special_days, birthdays, leaves, filename=f"calendar_A4_{year}.pdf", page_size="A4")
    create_calendar(year, holidays, special_days, birthdays, leaves, filename=f"calendar_A0_{year}.pdf", page_size="A0")
