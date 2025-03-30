import calendar
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import pandas as pd
import textwrap
from collections import defaultdict  # <-- YENİ

# Use Garamond font
plt.rcParams["font.family"] = "Garamond"

def load_data(year, file_path="pycalendar_data.xlsx"):
    """Loads public holidays, special days, birthdays, and leave days from an Excel file."""
    try:
        xls = pd.ExcelFile(file_path)

        def read_sheet(sheet_name):
            if sheet_name in xls.sheet_names:
                return pd.read_excel(xls, sheet_name)
            else:
                return pd.DataFrame()

        holidays_df = read_sheet("Public_Holidays")
        special_days_df = read_sheet("Special_Days")
        birthdays_df = read_sheet("Birthdays")
        leaves_df = read_sheet("Leave_Days")

        # Multiple registrations in one day 
        holidays = defaultdict(list)
        special_days = defaultdict(list)
        leaves = defaultdict(list)
        birthdays = defaultdict(list)

        for _, row in holidays_df.iterrows():
            holidays[row['Date']].append((row['Description'], row['Grid Color'], row['Text Color']))

        for _, row in special_days_df.iterrows():
            special_days[row['Date']].append((row['Description'], row['Grid Color'], row['Text Color']))

        for _, row in leaves_df.iterrows():
            leaves[row['Date']].append((row['Description'], row['Grid Color'], row['Text Color']))

        for _, row in birthdays_df.iterrows():
            key = row['Date'][:5]
            birthdays[key].append((row['Description'], row['Grid Color'], row['Text Color']))

        return holidays, special_days, birthdays, leaves
    except FileNotFoundError:
        print("Excel file not found! Please check the file location.")
        return {}, {}, {}, {}

def create_calendar(year, holidays, special_days, birthdays, leaves, filename="calendar_A4.pdf", page_size="A4"):
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
            fig, axes = plt.subplots(3, 4, figsize=figsize)
            plt.subplots_adjust(wspace=0.05, hspace=0.2)
            for month in range(1, 13):
                row, col = (month - 1) // 4, (month - 1) % 4
                create_monthly_calendar(axes[row, col], month, year, holidays, special_days, birthdays, leaves, scale=1.3)
            pdf.savefig(fig, bbox_inches='tight')
            plt.close()

    print(f"PDF successfully created: {filename}")

def create_monthly_calendar(ax, month, year, holidays, special_days, birthdays, leaves, scale=1.0):
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
                event_lines = []

                # Collect all activities during the day
                if date_str in holidays:
                    for desc, gc, tc in holidays[date_str]:
                        event_lines.append((desc, gc, tc))

                if date_str in special_days:
                    for desc, gc, tc in special_days[date_str]:
                        event_lines.append((desc, gc, tc))

                if birthday_key in birthdays:
                    for desc, gc, tc in birthdays[birthday_key]:
                        event_lines.append((desc, gc, tc))

                if date_str in leaves:
                    for desc, gc, tc in leaves[date_str]:
                        event_lines.append((desc, gc, tc))

                # Number of lines and height adjustment
                num_lines = max(1, len(event_lines))
                line_height = 1.0 / num_lines

                # Draw a background box for each row
                for idx in range(num_lines):
                    y_bottom = row_idx + 1 + idx * line_height
                    color = event_lines[idx][1] if idx < len(event_lines) else "white"
                    ax.add_patch(plt.Rectangle((col_idx, y_bottom), 1, line_height, color=color, alpha=0.5))

                # If there are no events and the day falls on a weekend, make the box gray
                if not event_lines and col_idx >= 5:
                    ax.add_patch(plt.Rectangle((col_idx, row_idx + 1), 1, 1, color="#808080", alpha=0.5))

                # Place the day number in the top left corner of the box
                ax.text(
                    col_idx + 0.05,
                    row_idx + 1.05,
                    str(day),
                    fontsize=9 * scale,
                    ha="left",
                    va="top",
                    color="black",
                    fontweight="bold"
                )

                if event_lines:  # draw if there is activity
                    num_lines = len(event_lines)
                    line_height = 1.0 / num_lines
                    font_size = max(3 * scale, 7 * scale - (num_lines - 3) * 0.5)
                    group_offset = (1.0 - num_lines * line_height) / 2

                    for idx, (text, _, tc) in enumerate(event_lines):
                        wrapped = "\n".join(textwrap.wrap(text, width=20))
                        y_center = row_idx + 1 + group_offset + (idx + 0.5) * line_height

                        ax.text(
                            col_idx + 0.5,
                            y_center,
                            wrapped,
                            fontsize=font_size,
                            ha="center",
                            va="center",
                            color=tc
                        )


    # Grid lines
    for x in range(8):
        ax.plot([x, x], [0, 6], color="gray", linewidth=0.2)
    for y in range(7):
        ax.plot([0, 7], [y, y], color="gray", linewidth=0.2)

if __name__ == "__main__":
    year = int(input("What is the calendar year?: "))
    holidays, special_days, birthdays, leaves = load_data(year)
    create_calendar(year, holidays, special_days, birthdays, leaves, filename=f"calendar_A4_{year}.pdf", page_size="A4")
    create_calendar(year, holidays, special_days, birthdays, leaves, filename=f"calendar_A0_{year}.pdf", page_size="A0")
