"""
Create daily income spreadsheet for Mom
"""

from datetime import datetime
from pathlib import Path

import pandas as pd


def create_date(month: int, year: int) -> pd.DatetimeIndex:
    """
    Create a DateTimeindex of dates for the specified month and year
    """

    month_start = datetime(year, month, 1)
    dates = pd.date_range(
        month_start, month_start + pd.offsets.MonthBegin(1), inclusive="left"
    )

    return dates


def create_weekday(dates: pd.DatetimeIndex) -> pd.Series:
    """
    Create a series of weekdays based on specified dates
    """

    DAYS_OF_WEEK = {
        0: "Mon",
        1: "Tue",
        2: "Wed",
        3: "Thu",
        4: "Fri",
        5: "Sat",
        6: "Sun",
    }

    s = dates.to_series()

    weekdays = s.dt.dayofweek

    return weekdays.map(DAYS_OF_WEEK)


def _get_month_name(month_num: int) -> str:
    month_name = datetime.strptime(str(month_num), "%m").strftime("%b")

    return month_name


def run(month: int, year: int, output_path: Path):

    days = create_date(month, year)
    weekdays = create_weekday(days)

    month_name = _get_month_name(month)

    dates_df = pd.DataFrame({"Date": days.values, "weekdays": weekdays.values})
    weekdays_list = dates_df["weekdays"].to_list()
    first_sunday = weekdays_list.index("Sun")

    with pd.ExcelWriter(
        output_path / f"Daily Income {month_name}.{year}.xlsx", engine="xlsxwriter"
    ) as writer:
        dates_df.to_excel(writer, sheet_name="Sheet1", startrow=1, index=False)
        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]

        TITLE_FORMAT = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 24,
                "align": "center",
                "right": 0,
            }
        )

        HEADER_FORMAT = workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 14,
                "align": "center",
                "valign": "vcenter",
                "border": 1,
            }
        )

        DATE_FORMAT = workbook.add_format(
            {
                "font_name": "Arial",
                "font_size": 14,
                "align": "center",
                "valign": "vcenter",
                "num_format": "m/d",
                "border": 1,
            }
        )

        worksheet.set_default_row(19.5)
        worksheet.set_row(0, 29)
        worksheet.set_row(1, 23)

        # according to https://stackoverflow.com/questions/61392660/xlsxwriter-custom-column-formatting-issue
        # set_column doesn't work with date columns, loop workaround
        (max_row, _) = dates_df.shape

        for row in range(0, max_row):
            worksheet.write(row + 2, 0, dates_df.iloc[row, 0], DATE_FORMAT)

        worksheet.set_column("A:B", 5, DATE_FORMAT)
        worksheet.set_column("C:F", 10.33, workbook.add_format({"border": 1}))
        worksheet.set_column("G:G", 21.67, workbook.add_format({"border": 1}))

        worksheet.merge_range(
            "A1:G1", f"Daily Income {month_name}.{year}", TITLE_FORMAT
        )

        worksheet.merge_range("A2:B2", "Date", HEADER_FORMAT)
        worksheet.write("C2", "Cash", HEADER_FORMAT)
        worksheet.write("D2", "Credit", HEADER_FORMAT)
        worksheet.write("E2", "Total", HEADER_FORMAT)
        worksheet.write("F2", "Tip", HEADER_FORMAT)
        worksheet.write("G2", "Note", HEADER_FORMAT)

        if first_sunday >= 4:
            COLOR_BG = True
        else:
            COLOR_BG = False

        if first_sunday > 0:
            worksheet.merge_range(
                2, 6, first_sunday + 1, 6, "", workbook.add_format({"right": 1})
            )

        for row in range(first_sunday + 2, max_row, 7):
            if COLOR_BG:
                # merge_range 0 indexed
                worksheet.merge_range(
                    row,
                    6,
                    min(row + 6, max_row + 1),
                    6,
                    "",
                    workbook.add_format({"bg_color": "#F2F2F2", "right": 1, "top": 1}),
                )
                # conditional_format matches row number in sheet that's why row + 1
                worksheet.conditional_format(
                    f"A{row+1}:F{min(row + 7, max_row + 1)}",
                    {
                        "type": "formula",
                        "criteria": f'=$G{row}=""',
                        "format": workbook.add_format({"bg_color": "#F2F2F2"}),
                    },
                )
                COLOR_BG = False
            else:
                worksheet.merge_range(row, 6, min(row + 6, max_row + 1), 6, "")
                COLOR_BG = True

        worksheet.print_area(f"A1:G{max_row+4}")
