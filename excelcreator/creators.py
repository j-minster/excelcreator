import logging

import pandas as pd
import xlsxwriter

from xlsxwriter.utility import xl_rowcol_to_cell
from .utils import (
    NestedDict,
    df_to_dict,
    get_groups,
    get_scenarios,
    get_sheetnames,
    shorten_long_sheetnames,
    vals_are_lists,
)


def create_sheet_df(in_df: pd.DataFrame, sheetname: str) -> pd.DataFrame:
    """
    Extract entries from `in_df` for the relevant `sheetname`
    """
    majorcol = list(in_df.columns)[0]
    sub_df = in_df[in_df[majorcol].isin([sheetname])]
    return sub_df


# create individual dictionaries for sheets rather than one huge dictionary for the whole dataframe
def create_sheet_dict(in_df: pd.DataFrame, sheetname: str) -> NestedDict:
    sheetnames = get_sheetnames(in_df)
    sub_df = create_sheet_df(in_df, sheetname)
    d1 = df_to_dict(sub_df)

    idx = list(sheetnames).index(sheetname)
    logging.info(f"made '{sheetname}' dict")
    logging.info(f"(Creating {idx+1} of {len(sheetnames)} total sheets)")

    return d1


def create_format_dict(workbook: xlsxwriter.Workbook) -> dict:
    format_dict = {}
    bordercolor = "#9B9B9B"
    orangecolor = "#F7D8AA"
    greycolor = "#F0F0F0"

    format_dict["sheetname"] = workbook.add_format(
        {
            "valign": "vcenter",
            "bg_color": orangecolor,
            "font_name": "Segoe UI Light (Heading)",
            "font_size": 14,
            "bottom": 1,
            "right": 1,
            "border_color": bordercolor,
            "indent": 1,
        }
    )

    format_dict["metric"] = workbook.add_format(
        {
            "bold": True,
            "bg_color": greycolor,
            "valign": "vcenter",
            "font_name": "Segoe UI (Body)",
            "font_size": 8,
            "border": 1,
            "border_color": bordercolor,
            "indent": 1,
        }
    )

    format_dict["scenario"] = workbook.add_format(
        {
            "bold": False,
            "bg_color": greycolor,
            "border_color": bordercolor,
            "align": "centre",
            "valign": "vcentre",
            "font_name": "Segoe UI (Body)",
            "font_size": 8,
            "top": 1,
            "bottom": 1,
            "left": 0,
            "right": 0,
        }
    )
    format_dict["l_scenario"] = workbook.add_format(
        {
            "bold": False,
            "bg_color": greycolor,
            "border_color": bordercolor,
            "align": "centre",
            "valign": "vcentre",
            "font_name": "Segoe UI (Body)",
            "font_size": 8,
            "top": 1,
            "bottom": 1,
            "left": 1,
            "right": 0,
        }
    )
    format_dict["r_scenario"] = workbook.add_format(
        {
            "bold": False,
            "bg_color": greycolor,
            "border_color": bordercolor,
            "align": "centre",
            "valign": "vcentre",
            "font_name": "Segoe UI (Body)",
            "font_size": 8,
            "top": 1,
            "bottom": 1,
            "left": 0,
            "right": 1,
        }
    )

    format_dict["dropdown"] = workbook.add_format(
        {
            "bottom": 1,
            "border_color": bordercolor,
            "align": "centre",
            "valign": "vcentre",
            "font_name": "Segoe UI (Body)",
            "font_size": 8,
            "bg_color": "#DAEDF8",
            "fg_color": "#FFFFFF",
            "pattern": 16,
        }
    )

    format_dict["pformat"] = workbook.add_format(
        {
            "bottom": 1,
            "border_color": bordercolor,
            "align": "right",
            "valign": "vcentre",
            "font_name": "Segoe UI (Body)",
            "font_size": 8,
            "bg_color": "#DAEDF8",
        }
    )

    format_dict["comp"] = workbook.add_format(
        {
            "top": 1,
            "border_color": bordercolor,
            "align": "left",
            "valign": "vcentre",
            "font_name": "Segoe UI Light (Headings)",
            "font_size": 8,
            "bg_color": "#DAEDF8",
            "indent": 1,
        }
    )

    format_dict["lilcell"] = workbook.add_format(
        {
            "top": 1,
            "right": 1,
            "border_color": bordercolor,
            "align": "left",
            "valign": "vcentre",
            "font_name": "Segoe UI Light (Headings)",
            "font_size": 8,
            "bg_color": "#DAEDF8",
            "indent": 1,
        }
    )

    format_dict["r"] = workbook.add_format({"right": 1, "border_color": "#9B9B9B"})

    format_dict["group"] = workbook.add_format(
        {
            "bold": False,
            "font_name": "Segoe UI (Body)",
            "font_size": 8,
            "right": 1,
            "border_color": "#9B9B9B",
        }
    )
    format_dict["index_group"] = workbook.add_format(
        {
            "bold": True,
            "font_name": "Arial Narrow",
            "font_size": 11,
            "font_color": "blue",
            "underline": 1,
            "indent": 1,
        }
    )
    format_dict["index_header"] = workbook.add_format(
        {
            "bold": True,
            "bottom": 1,
            "font_name": "Arial Narrow",
            "font_size": 16,
            "indent": 0,
        }
    )
    format_dict["index_ul"] = workbook.add_format({"bottom": 1})
    format_dict["num"] = workbook.add_format(
        {"font_name": "Segoe UI (Body)", "font_size": 8, "num_format": "#,##0.000"}
    )
    format_dict["pct"] = workbook.add_format(
        {"font_name": "Segoe UI (Body)", "font_size": 8, "num_format": "0.0%"}
    )
    format_dict["l"] = workbook.add_format({"left": 1, "border_color": "#9B9B9B"})

    return format_dict


# function to create the top block of a sheet. Returns nothing. Just alters `workbook` in memory.
def create_header_block(
    sheetname: str,
    worksheet: xlsxwriter.worksheet.Worksheet,
    sheet_dict: NestedDict,
    workbook: xlsxwriter.workbook.Workbook,
    groupnames: list[str],
    scenarionames: list[str],
    format_dict: dict,
) -> None:
    # merge cells A1 and A2
    # write the sheet name in the merged cells, large font, orange cell
    worksheet.merge_range("A1:A2", sheetname, format_dict["sheetname"])

    # set height of sheet name cell
    worksheet.set_row_pixels(0, 24)
    worksheet.set_row_pixels(1, 24)

    # set sheetname and metrics column (A) widths
    worksheet.set_column("A:A", 34.83)

    # write 'Metric' in cell A3, bold format, make the cell vertically taller (36 px), grey background
    worksheet.write("A3", "Metric", format_dict["metric"])
    worksheet.set_row_pixels(2, 48)

    # write all scenario* names in {H...}:3
    scenario_col_offset = 6
    for offset, name in enumerate(scenarionames):
        if offset == 0:
            worksheet.write(
                2, scenario_col_offset + offset, name, format_dict["l_scenario"]
            )

        elif offset == len(scenarionames) - 1:
            worksheet.write(
                2, scenario_col_offset + offset, name, format_dict["r_scenario"]
            )

        else:
            worksheet.write(
                2, scenario_col_offset + offset, name, format_dict["scenario"]
            )

    # create dropdowns (validation)
    input_cell_1 = "B$3"
    worksheet.data_validation(
        input_cell_1,
        {
            "validate": "list",
            # 'source': scenarionames,
            "source": "=$G$3:$XFD$3",
            "input_title": "Pick a scenario",
        },
    )
    worksheet.write(input_cell_1, scenarionames[0], format_dict["dropdown"])

    input_cell_2 = "C$3"
    worksheet.data_validation(
        input_cell_2,
        {
            "validate": "list",
            # 'source': scenarionames,
            "source": "=$G$3:$XFD$3",
            "input_title": "Pick a scenario",
        },
    )
    worksheet.write(input_cell_2, scenarionames[1], format_dict["dropdown"])

    # create +/- headings
    pmcell = "D$3"
    worksheet.write(pmcell, "+/-", format_dict["pformat"])
    pcell = "E$3"
    worksheet.write(pcell, "%", format_dict["pformat"])

    # merge cells {B, C, D, E}:2
    # write 'compare loaded scenarios...' in the merged cells, blue cell
    worksheet.merge_range(
        "B2:E2", "Compare two loaded scenarios (use dropdowns)", format_dict["comp"]
    )
    worksheet.write("F2", None, format_dict["lilcell"])
    worksheet.write("F3", None, format_dict["pformat"])


def create_dynamic_block(
    worksheet: xlsxwriter.worksheet.Worksheet,
    workbook: xlsxwriter.workbook.Workbook,
    format_dict: dict,
) -> None:
    # make column F small and G zero-width
    worksheet.set_column("F:F", 2.33, format_dict["r"])


# input the data for each sheet (warning: recursion)
def create_data_rows(
    worksheet: xlsxwriter.worksheet.Worksheet,
    in_dict: NestedDict,
    workbook: xlsxwriter.workbook.Workbook,
    groupnames: list[str],
    scenarionames: list[str],
    ind_level: int,
    sheetname: str,
    format_dict: dict,
    writeIndexHeader: bool,
) -> None:
    nums_offset = 6
    global row_offset
    global index_row_offset

    if writeIndexHeader:
        index_row_offset += 1
        link_string = f"internal:{sheetname!r}!A1"
        index_sheet.write_url(index_row_offset, 1, link_string)
        index_sheet.write(index_row_offset, 1, sheetname, format_dict["index_header"])
        index_row_offset += 1

    # if at leaf level, write row name and data at proper indentation, push row counter +1
    # else, write metric name and recursively call `create_data_rows`, on items in `in_dict` pushing indentation counter +1
    if vals_are_lists(in_dict):
        for name, datavec in in_dict.items():
            if name == "--":
                row_offset -= 1
                worksheet.write_row(
                    row_offset, nums_offset, datavec, format_dict["num"]
                )

                formula_offset = row_offset + 1
                formulers = [
                    f'=IFERROR(OFFSET($F{formula_offset}, 0, MATCH(B$3, $G$3:$DB$3, 0)), "-")',
                    f'=IFERROR(OFFSET($F{formula_offset}, 0, MATCH(C$3, $G$3:$DB$3, 0)), "-")',
                    f'=IFERROR(C{formula_offset}-B{formula_offset}, "-")',
                ]
                pct_cell = f'=IFERROR(C{formula_offset}/B{formula_offset}-1, "-")'

                worksheet.write_row(row_offset, 1, formulers, format_dict["num"])
                worksheet.write(row_offset, 4, pct_cell, format_dict["pct"])

                row_offset += 1
            else:
                format_dict["group"].set_indent(ind_level + 1)
                worksheet.write(row_offset, 0, name, format_dict["group"])
                worksheet.write_row(
                    row_offset, nums_offset, datavec, format_dict["num"]
                )
                formula_offset = row_offset + 1
                formulers = [
                    f'=IFERROR(OFFSET($F{formula_offset}, 0, MATCH(B$3, $G$3:$DB$3, 0)), "-")',
                    f'=IFERROR(OFFSET($F{formula_offset}, 0, MATCH(C$3, $G$3:$DB$3, 0)), "-")',
                    f'=IFERROR(C{formula_offset}-B{formula_offset}, "-")',
                ]
                pct_cell = f'=IFERROR(C{formula_offset}/B{formula_offset}-1, "-")'
                worksheet.write_row(row_offset, 1, formulers, format_dict["num"])
                worksheet.write(row_offset, 4, pct_cell, format_dict["pct"])
                row_offset += 1

        # worksheet.write(row_offset, 0, None, groupformat)
        row_offset += 1
    else:
        for name, nested_dict in in_dict.items():
            format_dict["group"].set_indent(ind_level)
            if ind_level == 0:
                # write `bigname` to sheet
                bigname = "-- " + name + " --"
                format_dict["group"].set_bold(True)
                format_dict["group"].set_font_size(9)
                worksheet.write(row_offset, 0, bigname, format_dict["group"])
                worksheet.write(row_offset, nums_offset, None, format_dict["l"])

                # create index links
                to_cell = xl_rowcol_to_cell(row_offset, 0)
                link_string = f"internal:{sheetname!r}!{to_cell}"
                index_sheet.write_url(index_row_offset, 1, link_string)
                index_sheet.write(index_row_offset, 1, name, format_dict["index_group"])
                index_row_offset += 1

                # bump `row_offset` and `next_ind` and recurse again for each `nested_dict`
                row_offset += 1
                next_ind = ind_level + 1
                create_data_rows(
                    worksheet,
                    nested_dict,
                    workbook,
                    groupnames,
                    scenarionames,
                    next_ind,
                    sheetname,
                    format_dict,
                    False,
                )
            elif ind_level == 1:
                format_dict["group"].set_bold(True)
                worksheet.write(row_offset, 0, name, format_dict["group"])
                worksheet.write(row_offset, nums_offset, None, format_dict["l"])
                row_offset += 1
                next_ind = ind_level + 1
                create_data_rows(
                    worksheet,
                    nested_dict,
                    workbook,
                    groupnames,
                    scenarionames,
                    next_ind,
                    sheetname,
                    format_dict,
                    False,
                )
            else:
                if name == "--":
                    next_ind = ind_level
                    create_data_rows(
                        worksheet,
                        nested_dict,
                        workbook,
                        groupnames,
                        scenarionames,
                        next_ind,
                        sheetname,
                        format_dict,
                        False,
                    )
                else:
                    # worksheet.write(row_offset, 0, None, groupformat)
                    # row_offset += 1
                    worksheet.write(row_offset, 0, name, format_dict["group"])
                    row_offset += 1
                    next_ind = ind_level + 1
                    create_data_rows(
                        worksheet,
                        nested_dict,
                        workbook,
                        groupnames,
                        scenarionames,
                        next_ind,
                        sheetname,
                        format_dict,
                        False,
                    )


# Creates excel file and writes to disk at the end.
def create_xl_from_df(in_df: pd.DataFrame, excel_out_path) -> None:
    workbook = xlsxwriter.Workbook(excel_out_path)

    # create the index sheet
    workbook.add_worksheet("Index")

    # create a dict of formats used in the workbook
    format_dict = create_format_dict(workbook)

    global index_sheet
    index_sheet = workbook.get_worksheet_by_name("Index")
    global index_row_offset
    index_row_offset = 0

    in_df = shorten_long_sheetnames(in_df)
    in_df = in_df.fillna("")
    sheetnames = get_sheetnames(in_df)

    for sheetname in sheetnames:
        logging.info(f"creating {sheetname} sheet")
        global row_offset
        row_offset = 3

        workbook.add_worksheet(sheetname)
        worksheet = workbook.get_worksheet_by_name(sheetname)
        worksheet.set_default_row(18)
        worksheet.hide_gridlines(2)

        sheet_df = create_sheet_df(in_df, sheetname)
        sheet_dict = create_sheet_dict(in_df, sheetname)
        scenarionames = get_scenarios(sheet_df)
        groupnames = get_groups(sheet_df)

        create_data_rows(
            worksheet,
            sheet_dict,
            workbook,
            groupnames,
            scenarionames,
            0,
            sheetname,
            format_dict,
            writeIndexHeader=True,
        )
        create_header_block(
            sheetname,
            worksheet,
            sheet_dict,
            workbook,
            groupnames,
            scenarionames,
            format_dict,
        )
        create_dynamic_block(worksheet, workbook, format_dict)

        worksheet.set_default_row(hide_unused_rows=True)
        worksheet.autofit()

        # setting column widths
        worksheet.set_column("A:A", 33.33)
        worksheet.set_column("B:B", 10.67)
        worksheet.set_column("C:C", 10.67)
        worksheet.set_column("D:D", 10.67)
        worksheet.set_column("E:E", 5)
        worksheet.set_column("G:XFD", 10.67)
        logging.info(f"{sheetname} sheet done")

    logging.info(f"Writing excel file to disk as {excel_out_path}")
    index_sheet.autofit()
    index_sheet.set_column("B:B", 33.33)
    index_sheet.hide_gridlines(2)
    workbook.close()
