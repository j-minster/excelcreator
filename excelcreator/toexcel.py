import pandas as pd
import os
import pathlib
from itertools import filterfalse
import click
import xlsxwriter
import functools
import itertools
import operator
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell
import sys
import re
import multiprocessing as mp


# helper functions
def compose(*functions):
    """
    Compose a series of functions together to be called as a single function.
    """

    def compose2(f, g):
        return lambda x: f(g(x))

    return functools.reduce(compose2, functions, lambda x: x)


def is_text(df: pd.DataFrame) -> list[bool]:
    """
    Return a list with a boolean for each column of `df`, depending on whether or not
    the column contains str type data.
    """
    keys = list(df.keys())
    is_text = map(lambda key: type(df[key][0]) == str, keys)
    return list(is_text)


def get_groups(df: pd.DataFrame) -> list[str]:
    """
    Get all column names in `df` which don't contain a year (ie. 2049).
    Should return all columns besides those with names of scenarios.
    """
    allcols = list(df.columns)
    groupnames = [
        name for name in allcols if not re.search("(19|[2-9][0-9])\d{2}", name)
    ]
    return groupnames


def drop_rows_containing(df: pd.DataFrame, string: str) -> pd.DataFrame:
    """
    Drop rows in `df` which contain entries matching `string`.
    """
    rowlists = [row.to_list() for _, row in df.iterrows()]
    indices = [i for i, row in enumerate(rowlists) if string in row]
    return df.drop(indices)


def df_from_clargs(input_csv_path: str, output_excel_path: str) -> pd.DataFrame:
    """
    Read command-line args and:
    + set the output excel file path
    + read the input csv as a pandas DataFrame
    """
    global excel_out_path
    excel_out_path = output_excel_path
    input_df = pd.read_csv(input_csv_path)

    return input_df


def get_scenarios(df: pd.DataFrame) -> list[str]:
    # will ignore empty scenarios only containing ' '
    """
    Get non-empty data for scenarios - for use in writing excel rows.
    """
    groups = get_groups(df)
    scen_cols = list(df.columns)
    for group in groups:
        scen_cols.remove(group)
    scen_data = [list(df.get(col)) for col in scen_cols]
    scen_sets = [set(d) for d in scen_data]
    nonempty_scens = [col for col, s in zip(scen_cols, scen_sets) if len(s) > 1]
    return nonempty_scens


# infinite dict class
class NestedDict(dict):
    def __getitem__(self, key):
        if key in self:
            return self.get(key)
        else:
            value = NestedDict()
            self[key] = value
            return value


def df_to_dict(in_df: pd.DataFrame) -> NestedDict:
    """
    Convert `in_df` to NestedDict structure.
    """
    # create emtpy infinite dict to fill
    big_dict = NestedDict()
    groupnames = get_groups(in_df)
    scenarionames = get_scenarios(in_df)
    # fill the dict with values from `input_df`
    for i, row in in_df.iterrows():
        dim_name = row.Mode
        scenario_data = row[scenarionames].to_list()
        upd_d = {dim_name: scenario_data}

        # create call to dict: ie. dict[group1][group2]...[groupn]
        gbrackets = [f"[row.{group}]" for group in groupnames]
        dict_accessor = "big_dict" + functools.reduce(operator.concat, gbrackets[:-1])
        ex_d = eval(dict_accessor)  # old dictionary, if any
        new_d = {**ex_d, **upd_d}  # update and create new dictionary
        dict_set = dict_accessor + " = new_d"  # expression to assign old dict to new
        exec(dict_set)  # execute dict_set

    return big_dict


def shorten_long_sheetnames(in_df: pd.DataFrame) -> pd.DataFrame:
    """
    Shorten the names of sheets to meet excel's 31 character limit.
    """
    rlist = [
        ("Average", "Avg."),
        ("Distance", "Dist."),
        ("Distances", "Dists"),
        ("Terminating", "Term."),
        ("Originating", "Orig."),
        ("Population", "Pop."),
    ]

    def replace_multi(rlist, name):
        for frm, to in rlist:
            if len(name) > 31:
                name = name.replace(frm, to)
        return name

    sheetname_col = list(in_df.columns)[0]
    sheetnames = list(in_df.loc[:, sheetname_col])
    replaced_sheetnames = [replace_multi(rlist, name) for name in sheetnames]

    if in_df[sheetname_col].to_list() != replaced_sheetnames:
        print("Sheet names shortened to be < 31 chars")

    in_df[sheetname_col] = replaced_sheetnames
    return in_df


def get_sheetnames(in_df: pd.DataFrame) -> set[str]:
    """
    Get the names of the sheets that will exist in the final excel file.
    """
    groups = get_groups(in_df)
    sheetname_col = groups[0]
    sheetnames = in_df.loc[:, sheetname_col]
    sheetset = set(sheetnames)
    return sheetset


def create_sheet_df(in_df: pd.DataFrame, sheetname: str) -> pd.DataFrame:
    """
    Extract entries from `in_df` for the relevant `sheetname`
    """
    majorcol = list(in_df.columns)[0]
    sub_df = in_df[in_df[majorcol].isin([sheetname])]
    return sub_df


### create individual dictionaries for sheets rather than one huge dictionary for the whole dataframe
def create_sheet_dict(in_df: pd.DataFrame, sheetname: str) -> NestedDict:
    sheetnames = get_sheetnames(in_df)
    sub_df = create_sheet_df(in_df, sheetname)
    d1 = df_to_dict(sub_df)

    idx = list(sheetnames).index(sheetname)
    print(f"made '{sheetname}' dict")
    print(f"(Creating {idx+1} of {len(sheetnames)} total sheets)")

    return d1


### function to create the top block of a sheet. Returns nothing. Just alters `workbook` in memory.
def create_header_block(
    sheetname: str,
    worksheet: xlsxwriter.worksheet.Worksheet,
    sheet_dict: NestedDict,
    workbook: xlsxwriter.workbook.Workbook,
    groupnames: list[str],
    scenarionames: list[str],
) -> None:
    # constants throughout
    bordercolor = "#9B9B9B"
    orangecolor = "#F7D8AA"
    greycolor = "#F0F0F0"

    # merge cells A1 and A2
    # write the sheet name in the merged cells, large font, orange cell
    sheetname_format = workbook.add_format(
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

    worksheet.merge_range("A1:A2", sheetname, sheetname_format)

    # set height of sheet name cell to 36
    worksheet.set_row_pixels(0, 24)
    worksheet.set_row_pixels(1, 24)

    # set sheetname and metrics column (A) to width 43.33
    worksheet.set_column("A:A", 34.83)

    # write 'Metric' in cell A3, bold format, make the cell vertically taller (36 px), grey background
    metric_format = workbook.add_format(
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
    worksheet.write("A3", "Metric", metric_format)
    worksheet.set_row_pixels(2, 48)

    # write all scenario* names in {H...}:3
    scenario_format = workbook.add_format(
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
    l_scenario_format = workbook.add_format(
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
    r_scenario_format = workbook.add_format(
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

    scenario_col_offset = 6
    for offset, name in enumerate(scenarionames):
        if offset == 0:
            worksheet.write(2, scenario_col_offset + offset, name, l_scenario_format)

        elif offset == len(scenarionames) - 1:
            worksheet.write(2, scenario_col_offset + offset, name, r_scenario_format)

        else:
            worksheet.write(2, scenario_col_offset + offset, name, scenario_format)

    # create dropdowns
    dropdown_format = workbook.add_format(
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
    worksheet.write(input_cell_1, scenarionames[0], dropdown_format)

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
    worksheet.write(input_cell_2, scenarionames[1], dropdown_format)

    # create +/- headings
    pmformat = workbook.add_format(
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
    pmcell = "D$3"
    worksheet.write(pmcell, "+/-", pmformat)
    pcell = "E$3"
    worksheet.write(pcell, "%", pmformat)

    # merge cells {B, C, D, E}:2
    # write 'compare loaded scenarios...' in the merged cells, blue cell

    comp_format = workbook.add_format(
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
    lilcell_format = workbook.add_format(
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

    worksheet.merge_range(
        "B2:E2", "Compare two loaded scenarios (use dropdowns)", comp_format
    )
    worksheet.write("F2", None, lilcell_format)
    worksheet.write("F3", None, pmformat)


def create_dynamic_block(
    worksheet: xlsxwriter.worksheet.Worksheet, workbook: xlsxwriter.workbook.Workbook
) -> None:
    # make column F small and G zero-width
    rformat = workbook.add_format({"right": 1, "border_color": "#9B9B9B"})
    worksheet.set_column("F:F", 2.33, rformat)
    # worksheet.set_column('G:G', 0)


# https://stackoverflow.com/questions/23499017/know-the-depth-of-a-dictionary
def dict_depth(d) -> int:
    if isinstance(d, dict):
        return 1 + (max(map(dict_depth, d.values())) if d else 0)
    return 0


### check whether the values (not the keys) in the dictionary `d` are lists
def vals_are_lists(d: NestedDict) -> bool:
    boollist = [isinstance(val, list) for _, val in d.items()]
    return all(boollist)


### input the data for each sheet (warning: recursion)
def create_data_rows(
    worksheet: xlsxwriter.worksheet.Worksheet,
    in_dict: NestedDict,
    workbook: xlsxwriter.workbook.Workbook,
    groupnames: list[str],
    scenarionames: list[str],
    ind_level: int,
    sheetname: str,
    writeIndexHeader: bool,
) -> None:
    nums_offset = 6
    global row_offset
    global index_row_offset

    groupformat = workbook.add_format(
        {
            "bold": False,
            "font_name": "Segoe UI (Body)",
            "font_size": 8,
            "right": 1,
            "border_color": "#9B9B9B",
        }
    )
    index_groupformat = workbook.add_format(
        {
            "bold": True,
            "font_name": "Arial Narrow",
            "font_size": 11,
            "font_color": "blue",
            "underline": 1,
            "indent": 1,
        }
    )
    index_headerformat = workbook.add_format(
        {
            "bold": True,
            "bottom": 1,
            "font_name": "Arial Narrow",
            "font_size": 16,
            "indent": 0,
        }
    )
    index_ulformat = workbook.add_format({"bottom": 1})
    numformat = workbook.add_format(
        {"font_name": "Segoe UI (Body)", "font_size": 8, "num_format": "#,##0.000"}
    )
    pctformat = workbook.add_format(
        {"font_name": "Segoe UI (Body)", "font_size": 8, "num_format": "0.0%"}
    )
    lformat = workbook.add_format({"left": 1, "border_color": "#9B9B9B"})
    if writeIndexHeader:
        index_row_offset += 1
        link_string = f"internal:{sheetname!r}!A1"
        index_sheet.write_url(index_row_offset, 1, link_string)
        index_sheet.write(index_row_offset, 1, sheetname, index_headerformat)
        index_row_offset += 1

    # if at leaf level, write row name and data at proper indentation, push row counter +1
    # else, write metric name and recursively call `create_data_rows`, on items in `in_dict` pushing indentation counter +1
    if vals_are_lists(in_dict):
        for name, datavec in in_dict.items():
            if name == "--":
                row_offset -= 1
                worksheet.write_row(row_offset, nums_offset, datavec, numformat)

                formula_offset = row_offset + 1
                formulers = [
                    f'=IFERROR(OFFSET($F{formula_offset}, 0, MATCH(B$3, $G$3:$DB$3, 0)), "-")',
                    f'=IFERROR(OFFSET($F{formula_offset}, 0, MATCH(C$3, $G$3:$DB$3, 0)), "-")',
                    f'=IFERROR(C{formula_offset}-B{formula_offset}, "-")',
                ]
                pct_cell = f'=IFERROR(C{formula_offset}/B{formula_offset}-1, "-")'

                worksheet.write_row(row_offset, 1, formulers, numformat)
                worksheet.write(row_offset, 4, pct_cell, pctformat)

                row_offset += 1
            else:
                groupformat.set_indent(ind_level + 1)
                worksheet.write(row_offset, 0, name, groupformat)
                worksheet.write_row(row_offset, nums_offset, datavec, numformat)
                formula_offset = row_offset + 1
                formulers = [
                    f'=IFERROR(OFFSET($F{formula_offset}, 0, MATCH(B$3, $G$3:$DB$3, 0)), "-")',
                    f'=IFERROR(OFFSET($F{formula_offset}, 0, MATCH(C$3, $G$3:$DB$3, 0)), "-")',
                    f'=IFERROR(C{formula_offset}-B{formula_offset}, "-")',
                ]
                pct_cell = f'=IFERROR(C{formula_offset}/B{formula_offset}-1, "-")'
                worksheet.write_row(row_offset, 1, formulers, numformat)
                worksheet.write(row_offset, 4, pct_cell, pctformat)
                row_offset += 1

        # worksheet.write(row_offset, 0, None, groupformat)
        row_offset += 1
    else:
        for name, nested_dict in in_dict.items():
            groupformat.set_indent(ind_level)
            if ind_level == 0:
                # write `bigname` to sheet
                bigname = "-- " + name + " --"
                groupformat.set_bold(True)
                groupformat.set_font_size(9)
                worksheet.write(row_offset, 0, bigname, groupformat)
                worksheet.write(row_offset, nums_offset, None, lformat)

                # create index links
                to_cell = xl_rowcol_to_cell(row_offset, 0)
                link_string = f"internal:{sheetname!r}!{to_cell}"
                index_sheet.write_url(index_row_offset, 1, link_string)
                index_sheet.write(index_row_offset, 1, name, index_groupformat)
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
                    False,
                )
            elif ind_level == 1:
                groupformat.set_bold(True)
                worksheet.write(row_offset, 0, name, groupformat)
                worksheet.write(row_offset, nums_offset, None, lformat)
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
                        False,
                    )
                else:
                    # worksheet.write(row_offset, 0, None, groupformat)
                    # row_offset += 1
                    worksheet.write(row_offset, 0, name, groupformat)
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
                        False,
                    )


# Creates excel file and writes to disk at the end.
def create_xl_from_df(in_df: pd.DataFrame) -> None:
    global excel_out_path
    workbook = xlsxwriter.Workbook(excel_out_path)

    # create the index sheet
    workbook.add_worksheet("Index")
    global index_sheet
    index_sheet = workbook.get_worksheet_by_name("Index")
    global index_row_offset
    index_row_offset = 0

    in_df = shorten_long_sheetnames(in_df)
    in_df = in_df.fillna("")
    sheetnames = get_sheetnames(in_df)

    for sheetname in sheetnames:
        print(f"creating {sheetname} sheet")
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
            writeIndexHeader=True,
        )
        create_header_block(
            sheetname, worksheet, sheet_dict, workbook, groupnames, scenarionames
        )
        create_dynamic_block(worksheet, workbook)

        worksheet.set_default_row(hide_unused_rows=True)
        worksheet.autofit()

        # setting column widths
        worksheet.set_column("A:A", 33.33)
        worksheet.set_column("B:B", 10.67)
        worksheet.set_column("C:C", 10.67)
        worksheet.set_column("D:D", 10.67)
        worksheet.set_column("E:E", 5)
        worksheet.set_column("G:XFD", 10.67)
        print(f"{sheetname} sheet done")

    print(f"Writing excel file to disk as {excel_out_path}")
    index_sheet.autofit()
    index_sheet.set_column("B:B", 33.33)
    index_sheet.hide_gridlines(2)
    workbook.close()


@click.command()
@click.argument("input_csv_path", required=True)
@click.option(
    "--output_folder",
    "-o",
    required=False,
    default=os.path.join("outputs"),
)
@click.option("--output_filename", "-n", required=False, default=r"output.xlsx")
def run(input_csv_path: str, output_folder: str, output_filename: str) -> None:
    """
    INPUT_CSV_PATH: relative path to the CSV file to be converted
    """

    if "/" in output_folder or "\\" in output_folder:
        click.echo(f"{output_folder} contains slashes")
        click.echo("Please supply a folder name only, not a path")
        sys.exit()

    if "/" in output_filename or "\\" in output_filename:
        click.echo(f"{output_filename} contains slashes")
        click.echo("Please supply a file name only, not a path")
        sys.exit()

    if not os.path.exists(output_folder):
        click.echo(f"Creating output folder: {output_folder}")
        os.mkdir(output_folder)

    if pathlib.Path(output_filename).suffix != ".xlsx":
        click.echo("Funny business detected, changing extension of output to xlsx...")
        output_filename = os.path.splitext(output_filename)[0] + ".xlsx"

    output_excel_path = os.path.join(output_folder, output_filename)

    input_df = df_from_clargs(input_csv_path, output_excel_path)
    create_xl_from_df(input_df)
