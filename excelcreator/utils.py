# https://stackoverflow.com/questions/23499017/know-the-depth-of-a-dictionary
import functools
import logging
import operator
import re

import pandas as pd


# infinite dict class
class NestedDict(dict):
    def __getitem__(self, key):
        if key in self:
            return self.get(key)
        else:
            value = NestedDict()
            self[key] = value
            return value


def dict_depth(d) -> int:
    """
    Calculates the depth of the dictionary, `d`
    """
    if isinstance(d, dict):
        return 1 + (max(map(dict_depth, d.values())) if d else 0)
    return 0


def vals_are_lists(d: NestedDict) -> bool:
    """
    check whether the values (not the keys) in the dictionary `d` are lists
    """
    boollist = [isinstance(val, list) for _, val in d.items()]
    return all(boollist)


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


def df_from_clargs(input_csv_path: str) -> pd.DataFrame:
    """
    Read command-line args and:
    + set the output excel file path
    + read the input csv as a pandas DataFrame
    """
    input_df = pd.read_csv(input_csv_path)

    return input_df


def get_scenarios(df: pd.DataFrame) -> list[str]:
    """
    Get non-empty data for scenarios - for use in writing excel rows.
    Will ignore empty scenarios only containing ' '
    """
    groups = get_groups(df)
    scen_cols = list(df.columns)
    for group in groups:
        scen_cols.remove(group)
    scen_data = [list(df.get(col)) for col in scen_cols]
    scen_sets = [set(d) for d in scen_data]
    nonempty_scens = [col for col, s in zip(scen_cols, scen_sets) if len(s) > 1]
    return nonempty_scens


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
        gbrackets = [f"[row[{group!r}]]" for group in groupnames]
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
        if len(name) > 31:
            name = name.replace(" ", "")
        if len(name) > 31:
            name = name[0:31]
        return name

    sheetname_col = list(in_df.columns)[0]
    sheetnames = list(in_df.loc[:, sheetname_col])
    replaced_sheetnames = [replace_multi(rlist, name) for name in sheetnames]

    if in_df[sheetname_col].to_list() != replaced_sheetnames:
        logging.info("Sheet names shortened to be < 31 chars")

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
