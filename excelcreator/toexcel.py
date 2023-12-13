import logging
import os
import pathlib
import sys

import click

from .creators import create_xl_from_df
from .utils import df_from_clargs

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s",
)


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

    input_df = df_from_clargs(input_csv_path)
    create_xl_from_df(input_df, output_excel_path)
