# excelcreator
## Installation

``` sh
git clone <THIS_PROJECT_URL>
cd excelcreator
poetry install
poetry shell
```

## Required CSV structure
The input CSV file should be structured like this:
| Group  | Boundaries | ( n grouping columns ) | Scenario 1 YEAR | ( ... ) | Scenario n YEAR |
|--------|------------|------------------------|-----------------|---------|-----------------|
| String | String     | Strings                | Number          | Numbers | Number          |

+ The output excel file will have as many sheets as there are unique 'Groups' in the 'Group' column
+ The remaining non-scenario columns will be used to group the data into blocks
+ Scenario names must contain a **YEAR** to be recognised.

## Usage

``` sh
poetry run createxl 'inputs/data.csv' \
    -o 'outputs' \
    -n 'output.xlsx'
```


