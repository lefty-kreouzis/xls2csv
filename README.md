# xls2csv

A command-line utility that converts Excel `.xlsx` workbooks to CSV using Spring Shell and Apache POI.

## Requirements

- Java 17 or newer
- Maven 3.9 or newer, or the included Maven wrapper

## Build

On Windows:

```powershell
./mvnw.cmd clean package
```

On Linux or macOS:

```bash
./mvnw clean package
```

The packaged application is created under `target/`.

## Usage

Start the interactive shell with:

```powershell
java -jar target/xls2csv-0.0.1-SNAPSHOT.jar
```

At the shell prompt, convert a workbook to standard output:

```text
xls2csv --fromFile input.xlsx
```

Write the CSV to a file instead:

```text
xls2csv --fromFile input.xlsx --toFile output.csv
```

`--toFile` defaults to `stdout`. The command processes every worksheet in workbook order and writes the rows to one CSV stream.

## Conversion details

- String cells are written as text.
- Numeric cells are written without unnecessary trailing zeroes.
- Date-formatted cells use `yyyy-MM-dd HH:mm:ss`.
- Formula cells use their cached result when it is a string or number.
- Empty or missing cells are written as empty CSV fields.

## Tests

Run the test suite with:

```powershell
./mvnw.cmd test
```

On Linux or macOS, use `./mvnw test`.
