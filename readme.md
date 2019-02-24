# Excel-Extraction


[![CircleCI](https://circleci.com/gh/horitaku1124/Excel-Extraction.svg?style=svg)](https://circleci.com/gh/horitaku1124/Excel-Extraction)

## How to build
```bash
./gradlew build
```

## Features

* ExtractBySheet - create csv files by each sheets
* ExcelQuery - prints cell data with like SQL
* DumpData - prints all cell texts

## How to execute

### ExtractBySheet
```bash
./gradlew run -Pargs="-sheets 新宿,東京 -in ./data/sample1.xlsx -out ./out/data -divide 3"
or
java -cp build/libs/excel_extraction-1.0-SNAPSHOT.jar ExtractBySheet -sheets 新宿,東京 -in ./data/sample1.xlsx -out ./out/data -divide 3
```

### ExcelQuery
```bash
java -cp build/libs/excel_extraction-1.0-SNAPSHOT.jar ExcelQuery "select 日付 from `./data/sample1.xlsx`.東京"
```

### DumpData
```bash
java -cp build/libs/excel_extraction-1.0-SNAPSHOT.jar DumpData ./data/sample1.xlsx
```