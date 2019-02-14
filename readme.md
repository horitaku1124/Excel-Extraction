# Excel-Extraction

## How to build
```bash
./gradlew build
```

## How to execute
```bash
./gradlew run -Pargs="-sheets 新宿,東京 -in ./data/sample1.xlsx -out ./out/data -divide 3"
or
java -cp build/libs/excel_extraction-1.0-SNAPSHOT.jar ExtractBySheet -sheets 新宿,東京 -in ./data/sample1.xlsx -out ./out/data -divide 3
```