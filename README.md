# Command line utility to run Excel workbooks.

## Usage examples
```
xlrun.exe  -xlFileOpen MyWorkbook.xlsx  -xlRefreshLeftToRight  -xlRngGet Summary!TestStatus

xlrun.exe  -xlFileOpen MyMacrobook.xlsm  -xlEvalMacro MyMacro  -xlRngGet Summary!B4  -xlFileSave

xlrun.exe  -xlFileNew  -xlRngSet A1 1.0  -xlRngGet A1  -xlRngSet A2 =today()  -xlRngGet A2 -xlFileSaveAs MyGeneratedBook.xlsx
```


## Requirements & Build instructions

A local install of Excel is required. 
Excel is invoked using COM Automation.


Project file was created with (manually targeting net48 for COM support):
```
dotnet new console
dotnet add package Microsoft.Office.Interop.Excel
```

Project is built with:
```
dotnet build
```

