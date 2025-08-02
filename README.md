# Unity_ExcelFileHandler

This tool allows you to write and read excel files.

This tool uses EPPlus v8.0.8.

## Demo
You can run apk on Android or run project on Unity version 2022.3.61f1.

To work with files in demo, following tools were used:
* UnityNativeFilePicker from [GitHub](https://github.com/yasirkula/UnityNativeFilePicker) or [AssetStore](https://assetstore.unity.com/packages/tools/integration/native-file-picker-for-android-ios-173238).
* Unity_FilePicker from [GitHub](https://github.com/LordOfTheFridge/Unity_FilePicker)

## Installation

1. Import [NuGetForUnity](https://github.com/GlitchEnzo/NuGetForUnity).
2. Install EPPlus using NuGetForUnity.
3. Import unity package from this repository.


## Usage
### CreateFile
This method creates excel file. In callback, you write data to file or perform necessary manipulations. After that, method will close this file.

```csharp
ExcelFileHandler.CreateFile(path, FileCreatedCallback);

private void FileCreatedCallback(ExcelPackage package)
{
    FillElectricity(package, electricityDb);
    FillWater(package, waterDb);
    FillHeating(package, heatingDb);
}
```

### WriteCells
This method writes passed data to cells. You simply specify starting positions (starting from 1, not 0!) and data set.

```csharp
ExcelFileHandler.WriteCells(worksheet, 1, 1, data);
```

### ReadAllWorksheet
This method reads all worksheet.

```csharp
var data = ExcelFileHandler.ReadAllWorksheet(worksheet);
```

### ReadWorksheet
This method reads only specified worksheet positions (starting from 1, not 0!).

```csharp
var data = ExcelFileHandler.ReadWorksheet(worksheet, 2, 4, 1, 3);
```

### SetColumnsWidth
This method changes width of specified columns.

```csharp
ExcelFileHandler.SetColumnsWidth(worksheet, 20, 3);
```

### SetColumnsWidth
This method changes height of specified rows.

```csharp
ExcelFileHandler.SetRowsHeight(worksheet, 25, 3);
```
