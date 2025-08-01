using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using UnityEngine;
using UnityEngine.UI;
using TMPro;
using Utilities.FileControl;

namespace DemoExcelFileHandler
{
    public class FileMenu : MonoBehaviour
    {
        [Header("UI")]
        [SerializeField] private Button ButtonLoad;
        [SerializeField] private Button ButtonUpload;
        [SerializeField] private TMP_Text OutputText;

        [Header("FileControl")]
        [SerializeField] private FilePicker FilePicker;
        [SerializeField] private FileHandler FileHandler;
        [SerializeField] private ExcelFileHandler ExcelFileHandler;

        private const string electricityName = "Electricity";
        private string[,] electricityDb = new string[,] {
            { "Date", "Day", "Night" },
            { "1.12.2023" , "123", "12345" },
            { "11.3.2024" , "425", "4562" },
            { "6.6.2025" , "556", "6467" }
        };
        private const string waterName = "Water";
        private string[,] waterDb = new string[,] {
            { "Date", "Day", "Night" },
            { "2.10.2023" , "353", "35467" },
            { "15.4.2024" , "456", "5677" },
            { "4.7.2025" , "576", "7456" }
        };
        private const string heatingName = "Heating";
        private string[,] heatingDb = new string[,] {
            { "Date", "Heating" },
            { "2.10.2023" , "353" },
            { "15.4.2024" , "456" },
            { "4.7.2025" , "576" }
        };

        private const int columnWidth = 20;
        private const int rowHeight = 25;

        void Start()
        {
            ButtonLoad.onClick.AddListener(OnClickButtonLoad);
            ButtonUpload.onClick.AddListener(OnClickButtonUpload);
        }

        private void OnClickButtonUpload()
        {
            var path = Path.Combine(Application.temporaryCachePath, "Test.xlsx");
            CreateFile(path);
            FilePicker.ExportFile(path);
        }

        private void OnClickButtonLoad()
        {
            FilePicker.PickFile("xlsx", PickFileCallback);
        }

        private void PickFileCallback(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) {
                Debug.Log("Operation cancelled.");
            } else {
                Debug.Log("Picked file: " + path);
                var data = ReadDataFromFile(path);
                OutputText.text = ArrayToString(data.Electricity) + ArrayToString(data.Water) + ArrayToString(data.Heating);
            }
        }

        private void CreateFile(string path)
        {
            FileHandler.DeleteFileIfExist(path);

            ExcelFileHandler.CreateFile(path, FileCreatedCallback);
        }

        private void FileCreatedCallback(ExcelPackage package)
        {
            FillElectricity(package, electricityDb);
            FillWater(package, waterDb);
            FillHeating(package, heatingDb);
        }

        private void FillElectricity(ExcelPackage package, string[,] electricity)
        {
            var worksheet = package.Workbook.Worksheets.Add(electricityName);
            ExcelFileHandler.WriteCells(worksheet, 1, 1, electricity);
            ExcelFileHandler.SetRowsHeight(worksheet, rowHeight, 4);
            ExcelFileHandler.SetColumnsWidth(worksheet, columnWidth, 3);
        }

        private void FillWater(ExcelPackage package, string[,] water)
        {
            var worksheet = package.Workbook.Worksheets.Add(waterName);
            ExcelFileHandler.WriteCells(worksheet, 1, 1, water);
            ExcelFileHandler.SetRowsHeight(worksheet, rowHeight, 4);
            ExcelFileHandler.SetColumnsWidth(worksheet, columnWidth, 3);
        }

        private void FillHeating(ExcelPackage package, string[,] heating)
        {
            var worksheet = package.Workbook.Worksheets.Add(heatingName);
            ExcelFileHandler.WriteCells(worksheet, 1, 1, heating);
            ExcelFileHandler.SetRowsHeight(worksheet, rowHeight, 4);
            ExcelFileHandler.SetColumnsWidth(worksheet, columnWidth, 2);
        }

        private (string[,] Electricity, string[,] Water, string[,] Heating) ReadDataFromFile(string path)
        {
            if (!File.Exists(path)) {
                Debug.LogError("File not exist!");
                return (null, null, null);
            }
            var fileInfo = new FileInfo(path);

            string[,] electricity;
            string[,] water;
            string[,] heating;
            using (var package = new ExcelPackage(fileInfo)) {
                var worksheet = package.Workbook.Worksheets[electricityName];
                electricity = ExcelFileHandler.ReadAllWorksheet(worksheet);
                //electricity = ExcelFileHandler.ReadWorksheet(worksheet, 2, 4, 1, 3); // This is for targeted data loading. This example loads data without headers.

                worksheet = package.Workbook.Worksheets[waterName];
                water = ExcelFileHandler.ReadAllWorksheet(worksheet);
                //water = ExcelFileHandler.ReadWorksheet(worksheet, 2, 4, 1, 3);

                worksheet = package.Workbook.Worksheets[heatingName];
                heating = ExcelFileHandler.ReadAllWorksheet(worksheet);
                //heating = ExcelFileHandler.ReadWorksheet(worksheet, 2, 4, 1, 2);
            }
            return (electricity, water, heating);
        }

        private string ArrayToString(string[,] array)
        {
            var rows = array.GetUpperBound(0) + 1;
            var columns = array.Length / rows;

            var result = new string("");
            for (int i = 0; i < rows; i++) {
                for (int j = 0; j < columns; j++) {
                    result = result + array[i, j] + "   ";
                }
                result = result + "\n";
            }
            return result;
        }

        void OnDestroy()
        {
            ButtonLoad.onClick.RemoveListener(OnClickButtonLoad);
            ButtonUpload.onClick.RemoveListener(OnClickButtonUpload);
        }
    }
}