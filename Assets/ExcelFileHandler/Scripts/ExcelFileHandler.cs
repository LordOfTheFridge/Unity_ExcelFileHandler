using OfficeOpenXml;
using System.IO;
using UnityEngine;

namespace Utilities.FileControl
{
    public class ExcelFileHandler : MonoBehaviour
    {
        public delegate void FileCreatedCallback(ExcelPackage package);

        void Start()
        {
            ExcelPackage.License.SetNonCommercialPersonal("test");
        }

        public void CreateFile(string path, FileCreatedCallback callback)
        {
            var fileInfo = new FileInfo(path);

            using (var package = new ExcelPackage(fileInfo)) {
                callback?.Invoke(package);

                package.Save();
            }
        }

        public void SetColumnsWidth(ExcelWorksheet worksheet, int width, int endPosition, int startPosition = 1)
        {
            for (var col = startPosition; col <= endPosition; col++) {
                worksheet.Column(col).Width = width;
            }
        }

        public void SetRowsHeight(ExcelWorksheet worksheet, int height, int endPosition, int startPosition = 1)
        {
            for (var row = startPosition; row <= endPosition; row++) {
                worksheet.Row(row).Height = height;
            }
        }

        public ExcelPackage GetFilePackage(string path)
        {
            if (!File.Exists(path)) {
                Debug.LogError("File does not exist!");
                return null;
            }
            var fileInfo = new FileInfo(path);
            return new ExcelPackage(fileInfo);
        }

        public void WriteCells(ExcelWorksheet worksheet, int rowStart, int columnStart, string[,] data)
        {
            if(worksheet == null) {
                Debug.LogError("Worksheet does not exist!");
                return;
            }

            if(
                data == null || 
                data.Length < 1 ||
                rowStart < 1 ||
                columnStart < 1
            ) {
                Debug.LogError("Entered data is incorrect!");
                return;
            }

            var rowCount = data.GetUpperBound(0) + 1;
            var columnCount = data.Length / rowCount;
            var i = rowStart;
            var j = columnStart;
            for (var row = 0; row < rowCount; row++) {
                for(var column = 0; column < columnCount; column++) {
                    worksheet.Cells[i, j].Value = data[row, column];
                    j++;
                }
                j = columnStart;
                i++;
            }
        }

        public string[,] ReadAllWorksheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null) {
                Debug.LogError("Worksheet does not exist!");
                return null;
            }

            var totalColumn = worksheet.Dimension.End.Column;
            var totalRow = worksheet.Dimension.End.Row;
            var result = new string[totalRow, totalColumn];
            for (var row = 1; row <= totalRow; row++) {
                for(var column = 1; column <= totalColumn; column++) {
                    result[row - 1, column - 1] = worksheet.Cells[row, column].Text;
                }
            }
            return result;
        }

        public string[,] ReadWorksheet(ExcelWorksheet worksheet, int rowStart, int rowEnd, int columnStart, int columnEnd)
        {
            if (worksheet == null) {
                Debug.LogError("Worksheet does not exist!");
                return null;
            }
            if (
                rowStart < 1 || columnStart < 1 ||
                rowEnd - rowStart < 1 || columnEnd - columnStart < 1
            ) {
                Debug.LogError("Entered data is incorrect!");
                return null;
            }
            if (rowEnd > worksheet.Dimension.End.Row || columnEnd > worksheet.Dimension.End.Column) {
                Debug.LogError("Entered data is incorrect!");
                return null;
            }

            var rowCount = rowEnd - rowStart + 1;
            var columnCount = columnEnd - columnStart + 1;
            var result = new string[rowCount, columnCount];
            var i = 0;
            var j = 0;
            for (var row = rowStart; row <= rowEnd; row++) {
                for (var column = columnStart; column <= columnEnd; column++) {
                    result[i, j] = worksheet.Cells[row, column].Text;
                    j++;
                }
                i++;
                j = 0;
            }
            return result;
        }
    }
}
