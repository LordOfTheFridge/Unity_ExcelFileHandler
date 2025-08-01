using System.IO;
using UnityEngine;

namespace Utilities.FileControl
{
    public class FilePicker : MonoBehaviour
    {
        public void PickFile(string fileType, NativeFilePicker.FilePickedCallback callback)
        {
            var convertedFileType = NativeFilePicker.ConvertExtensionToFileType(fileType);
            NativeFilePicker.PickFile(callback, new string[] { convertedFileType });
        }

        public void ExportFile(string fileFullPath)
        {
            NativeFilePicker.ExportFile(fileFullPath, ExportFileCallback);
        }

        public void ExportFile(string folderPath, string fileName, string fileType)
        {
            var fileFullPath = Path.Combine(folderPath, fileName + "." + fileType);
            ExportFile(fileFullPath);
        }

        private void ExportFileCallback(bool success)
        {
            Debug.Log("SaveFile success: " + success.ToString());
        }
    }
}
