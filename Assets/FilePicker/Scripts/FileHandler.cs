using System.IO;
using UnityEngine;

namespace Utilities.FileControl
{
    public class FileHandler : MonoBehaviour
    {
        public static void DeleteFileIfExist(string path)
        {
            if (!File.Exists(path)) {
                return;
            }

            File.Delete(path);
        }

        public static void WriteTextToFile(string path, string data)
        {
            File.WriteAllText(path, data);
        }

        public static void WriteTextToFile(string path, string[] data)
        {
            File.WriteAllLines(path, data);
        }

        public static string[] ReadDataFromFile(string path)
        {
            if (!File.Exists(path)) {
                Debug.LogError("File not exist!");
                return null;
            }

            var lines = File.ReadAllLines(path);
            return lines;
        }
    }
}
