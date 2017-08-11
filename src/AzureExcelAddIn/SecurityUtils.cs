using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelAddIn1
{
    internal static class SecurityUtils
    {
        private const string DataPersistenceFileName = "UsageReportSavedParameters.dat";
        private const string DataPersistenceFolderName = ".azureexceladdin";

        private static string GetDataPersistenceFolder()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), DataPersistenceFolderName);
        }

        private static string GetDataPersistenceFile(string filename)
        {
            var folderPath = GetDataPersistenceFolder();
            DirectoryInfo directoryInfo = new DirectoryInfo(folderPath);
            if (!directoryInfo.Exists)
            {
                directoryInfo.Create();
            }

            return Path.Combine(directoryInfo.FullName, filename);
        }

        private static string ReadProtectedData(string file)
        {
            var bytes = File.ReadAllBytes(file);
            bytes = ProtectedData.Unprotect(bytes, null, DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(bytes);
        }

        private static void WriteProtectedData(string file, string content)
        {
            var bytes = Encoding.UTF8.GetBytes(content);
            bytes = ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser);
            File.WriteAllBytes(file, bytes);
        }

        public static void SaveUsageReportParameters(PersistedData persistedData)
        {
            var json = JObject.FromObject(persistedData);
            WriteProtectedData(GetDataPersistenceFile(DataPersistenceFileName), json.ToString());
        }

        public static PersistedData GetSavedUsageReportParameters()
        {
            var file = GetDataPersistenceFile(DataPersistenceFileName);
            if (!File.Exists(file))
            {
                return new PersistedData();
            }

            return JsonConvert.DeserializeObject<PersistedData>(ReadProtectedData(file));
        }
    }
}
