using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelAddIn1
{
    internal static class JsonUtils
    {
        public static string ExtractTagsFromInstanceData(string instanceData)
        {
            if (String.IsNullOrWhiteSpace(instanceData))
            {
                return String.Empty;
            }

            JObject instanceDataObject = JsonConvert.DeserializeObject<JObject>(instanceData);
            JToken resourcesToken = instanceDataObject.First.Value<JToken>();
            if (resourcesToken.Children()["tags"].Any())
            {
                JToken tagsToken = resourcesToken.Children()["tags"].First();
                return tagsToken.HasValues ? tagsToken.ToString() : String.Empty;
            }

            return String.Empty;
        }

        public static string ExtractInfoFields(JObject infoFields)
        {
            if (infoFields != null)
            {
                return infoFields.ToString();
            }

            return String.Empty;
        }
    }
}
