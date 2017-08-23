using Newtonsoft.Json;

namespace ExcelAddIn1
{
    public class InstanceData
    {
        [JsonProperty("Microsoft.Resources")]
        public MicrosoftResources MicrosoftResources { get; set; }
    }
}
