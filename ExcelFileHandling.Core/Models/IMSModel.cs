using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace ExcelFileHandling.Core.Models
{
    public class IMSModel
    {
        [JsonPropertyName("Product List")]
        public string product { get; set; }
        public string receivedProduct { get; set; }
        public string Sales { get; set; }
        public string Stock { get; set; }
        public string Page { get; set; }
        public int Views { get; set; }
    }
}
