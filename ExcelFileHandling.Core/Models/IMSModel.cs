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
        public string ProductList { get; set; }
        public string ReceivedProducts { get; set; }
        public string Sales { get; set; }
        public string StocksInventory { get; set; }
        public string Pages { get; set; }
        public int Views { get; set; }
    }
}
