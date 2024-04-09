using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BarcodeGenerator.Model
{
    public class ProductNumberModel
    {
        public string ProductNumber { get; set; }
        public string ProductName { get; set; }
        public string ItemNo { get; set; }
        public int ProductSpecificationCode { get; set; }
        public int CapacityOfReagentCode { get; set; }
    }
}
