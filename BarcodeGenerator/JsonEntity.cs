using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BarcodeGenerator
{
    public class JsonEntity
    {
        public string Model { get; set; }
        public string Project { get; set; }
        public string Lot { get; set; }
        public string ExpiredDate { get; set; }
        public string ReagentSnr { get; set; }
        public string StdCalibrateLot { get; set; }
        public string StdItemAdensity { get; set; }
        public string StdItemALight { get; set; }
        public string StdItemBdensity { get; set; }
        public string StdItemBLight { get; set; }
        public string StdItemCdensity { get; set; }
        public string StdItemCLight { get; set; }
        public string StdItemDdensity { get; set; }
        public string StdItemDLight { get; set; }
        public string StdItemEdensity { get; set; }
        public string StdItemELight { get; set; }
        public string StdItemFdensity { get; set; }
        public string StdItemFLight { get; set; }
        public string StdItemGdensity { get; set; }
        public string StdItemGLight { get; set; }
        public string StdItemHdensity { get; set; }
        public string StdItemHLight { get; set; }
        public string S1Lot { get; set; }
        public string S1Density { get; set; }
        public string S1Uncertain { get; set; }
        public string S2Lot { get; set; }
        public string S2Density { get; set; }
        public string S2Uncertain { get; set; }
        public string C1Lot { get; set; }
        public string C1Target { get; set; }
        public string C1DensityDownLimit { get; set; }
        public string C1DensityUpLimit { get; set; }
        public string C2Lot { get; set; }
        public string C2Target { get; set; }
        public string C2DensityDownLimit { get; set; }
        public string C2DensityUpLimit { get; set; }
        public string StdCalibrateRev { get; set; }
        public string C3Lot { get; set; } = "0";
        public string C3Target { get; set; } = "0";
        public string C3DensityDownLimit { get; set; } = "0";
        public string C3DensityUpLimit { get; set; } = "0";
        public string ProductSpecification { get; set; } = "0";
        public string CapacityOfReagent { get; set; } = "0";
    }
}
