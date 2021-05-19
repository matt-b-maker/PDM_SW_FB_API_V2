using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDM_SW_FB_API_V2
{
    class MaterialItem
    {

        public MaterialItem(string materialName, int quantity, string partNo)
        {
            this.MaterialName = materialName;
            this.Quantity = quantity;
            this.PartNo = partNo;
        }

        public string MaterialName { get; set; }

        public int Quantity { get; set; }

        public string PartNo { get; set; }

    }
}
