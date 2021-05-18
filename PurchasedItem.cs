using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDM_SW_FB_API_V2
{
    class PurchasedItem
    {
        public PurchasedItem(string partNo, int quantity)
        {
            this.PartNo = partNo;
            this.Quantity = quantity;
        }

        public string PartNo { get; set; }

        public int Quantity { get; set; }
    }
}
