using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WebView.Controllers
{
    public class Item
    {
        public string ItemNo { get; set; }
        public float Quantity { get; set; }
        public string CotRt { get; set; }
        public string ContractNo { get; set; }

        public Item Copy()
        {
            Item output= new Item();
            output.ItemNo = ItemNo;
            output.Quantity = Quantity;
            output.CotRt = CotRt;
            output.ContractNo = ContractNo;
            return output;
        }


        internal bool Match(Item item)
        {
            return ItemNo == item.ItemNo && CotRt == item.CotRt && item.ContractNo == ContractNo;

        }
    }
}
