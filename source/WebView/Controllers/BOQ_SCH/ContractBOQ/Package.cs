using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WebView.Controllers;

namespace NEPS.BOQ.Classes.ContractBOQ
{
    public class Package
    {
        public class PackageItem:Item
        {
            public bool Mandatory { get; set; }
            public string ItemType { get; set; }
        
            public bool Matches(Item item, bool mandatoryOnly)
            {
                if (mandatoryOnly && !Mandatory)
                    return false;

                return ItemNo == item.ItemNo && CotRt == item.CotRt && ContractNo == item.ContractNo;
            }


            public bool FoundSpecifiedItem { get; set; }
        }

        public class ContentItem
        {
            public Item OriginalItem { get; set; }
            public float Quantity { get; set; }
        }

        public Item OriginalItem { get; set; } // the original item from which this package was derived
        public List<PackageItem> SpecifiedItems { get; set; } // specification items as in WV_PACKAGE_MAST table
        public List<ContentItem> ContentItems { get; set; } // actual content after assignment
        public bool HasBeenFilled { get; set; }

        public Package()
        {
            SpecifiedItems = new List<PackageItem>();
            ContentItems = new List<ContentItem>();
            HasBeenFilled = false;
            
        }



        internal bool Specifies(Item item, bool mandatoryOnly, out float specifiedQuantity)
        {
            foreach (PackageItem specItem in SpecifiedItems)
            {
                if (specItem.Matches(item, mandatoryOnly))
                {
                    specifiedQuantity = specItem.Quantity;
                    return true;
                }
            }

            specifiedQuantity = 0;
            return false;
        }

        internal void AddContentItem(Item item, float assignedQuantity)
        {
            ContentItem contentItem = new ContentItem();
            contentItem.OriginalItem = item;
            contentItem.Quantity = assignedQuantity;
            ContentItems.Add(contentItem);
            item.Quantity -= assignedQuantity;
        }

        internal bool CanBeFilled(List<Item> individualItems)
        {
            // if no items specified just return false
            if (SpecifiedItems.Count == 0)
                return false;

            // if there items that are less than the specified quantity
            // in the package, the package cannot be filled.
            foreach (PackageItem specifiedItem in SpecifiedItems)
                specifiedItem.FoundSpecifiedItem = false;

            foreach (Item item in individualItems)
            {
                foreach (PackageItem specifiedItem in SpecifiedItems)
                {
                    if (specifiedItem.Matches(item, true))
                    {
                        specifiedItem.FoundSpecifiedItem = true;
                        if (specifiedItem.Quantity > item.Quantity)
                            return false;
                    }
                }           
            }

            foreach (PackageItem specifiedItem in SpecifiedItems)
                if (specifiedItem.Mandatory && !specifiedItem.FoundSpecifiedItem)
                    return false;
  
            return true;
        }

        internal void Fill(List<Item> individualItems)
        {
            /// assumption, this is called after a CanBeCompleted function has been called
            /// i.e. it has been determined that there are adequate items to fill the package
            foreach (Item item in individualItems)
            {
                float specifiedQuantity = 0;
                if (Specifies(item, false, out specifiedQuantity))
                    AddContentItem(item, specifiedQuantity);
            }
            HasBeenFilled = true;

        }
    }
}
