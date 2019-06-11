using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NEPS.BOQ.Utilities;
using System.Data;
using Oracle.DataAccess.Client;
using WebView.BOQNrmService;

namespace NEPS.BOQ.Classes
{
    class BOQNrmEngine
    {
        private class Item
        {
            // copied straight from web service definition of BOQItem
            public string schemeName;
            public string contractNo;
            public string itemNo;
            public string materialID;
            public string featureState;
            public string quantity;
        }

        public bool GenerateBOQItems(string schemeName, OracleConnection conn, ref string errorMessage)
        {
            try
            {

                WebView.BOQNrmService.nrmboqClient svc = new WebView.BOQNrmService.nrmboqClient();

                string errorCode = "";
                string errorDesc = "";
                BOQItemList itemList = svc.generateBOQItems(out errorCode, out errorDesc, schemeName);
                List<Item> expandedList = ExpandWebServiceItems(itemList);

                if (string.IsNullOrEmpty(errorCode)) // no errors
                {
                    // insert into db
                    string sql = "DELETE FROM WV_ISP_TEMP WHERE SCHEME_NAME='" + schemeName + "'";
                    UtilityDb2.ExecuteSql(sql, conn);
                    using (UtilityDb2 db = new UtilityDb2())
                    {
                        db.connection = conn;

                        db.PrepareInsert("WV_ISP_TEMP");
                        foreach (Item item in expandedList)
                        {
                            DataRow row = db.Insert(null);
                            row["SCHEME_NAME"] = item.schemeName;
                            row["CONTRACT_NO"] = item.contractNo;
                            row["ITEM_NO"] = item.itemNo;

                            int quantity = 0;
                            if (int.TryParse(item.quantity, out quantity))
                                row["QUANTITY"] = quantity.ToString();

                            row["FEATURE_STATE"] = item.featureState;
                            db.Insert(row);
                        }
                        db.EndInsert();
                    }
                    return true;
                }
                else
                {
                    errorMessage = errorCode + ": " + errorDesc;
                    return true;
                }
            }
            catch (Exception ex)
            {
                errorMessage = "Exception: " + ex.Message;
                return false;
            }
        }

        private List<Item> ExpandWebServiceItems(BOQItemList wsItems)
        {
            List<Item> output = new List<Item>();
            foreach (BOQItem wsItem in wsItems.boqItem)
            {
                // create a list of item numbers, one from the main item number
                // and add all numbers in the material ID array.
                List<string> itemNumbers = new List<string>();
                itemNumbers.Add(wsItem.itemNo);

                // get the material ID
                string strMaterialIDs = wsItem.materialID;
                List<string> expandedMaterialIds = GetMaterialIDs(strMaterialIDs);

                itemNumbers.AddRange(expandedMaterialIds);

                // for each item number, create a new item in the output.
                foreach (string itemNumber in itemNumbers)
                {
                    string trimmedItemNumber = itemNumber.Trim();
                    if (!string.IsNullOrEmpty(trimmedItemNumber))
                    {
                        Item newItem = new Item();
                        newItem.contractNo = wsItem.contractNo;
                        newItem.featureState = wsItem.featureState;
                        newItem.itemNo = trimmedItemNumber;
                        newItem.quantity = wsItem.quantity;
                        newItem.schemeName = wsItem.schemeName;
                        output.Add(newItem);
                    }
                }
            }
            return output;
        }

        public static List<string> GetMaterialIDs(string strMaterialIDs)
        {
            if (string.IsNullOrEmpty(strMaterialIDs.Trim()))
                return new List<string>();

            char[] commaDelimiter = { ',' };
            string[] materialIds = strMaterialIDs.Split(commaDelimiter);

            // for material IDs it is possible to specify nn x qty
            // so expand it
            List<string> expandedMaterialIds = new List<string>();
            char[] multiplyDelimiter = { 'x', 'X' };
            foreach (string materialId in materialIds)
            {
                if (string.IsNullOrEmpty(materialId.Trim()))
                    continue;

                int qty = 1;
                string actualMaterialId = materialId;
                string[] parts = materialId.Split(multiplyDelimiter);

                // if a multiplier has been specified
                if (parts.Length == 2)
                {
                    actualMaterialId = parts[0];
                    int.TryParse(parts[1], out qty);
                }

                for (int count = 1; count <= qty; count++)
                    expandedMaterialIds.Add(actualMaterialId);
            }
            return expandedMaterialIds;
        }
    }
}
