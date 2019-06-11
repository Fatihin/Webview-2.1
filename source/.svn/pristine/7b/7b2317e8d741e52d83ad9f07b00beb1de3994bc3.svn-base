using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NEPS.BOQ.Utilities;
using Oracle.DataAccess.Client;
using WebView.Controllers;

namespace NEPS.BOQ.Classes.ContractBOQ
{
    public class Engine
    {
        public static void Run(bool writeOutputFile,
            OracleConnection conn,
            string schemeName,
            string packagesToCreate,
            out List<Package> packages,
            out List<Item> individualItems)
        {
            packages = new List<Package>();
            individualItems = new List<Item>();

            List<Item> source = GetSourceItems(conn, schemeName);
            //List<Item> source = TestGetSourceItems(conn, schemeName);

            // copy source items into individual items
            foreach (Item item in source)
            {
                Item copyItem = item.Copy();
                individualItems.Add(copyItem);
            }

            // create packages, one package per contract number
            char[] commaDelimiter = { ',' };
            string[] arPackagesToCreate = packagesToCreate.Split(commaDelimiter);
            Dictionary<string, string> contractNos = new Dictionary<string, string>();
            foreach (Item item in source)
            {
                string contractNo = item.ContractNo;
                if (!contractNos.ContainsKey(contractNo))
                {
                    foreach (string packageItemNo in arPackagesToCreate)
                        CreateItem(individualItems, contractNo, packageItemNo, 1, "");
                    contractNos[contractNo] = "x"; // does not matter what you keep there. this is just to indicate that the contract has been catered for
                }
            }

            // extract packages out of the individualItems
            List<Item> toRemove = new List<Item>();
            foreach (Item item in individualItems)
            {
                List<Package> createdPackages = GetPackages(conn, item, arPackagesToCreate);
                if (createdPackages != null)
                {
                    packages.AddRange(createdPackages);
                    toRemove.Add(item);
                }
            }
            foreach (Item item in toRemove)
                individualItems.Remove(item); // remove the items that have been identified as packages

            // by this time, we have separated the items into packages and individual items
            // so, we assign the individual items into the packages
            foreach (Package package in packages)
            {
                if (package.CanBeFilled(individualItems))
                    package.Fill(individualItems);

            }

            // remove items that have been taken by the packages
            toRemove = new List<Item>();
            foreach (Item item in individualItems)
            {
                if (item.Quantity == 0)
                    toRemove.Add(item);
            }
            foreach (Item item in toRemove)
            {
                individualItems.Remove(item);
            }

            // remove incomplete packages from list
            List<Package> completePackages = new List<Package>();
            foreach (Package package in packages)
            {
                if (package.HasBeenFilled)
                    completePackages.Add(package);
            }
            packages = completePackages;

            //if (writeOutputFile)
            //    WriteOutput("Output.txt", source, packages, individualItems, packagesToCreate);

            // delete the source items from database
            string deletesql = "DELETE FROM WV_BOQ_DATA WHERE SCHEME_NAME='" + schemeName + "'";
            UtilityDb2.ExecuteSql(deletesql, conn);


        }

        private static List<Item> GetSourceItems(OracleConnection conn, string schemeName)
        {
            List<Item> output = new List<Item>();
            string sql = "SELECT * FROM WV_BOQ_DATA WHERE SCHEME_NAME='" + schemeName + "'";
            using (OracleDataReader dr = UtilityDb2.GetDataReader(sql, conn))
            {
                while (dr.Read())
                {
                    Item item = new Item();
                    item.ContractNo = dr["CONTRACT_NO"].ToString();
                    item.ItemNo = dr["ITEM_NO"].ToString();
                    item.Quantity = Convert.ToSingle(dr["PU_QTY"]);

                    string isposp = dr["ISP_OSP"].ToString();
                    if (isposp.ToUpper() == "ISP")
                        item.CotRt = "COT";
                    else if (isposp.ToUpper() == "OSP")
                        item.CotRt = "RT";

                    output.Add(item);
                }
            }
            return output;
        }

        private static List<Item> TestGetSourceItems(OracleConnection conn, string schemeName)
        {
            List<Item> output = new List<Item>();
            string contractNumber = "3400005199";
            Item item = null;


            // create items
            item = new Item();
            item.ItemNo = "15";
            item.Quantity = 5;
            item.CotRt = "COT";
            item.ContractNo = contractNumber;
            output.Add(item);           
            
            item = new Item();
            item.ItemNo = "15";
            item.Quantity = 5;
            item.CotRt = "RT";
            item.ContractNo = contractNumber;
            output.Add(item);


            item = new Item();
            item.ItemNo = "29";
            item.Quantity = 5;
            item.CotRt = "COT";
            item.ContractNo = contractNumber;
            output.Add(item);

            item = new Item();
            item.ItemNo = "29";
            item.Quantity = 5;
            item.CotRt = "RT";
            item.ContractNo = contractNumber;
            output.Add(item);

            return output;
        }


        private static void WriteOutput(string filename, List<Item> source, List<Package> packages, List<Item> individualItems, string packagesToCreate)
        {
            using (System.IO.TextWriter file = System.IO.File.CreateText(filename))
            {
                file.WriteLine("SPECIFIED PACKAGES: " + packagesToCreate);

                file.WriteLine("1. SOURCE ITEMS (items from Kamal):");
                foreach (Item item in source)
                {
                    file.WriteLine("Contract No " + item.ContractNo + ", Item No " + item.ItemNo + ", C/R " + item.CotRt + ", Quantity " + item.Quantity);
                }
                file.WriteLine();

                file.WriteLine("2. PACKAGES (items identified as packages): ");
                foreach (Package package in packages)
                {
                    string completionStatus = package.HasBeenFilled ? "This package is complete." : "*** This package cannot be completed and will not be sent in output.";
                    file.WriteLine("Package Item No " + package.OriginalItem.ItemNo + ", " + completionStatus);
                    file.WriteLine("   Specified Items (items belonging to this package as specified in WV_PACKAGE_MAST):");
                    foreach (Package.PackageItem specItem in package.SpecifiedItems)
                    {
                        file.WriteLine("      Contract No " + specItem.ContractNo + ", Item No " + specItem.ItemNo + ", C/R " + specItem.CotRt + ", Quantity " + specItem.Quantity + ", Mandatory " + specItem.Mandatory.ToString());
                    }
                    file.WriteLine("   Content Items (items assigned to this package):");

                    foreach (Package.ContentItem contentItem in package.ContentItems)
                    {
                        file.WriteLine("      Contract No " + contentItem.OriginalItem.ContractNo + ", Item No " + contentItem.OriginalItem.ItemNo + ", C/R " + contentItem.OriginalItem.CotRt + ", Quantity " + contentItem.Quantity);
                    }

                    file.WriteLine();
                }
                file.WriteLine();

                file.WriteLine("3. INDIVIDUAL ITEMS LEFT: (items not assigned to packages) (");
                foreach (Item item in individualItems)
                {
                    file.WriteLine("Contract No " + item.ContractNo + ", Item No " + item.ItemNo + ", C/R " + item.CotRt + ", Quantity " + item.Quantity);
                }
                file.WriteLine();

            }
        }

        public static void CreateItem(List<Item> source, string contractNo, string itemNo, int quantity, string COTRT)
        {
            Item item = new Item();
            item.ItemNo = itemNo;
            item.CotRt = COTRT;
            item.Quantity = quantity;
            item.ContractNo = contractNo;
            source.Add(item);
        }

        private static List<Package> GetPackages(OracleConnection conn, 
            Item originalItem, string [] userSpecifiedPackages)
        {
            List<Package> packages = new List<Package>();

            // create packages according to the quantity of the item.
            for (int i = 0; i < originalItem.Quantity; i++)
            {
                Package package = new Package();
                package.OriginalItem = originalItem;
                packages.Add(package);
            }

            // find items for the packages
            bool hasItems = false;
            string sql = "SELECT * FROM WV_PACKAGE_MAST WHERE MANDATORY='Y' AND CONTRACT_NO='" +
                originalItem.ContractNo + "' AND PACKAGE_ITEM_NO='" + originalItem.ItemNo + "'";
            using (OracleDataReader dr = UtilityDb2.GetDataReader(sql, conn))
            {
                while (dr.Read())
                {
                    hasItems = true;

                    foreach (Package package in packages)
                    {
                        Package.PackageItem specItem = new Package.PackageItem();
                        specItem.ItemNo = dr["ITEM_NO"].ToString();
                        specItem.Quantity = Convert.ToSingle(dr["ITEM_QTY"]);
                        specItem.Mandatory = dr["MANDATORY"].ToString().ToUpper() == "Y";
                        specItem.CotRt = dr["COT-RT"].ToString();
                        specItem.ItemType = dr["TYPE"].ToString();
                        specItem.ContractNo = dr["CONTRACT_NO"].ToString();
                        package.SpecifiedItems.Add(specItem);
                    }
                }
            }

            // return the list of created packages if items were found
            // otherwise return null

            foreach (string strPackageItemNo in userSpecifiedPackages)
            {
                if (strPackageItemNo.Trim() == originalItem.ItemNo)
                    return packages;
            }

            return null;

         
        }

    }
}
