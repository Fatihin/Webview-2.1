using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WebView.Controllers;
using System.Data.OleDb;
using System.Data;

namespace WebView.Controllers

{
    class BOQEngine
    {
  
        public void StartProcess(string schemeName, OleDbConnection conn)
        {

                // get the min material
                List<string> minMaterials = new List<string>();
                string sql = "SELECT min_material FROM gc_netelem WHERE scheme_name='" + schemeName + "'";
                using (OleDbDataReader dr = UtilityDb.GetDataReader(sql, conn))
                {
                    while (dr.Read())
                    {
                        minMaterials.Add(dr["min_material"].ToString());
      
                    }
                }

                foreach (string mm in minMaterials)
                {
                    string minMaterial = mm;

                    // TEST CODE
                   // minMaterial = "24F FRA LOOSE AT";

                    // find lookup table. This is used for SCH processing later.
                    string lookupTableName = GetLookupTableName(conn, minMaterial);

                    // fetch the attribute mappings
                    Dictionary<string, string> mapping =
                        GetAttributeMapping(conn, sql, lookupTableName, minMaterial);

                    // find matches from the SCH table
                    List<SCHMatches> matches = GetMatches(conn, mapping);

                    // TEST CODE
                    //foreach (SCHMatches o in matches)
                    //{
                    //    o.PU_ID = "105987";
                    //    o.BILL_RATE = "D";
                    //}

                    // Fetch PU Master records
                    List<PUMaster> PUMasterRecords = GetPUMasterRecords(conn, matches);


                    // Insert into BOQ Data
                    InsertBOQData(conn, PUMasterRecords, schemeName, "OSP");
                }

        }


        private void InsertBOQData(OleDbConnection conn, List<PUMaster> PUMasterRecords,
            string schemeName, string ISPOSP)
        {

            UtilityDb db = new UtilityDb();
            db.connection = conn;

            db.PrepareInsert("WV_BOQ_DATA");
            foreach (PUMaster item in PUMasterRecords)
            {
                // first check for duplication.
                // find out if the same PU_ID and Billing_date already exists in the WV_BOQ_DATA table. If not, insert.

                if (BOQDataExists(conn, item.SCHReference.PU_ID, item.SCHReference.BILL_RATE))
                    continue;                

                DataRow row = db.Insert(null);
                row["SCHEME_NAME"] = schemeName;
                row["PU_ID"] = item.SCHReference.PU_ID;
                row["PU_DESC"] = item.PU_DESC;
                row["RATE_INDICATOR"] = item.SCHReference.BILL_RATE;
                row["BQ_MAT_PRICE"] = item.PU_MAT_PR;
                row["BQ_INSTALL_PRICE"] = item.PU_INST_PR;
                row["PU_QTY"] = 0; // need to count all records inside  with the same PU_ID and BILLING_RATE
                row["PU_UOM"] = item.PU_UOM;
                row["ISP_OSP"] = ISPOSP;
                db.Insert(row);
            }
            db.EndInsert();
        }

        private bool BOQDataExists(OleDbConnection conn, string PU_ID, string BILL_RATE)
        {
            string sql = string.Format("SELECT COUNT(*) FROM WV_BOQ_DATA WHERE TRIM(PU_ID)='{0}' AND RATE_INDICATOR='{1}'",
                    PU_ID, BILL_RATE);

            int count = Convert.ToInt32(UtilityDb.ExecuteScalar(sql, conn));
            return count > 0;
              
        }

        private List<PUMaster> GetPUMasterRecords(OleDbConnection conn, List<SCHMatches> matches)
        {
            List<PUMaster> output = new List<PUMaster>();

            foreach (SCHMatches sch in matches)
	        {
                string sql = string.Format("SELECT * FROM WV_PU_MAST WHERE TRIM(PU_ID)='{0}' AND BILL_RATE='{1}'",
                    sch.PU_ID, sch.BILL_RATE);

                using (OleDbDataReader dr = UtilityDb.GetDataReader(sql, conn))
                {
                    while (dr.Read())
                    {
                        PUMaster obj = new PUMaster();
                        obj.PU_DESC = dr["PU_DESC"].ToString();
                        obj.PU_INST_PR = Convert.ToSingle(dr["PU_INST_PR"]);
                        obj.PU_MAT_PR = Convert.ToSingle(dr["PU_MAT_PR"]);
                        obj.PU_UOM = dr["PU_UOM"].ToString();
                        obj.SCHReference = sch;
                        output.Add(obj);
                    }
                }
	        }
          

            return output;
        }

        private List<SCHMatches> GetMatches(OleDbConnection conn, 
            Dictionary<string, string> mapping)
        {
            List<SCHMatches> output = new List<SCHMatches>();

            string where = "";
            int i=0;
            foreach (string key in mapping.Keys)
            {
                string s = string.Format("(UPPER({0}) = UPPER('{1}') OR {0} = '-')",
                    key, mapping[key]);
                if (i != 0)
                    where += " AND "; 
                where += s;
                i++;
            }

            if (string.IsNullOrEmpty(where))
                return output;

            string sql = string.Format("SELECT * FROM WV_FEAT_SCH WHERE {0}", 
                where);
            using (OleDbDataReader dr = UtilityDb.GetDataReader(sql, conn))
            {
                while (dr.Read())
                {
                    SCHMatches match = new SCHMatches();
                    if (dr["MUL_FAC"] != DBNull.Value)
                        match.MUL_FAC = Convert.ToSingle(dr["MUL_FAC"]);

                    if (dr["PU_ID"] != DBNull.Value)
                        match.PU_ID = dr["PU_ID"].ToString();

                    if (dr["QTY_IND"] != DBNull.Value)
                        match.QTY_IND = dr["QTY_IND"].ToString();

                    if (dr["BILL_RATE"] != DBNull.Value)
                        match.BILL_RATE = dr["BILL_RATE"].ToString();

                    output.Add(match);
                }
            }

            return output;
        }

        private static string GetLookupTableName(OleDbConnection conn, string minMaterial)
        {
            string sql = string.Format("SELECT * FROM WV_FEAT_MAST WHERE min_material='{0}'", minMaterial);
            using (OleDbDataReader dr = UtilityDb.GetDataReader(sql, conn))
            {
                while (dr.Read())
                {
                    return dr["LOOKUP_TABLE"].ToString();
                }
            }
            return null;
        }

        private static Dictionary<string, string> GetAttributeMapping(OleDbConnection conn, 
            string sql, string lookupTableName, string minMaterial)
        {
            Dictionary<string, string> attMapping = new Dictionary<string, string>();
            const int maxAttColumn = 12;

            // determine the fields to look into
            sql = string.Format("SELECT * FROM wv_lookup_table WHERE lookup_table='{0}'", lookupTableName);
            using (OleDbDataReader dr = UtilityDb.GetDataReader(sql, conn))
            {
                while (dr.Read())
                {
                    for (int i = 1; i < maxAttColumn; i++)
                    {
                        string attFieldName = "ATT" + i;
                        for (int columnIndex = 0; columnIndex < dr.FieldCount; columnIndex++)
                        {
                            string columnName = dr.GetName(columnIndex);
                            if (columnName == attFieldName)
                            {
                                attMapping[columnName] = dr[columnName].ToString();
                                break;
                            }
                        }      
                    }
                    break;
                }
            }

            // TEST CODE
           // minMaterial = "207251/N";

            // fetch the actual values
            Dictionary<string, string> attMappingWithValues = new Dictionary<string, string>();
            sql = string.Format("SELECT * FROM {0} WHERE MIN_MATERIAL='{1}'",
                lookupTableName, minMaterial);
            using (OleDbDataReader dr = UtilityDb.GetDataReader(sql, conn))
            {
                while (dr.Read())
                {
                    foreach (string attFieldName in attMapping.Keys)
                    {
                        string actualFieldName = attMapping[attFieldName];
                        if (dr.HasColumn(actualFieldName))

                            attMappingWithValues[attFieldName] = 
                                dr[actualFieldName].ToString();
                    }
                }
            }

            return attMappingWithValues;
        }
    }
}
