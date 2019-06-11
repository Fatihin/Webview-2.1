using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PagedList;

namespace WebView.Controllers
{
    public class FeatureMastMaintenanceController : Controller
    {
        //
        // GET: /FeatureMastMaintenance/

        public ActionResult Index()
        {
            return View();
        }
        public ActionResult AddFeature()
        {
            return View();
        }

        public ActionResult AddFeatureData(string strSent)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            System.Diagnostics.Debug.WriteLine("Fase 3 success ");
            success = myWebService.AddFeatureTable(strSent);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult List(string searchKey, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.OSPFeatureMaintenance FeatureList = new WebService._base.OSPFeatureMaintenance();

            if (searchKey != null)
            {
                if (searchKey.Equals(""))
                    FeatureList = myWebService.GetOSPFeatureMaintenanceJKH(0, 10000, null);
                else
                {
                    FeatureList = myWebService.GetOSPFeatureMaintenanceJKH(0, 10000, searchKey);
                    ViewBag.searchKey = searchKey;
                }
            }
            else
            {
                FeatureList = myWebService.GetOSPFeatureMaintenanceJKH(0, 10000, null);
                ViewBag.searchKey = "";
            }

            ViewData["data7"] = FeatureList.FeatureMaintenanceList;

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            //return View();
            int pageSize = 15;
            int pageNumber = (page ?? 1);
            return View(FeatureList.FeatureMaintenanceList.ToPagedList(pageNumber, pageSize));
        }



        public ActionResult FeatureTypeListJKH()
        {
            string FeatureTypeList = "";
            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.WV_LOOKUPTABLE
                            where p.JKH_CONTRACT == "JKH"
                            select new { p.FEATURE_TYPE };

               
                foreach (var a in query)
                {
                    FeatureTypeList = FeatureTypeList + a.FEATURE_TYPE + "|";
                }
            }
            return Json(new
            {
                FeatureTypeList = FeatureTypeList
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult MinMaterialList(string featureType)
        {
            string MinMaterialList = "";
            using (Entities ctxData = new Entities())
            {
                string lookupTable = (from a in ctxData.WV_LOOKUPTABLE
                                      where a.FEATURE_TYPE == featureType
                                      select a.LOOKUP_TABLE).Single();
                System.Diagnostics.Debug.WriteLine("LOOKUP TABLE : " + lookupTable);
                if (lookupTable == "REF_FSPLICE")
                {
                    var query = from p in ctxData.REF_FSPLICE
                                select new { p.MIN_MATERIAL };

                    foreach (var a in query)
                    {
                        MinMaterialList = MinMaterialList + a.MIN_MATERIAL + "|";
                    }
                }
                else if (lookupTable == "REF_FCBL")
                {
                    var query = from p in ctxData.REF_FCBL
                                select new { p.MIN_MATERIAL };

                    foreach (var a in query)
                    {
                        MinMaterialList = MinMaterialList + a.MIN_MATERIAL + "|";
                    }
                }
                else if (lookupTable == "REF_FDP")
                {
                    var query = from p in ctxData.REF_FDP
                                select new { p.MIN_MATERIAL };

                    foreach (var a in query)
                    {
                        MinMaterialList = MinMaterialList + a.MIN_MATERIAL + "|";
                    }
                }
                else if (lookupTable == "REF_FDC")
                {
                    var query = from p in ctxData.REF_FDC
                                select new { p.MIN_MATERIAL };

                    foreach (var a in query)
                    {
                        MinMaterialList = MinMaterialList + a.MIN_MATERIAL + "|";
                    }
                }
                else if (lookupTable == "REF_CIV_MH")
                {
                    var query = from p in ctxData.REF_CIV_MH
                                select new { p.MIN_MATERIAL };

                    foreach (var a in query)
                    {
                        MinMaterialList = MinMaterialList + a.MIN_MATERIAL + "|";
                    }
                }
                else if (lookupTable == "REF_CIV_POLE")
                {
                    var query = from p in ctxData.REF_CIV_POLE
                                select new { p.MIN_MATERIAL };

                    foreach (var a in query)
                    {
                        MinMaterialList = MinMaterialList + a.MIN_MATERIAL + "|";
                    }
                }
                else if (lookupTable == "REF_CIV_DUCTPATH")
                {
                    var query = from p in ctxData.REF_CIV_DUCTPATH
                                select new { p.MIN_MATERIAL };

                    foreach (var a in query)
                    {
                        MinMaterialList = MinMaterialList + a.MIN_MATERIAL + "|";
                    }
                }
                else if (lookupTable == "REF_CIV_RISERATT")
                {
                    var query = from p in ctxData.REF_CIV_RISERATT
                                select new { p.MIN_MATERIAL };

                    foreach (var a in query)
                    {
                        MinMaterialList = MinMaterialList + a.MIN_MATERIAL + "|";
                    }
                }
                else if (lookupTable == "REF_CIV_INNERDUCT")
                {
                    var query = from p in ctxData.REF_CIV_INNERDUCT
                                select new { p.MIN_MATERIAL };

                    foreach (var a in query)
                    {
                        MinMaterialList = MinMaterialList + a.MIN_MATERIAL + "|";
                    }
                }
                else if (lookupTable == "REF_MUX")
                {
                    var query = from p in ctxData.REF_MUX
                                select new { p.MIN_MATERIAL };

                    foreach (var a in query)
                    {
                        MinMaterialList = MinMaterialList + a.MIN_MATERIAL + "|";
                    }
                }
                else if (lookupTable == "REF_FCARD")
                {
                    var query = from p in ctxData.REF_FCARD
                                select new { p.MIN_MATERIAL };

                    foreach (var a in query)
                    {
                        MinMaterialList = MinMaterialList + a.MIN_MATERIAL + "|";
                    }
                }

            }

            string FeatureStateList = "";
            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.REF_FEATURE_STATE
                            select new { p.FEATURE_STATE };


                foreach (var a in query)
                {
                    FeatureStateList = FeatureStateList + a.FEATURE_STATE + "|";
                }
            }
           
            return Json(new
            {
                MinMaterialList = MinMaterialList,
                FeatureStateList = FeatureStateList
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult FeatureStateList()
        {
            string FeatureStateList = "";
            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.REF_FEATURE_STATE
                            select new { p.FEATURE_STATE };


                foreach (var a in query)
                {
                    FeatureStateList = FeatureStateList + a.FEATURE_STATE + "|";
                }
            }

            return Json(new
            {
                FeatureStateList = FeatureStateList
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult NewFeatureMaster(string txtFeatureType, string txtMinMaterial, string txtFeatureState, string txtDay, string txtNight, string txtWeekend, string txtHoliday)
        {
            string errorMessage = "Failed";
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            System.Diagnostics.Debug.WriteLine("DAY :" + txtDay);
            bool validateDay = myWebService.ValidateNewFeatureMastMaintenance(txtMinMaterial,txtDay, "DAY");
            bool validateNight = myWebService.ValidateNewFeatureMastMaintenance(txtMinMaterial,txtNight, "NIGHT");
            bool validateWeekend = myWebService.ValidateNewFeatureMastMaintenance(txtMinMaterial,txtWeekend, "WEEKEND");
            bool validateHoliday = myWebService.ValidateNewFeatureMastMaintenance(txtMinMaterial,txtHoliday, "HOLIDAY");

            if (validateDay && validateNight && validateWeekend && validateHoliday)
            {
                WebService._base.FeatureMaintenance newFeatMain = new WebService._base.FeatureMaintenance();
                newFeatMain.FEATURE_TYPE = txtFeatureType;
                newFeatMain.MIN_MATERIAL = txtMinMaterial;
                newFeatMain.FEATURE_STATE = txtFeatureState;
                newFeatMain.DAY = txtDay;
                newFeatMain.NIGHT = txtNight;
                newFeatMain.WEEKEND = txtWeekend;
                newFeatMain.HOLIDAY = txtHoliday;
                success = myWebService.AddFeatureMastMaintenance(newFeatMain);
            }
            else
            {
                success = false;
                errorMessage = "PU Id ";
                if (!validateDay)
                {
                    if(errorMessage != "PU Id ")
                        errorMessage += ", ";
                    errorMessage += "Day";
                }
                if (!validateNight)
                {
                    if(errorMessage != "PU Id ")
                        errorMessage += ", ";
                    errorMessage += "Night";
                }
                if (!validateWeekend)
                {
                    if(errorMessage != "PU Id ")
                        errorMessage += ", ";
                    errorMessage += "Weekend";
                }
                if (!validateHoliday)
                {
                    if(errorMessage != "PU Id ")
                        errorMessage += ", ";
                    errorMessage += "Holiday";
                }

                errorMessage += " existed.";
            }
            return Json(new
            {
                errorMessage = errorMessage,
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult DeleteFeatureMaster(string txtFeatMastNo)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.DeleteFeatureMaintenanceJKH(txtFeatMastNo);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateFeatureMaster(string txtMinMaterial, string txtFeatMastNo, string txtFeatureState, string txtDay, string txtNight, string txtWeekend, string txtHoliday)
        {
            bool success = true;
            string errorMessage = "Failed";
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            string day = "";
            string night = "";
            string weekend = "";
            string holiday = "";
            using (Entities ctxdata = new Entities())
            {
                int featNo = Convert.ToInt16(txtFeatMastNo);
                var query = from a in ctxdata.WV_FEAT_MAST
                            where a.FEAT_MAST_NO == featNo
                            select new {a.DAY, a.NIGHT, a.WEEKEND, a.HOLIDAY };
                foreach (var data in query)
                {
                    day = data.DAY.ToString();
                    night = data.NIGHT.ToString();
                    weekend = data.WEEKEND.ToString();
                    holiday = data.HOLIDAY.ToString();
                }
            }
            bool validateDay = true;
            if (day != txtDay)
            {
                validateDay = myWebService.ValidateNewFeatureMastMaintenance(txtMinMaterial, txtDay, "DAY");
            }
            bool validateNight = true;
            if (night != txtNight)
            {
                validateNight = myWebService.ValidateNewFeatureMastMaintenance(txtMinMaterial, txtNight, "NIGHT");
            }
            bool validateWeekend = true;
            if (weekend != txtWeekend)
            {
                validateWeekend = myWebService.ValidateNewFeatureMastMaintenance(txtMinMaterial, txtWeekend, "WEEKEND");
            }
            bool validateHoliday = true;
            if (holiday != txtHoliday)
            {
                validateHoliday = myWebService.ValidateNewFeatureMastMaintenance(txtMinMaterial, txtHoliday, "HOLIDAY");
            }

            if (validateDay && validateNight && validateWeekend && validateHoliday)
            {
                WebService._base.FeatureMaintenance newFeatMain = new WebService._base.FeatureMaintenance();
                newFeatMain.FEAT_MAST_NO = txtFeatMastNo;
                newFeatMain.DAY = txtDay;
                newFeatMain.NIGHT = txtNight;
                newFeatMain.WEEKEND = txtWeekend;
                newFeatMain.HOLIDAY = txtHoliday;
                newFeatMain.FEATURE_STATE = txtFeatureState;
                success = myWebService.UpdateFeatMast(newFeatMain);
            }
            else
            {
                success = false;
                errorMessage = "PU Id ";
                if (!validateDay)
                {
                    if (errorMessage != "PU Id ")
                        errorMessage += ", ";
                    errorMessage += "Day";
                }
                if (!validateNight)
                {
                    if (errorMessage != "PU Id ")
                        errorMessage += ", ";
                    errorMessage += "Night";
                }
                if (!validateWeekend)
                {
                    if (errorMessage != "PU Id ")
                        errorMessage += ", ";
                    errorMessage += "Weekend";
                }
                if (!validateHoliday)
                {
                    if (errorMessage != "PU Id ")
                        errorMessage += ", ";
                    errorMessage += "Holiday";
                }

                errorMessage += " existed.";
            }
            return Json(new
            {
                errorMessage = errorMessage,
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        
        public ActionResult ListContract(string searchKey, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.OSPFeatureMaintenance FeatureList = new WebService._base.OSPFeatureMaintenance();

            if (searchKey != null)
            {
                if (searchKey.Equals(""))
                    FeatureList = myWebService.GetOSPFeatureMaintenanceContract(0, 10000, null);
                else
                {
                    FeatureList = myWebService.GetOSPFeatureMaintenanceContract(0, 10000, searchKey);
                    ViewBag.searchKey = searchKey;
                }
            }
            else
            {
                FeatureList = myWebService.GetOSPFeatureMaintenanceContract(0, 10000, null);
                ViewBag.searchKey = "";
            }
            //System.Diagnostics.Debug.WriteLine(FeatureList.FeatureMaintenanceList[0].CONTRACT_NO + "  :  " + FeatureList.FeatureMaintenanceList[0].ITEM_NO);
            ViewData["data7"] = FeatureList.FeatureMaintenanceList;

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            //return View();
            int pageSize = 15;
            int pageNumber = (page ?? 1);
            return View(FeatureList.FeatureMaintenanceList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult FeatureTypeListContract()
        {
            string FeatureTypeList = "";
            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.WV_LOOKUPTABLE
                            where p.JKH_CONTRACT == "CONTRACT"
                            select new { p.FEATURE_TYPE };


                foreach (var a in query)
                {
                    FeatureTypeList = FeatureTypeList + a.FEATURE_TYPE + "|";
                }
            }
            return Json(new
            {
                FeatureTypeList = FeatureTypeList
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult MinMaterialListContract(string featureType)
        {
            string MinMaterialList = "";
            using (Entities ctxData = new Entities())
            {
                string lookupTable = (from a in ctxData.WV_LOOKUPTABLE
                                      where a.FEATURE_TYPE == featureType
                                      select a.LOOKUP_TABLE).Single();
                System.Diagnostics.Debug.WriteLine("LOOKUP TABLE : " + lookupTable);
                if (lookupTable == "REF_MUX")
                {
                    var query = from p in ctxData.REF_FSPLICE
                                select new { p.MIN_MATERIAL };

                    foreach (var a in query)
                    {
                        MinMaterialList = MinMaterialList + a.MIN_MATERIAL + "|";
                    }
                }
                else if (lookupTable == "REF_FCARD")
                {
                    var query = from p in ctxData.REF_FCBL
                                select new { p.MIN_MATERIAL };

                    foreach (var a in query)
                    {
                        MinMaterialList = MinMaterialList + a.MIN_MATERIAL + "|";
                    }
                }
            }

            return Json(new
            {
                MinMaterialList = MinMaterialList
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult NewFeatureMasterContract(string txtFeatureType, string txtMinMaterial, string txtContractNo, string txtItemNo)
        {
            string errorMessage = "Failed";
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            //System.Diagnostics.Debug.WriteLine("DAY :" + txtDay);
            bool validateData = myWebService.ValidateNewFeatureMastContractMaintenance(txtMinMaterial, txtContractNo, txtItemNo);

            if (validateData)
            {
                WebService._base.FeatureMaintenance newFeatMain = new WebService._base.FeatureMaintenance();
                newFeatMain.FEATURE_TYPE = txtFeatureType;
                newFeatMain.MIN_MATERIAL = txtMinMaterial;
                newFeatMain.CONTRACT_NO = txtContractNo;
                newFeatMain.ITEM_NO = txtItemNo;
                success = myWebService.AddFeatureMastContractMaintenance(newFeatMain);
            }
            else
            {
                success = false;
                errorMessage = "Contract No and Item No existed.";
            }
            return Json(new
            {
                errorMessage = errorMessage,
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult DeleteFeatureMasterContract(string txtFeatMastNo)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.DeleteFeatureMaintenanceContract(txtFeatMastNo);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateFeatureMasterContract(string txtMinMaterial, string txtFeatMastNo, string txtContractNo, string txtItemNo)
        {
            bool success = true;
            string errorMessage = "Failed";
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            string contractNo = "";
            string itemNo = "";
            using (Entities ctxdata = new Entities())
            {
                int featNo = Convert.ToInt16(txtFeatMastNo);
                var query = from a in ctxdata.WV_FEAT_MAST_CONTRACT
                            where a.FEAT_MAST_CONTRACT_NO == featNo
                            select new { a.CONTRACT_NO, a.ITEM_NO};
                foreach (var data in query)
                {
                    contractNo = data.CONTRACT_NO;
                    itemNo = data.ITEM_NO;
                }
            }
            bool validateData = true;
            if ((contractNo != txtContractNo) && (itemNo != txtItemNo))
            {
                validateData = myWebService.ValidateNewFeatureMastContractMaintenance(txtMinMaterial, txtContractNo, txtItemNo);
            }
            if (validateData)
            {
                WebService._base.FeatureMaintenance newFeatMain = new WebService._base.FeatureMaintenance();
                newFeatMain.FEAT_MAST_CONTRACT_NO = txtFeatMastNo;
                newFeatMain.CONTRACT_NO = txtContractNo;
                newFeatMain.ITEM_NO = txtItemNo;
                success = myWebService.UpdateFeatMastContract(newFeatMain);
            }
            else
            {
                success = false;
                errorMessage = "Contract No and Item No existed.";
            }
            return Json(new
            {
                errorMessage = errorMessage,
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

    }
}
