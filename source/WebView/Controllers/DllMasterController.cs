using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using PagedList;
using WebView.Models;

namespace WebView.Controllers
{
    public class DllMasterController : Controller
    {
        public ActionResult List(string searchKey, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.ListDllMaster DllM = new WebService._base.ListDllMaster();
            System.Diagnostics.Debug.WriteLine(searchKey);
            System.Diagnostics.Debug.WriteLine("a");
            if (searchKey != null)
            {
                if (searchKey.Equals(""))
                {
                    DllM = myWebService.GetDllMasterMaintenance(0, 1000, null);
                    ViewBag.searchKey = searchKey;
                }
                else
                {
                    DllM = myWebService.GetDllMasterMaintenance(0, 1000, searchKey);
                    ViewBag.searchKey = searchKey;
                }
            }
            else
            {
                DllM = myWebService.GetDllMasterMaintenance(0, 1000, null);
                ViewBag.searchKey = "";
            }

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            int pageSize = 10;
            int pageNumber = (page ?? 1);
            return View(DllM.DllMasterList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult DllMaster_NewJob()
        {
            return View(new WebView.Models.DllMaster_NewJobModel());
        }

        [HttpPost]
        public ActionResult DllMaster_NewJob(DllMaster_NewJobModel model)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;

            WebService._base.DLLMaster DllMaster_newjob = new WebService._base.DLLMaster();
            DllMaster_newjob.DLL_NAME = model.DLL_NAME;
            DllMaster_newjob.DLL_DESCRIPTION = model.DLL_DESCRIPTION;
            DllMaster_newjob.CREATE_DATE = model.CREATE_DATE;
            success = myWebService.AddDllMaster(DllMaster_newjob);

            selected = true;

            if (ModelState.IsValid && selected)
            {
                if (success == true)
                    return RedirectToAction("NewSave");
                else
                    return RedirectToAction("NewSaveFail"); // store to db failed.
            }

            return View(model);
        }

        public ActionResult NewSave()
        {
            return View();
        }

        public ActionResult GetDetails_DllMaster(string id)
        {
            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_DLL_MASTER
                             where p.DLL_NAME == id
                             select p).Single();

                return Json(new
                {

                    Success = true,
                    name = id,
                    desc = query.DLL_DESCRIPTION,
                    createdAt = query.CREATE_DATE.ToString()
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        public ActionResult UpdateData_DllMaster(string txtName, string txtDesc)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.DLLMaster DllM_newjob = new WebService._base.DLLMaster();
            DllM_newjob.DLL_NAME = txtName;
            DllM_newjob.DLL_DESCRIPTION = txtDesc;
            success = myWebService.UpdateDllMaster(DllM_newjob);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult DeleteData(string targetId)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.Delete_DllMaster(targetId);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateDataVersion(string txtMasterV, string txtNameV, string txtDescV, string txtVersionV)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.DLLVersion upt_data = new WebService._base.DLLVersion();
            upt_data.DLL_MASTER = txtMasterV;
            upt_data.DLL_NAME = txtNameV;
            upt_data.DLL_VERSION = txtVersionV;
            upt_data.DLL_DESCRIPTION = txtDescV;
            success = myWebService.UpdateDllVersion(upt_data);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult BatchList(string FILENAME, string SERVICENAME, string HASERROR, string TYPE, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.BiBatch BiBatchMain = new WebService._base.BiBatch();

            if (FILENAME != null || SERVICENAME != null)
            {
                if (FILENAME.Equals("Select") && SERVICENAME.Equals("Select") && HASERROR.Equals("false") && TYPE.Equals("Select"))
                {
                    BiBatchMain = myWebService.GetBatch(0, 100000, null, null, null, null);
                }
                else
                {
                    BiBatchMain = myWebService.GetBatch(0, 100000, FILENAME, SERVICENAME, HASERROR, TYPE);
                }
            }
            else
            {
                BiBatchMain = myWebService.GetBatch(0, 100000, null, null, null, null);
            }

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            using (EntitiesNetworkElement ctxData = new EntitiesNetworkElement())
            {
                //filter FILENAME
                List<SelectListItem> list3 = new List<SelectListItem>();
                var queryFILENAME = from p in ctxData.BI_BATCH
                                    select new { Text = p.INSTANCE_ID, Value = p.INSTANCE_ID };

                list3.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in queryFILENAME.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list3.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }
                ViewBag.FILENAME = list3;

                //filter servicename
                List<SelectListItem> list2 = new List<SelectListItem>();
                var querySERVICENAME = from p in ctxData.BI_BATCH
                                       select new { Text = p.SERVICE_NAME, Value = p.SERVICE_NAME };

                list2.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in querySERVICENAME.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list2.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }
                ViewBag.SERVICENAME = list2;

                List<SelectListItem> list1 = new List<SelectListItem>();
                list1.Add(new SelectListItem() { Text = "", Value = "Select" });
                list1.Add(new SelectListItem() { Text = "INBOUND", Value = "INBOUND" });
                list1.Add(new SelectListItem() { Text = "OUTBOUND", Value = "OUTBOUND" });
                ViewBag.TYPE = list1;

                if (HASERROR == "true")
                {
                    ViewBag.HASERROR = "true";
                }
                else { ViewBag.HASERROR = "false";  }
            }

            int pageSize = 15;
            int pageNumber = (page ?? 1);
            ViewBag.FILENAME_1 = FILENAME;
            ViewBag.SERVICENAME_1 = SERVICENAME;
            ViewBag.HASERROR_1 = HASERROR;
            ViewBag.TYPE_1 = TYPE;
            return View(BiBatchMain.bacthList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult GetDetails_Batch(int id)
        {
            string outputATag = "";
            string outputATagNis = "";
            string output046 = "";
            string output047 = "";
            string output048 = "";
            string output049 = "";
            string output309 = "";
            string output310 = "";
            //string soutput310 = "";
            string output311 = "";
            string output312 = "";
            string output589 = "";
            string output591 = "";
            
            using (EntitiesNetworkElement ctxData = new EntitiesNetworkElement())
            {
                

                var query = (from p in ctxData.BI_BATCH
                             where p.BATCH_ID == id
                             select p).Single();

                var checkATag = (from a in ctxData.BI_ASSET_TAGGING
                                where a.BI_BATCH_ID == id
                                select a).Count();

                var checkATagNis = (from a in ctxData.BI_ASSET_TAGGING_NIS
                                where a.BI_BATCH_ID == id
                                select a).Count();

                var check046 = (from a in ctxData.BI_NNW_PU_MAST
                                where a.BI_BATCH_ID == id
                                select a).Count();
                var check047 = (from a in ctxData.BI_NNW_MAT_MAST
                                where a.BI_BATCH_ID == id
                                select a).Count();
                var check048 = (from a in ctxData.BI_NNW_PU_MAT_REF
                                where a.BI_BATCH_ID == id
                                select a).Count();
                var check049 = (from a in ctxData.BI_CONTRACT
                                where a.BI_BATCH_ID == id
                                select a).Count();
                var check309 = (from a in ctxData.BI_NNW_EXT_SCH_GEM
                                where a.BI_BATCH_ID == id
                                select a).Count();
                var check310 = (from a in ctxData.BI_NNW_GEM_PROJNO
                                where a.BI_BATCH_ID == id
                                select a).Count();
                var check311 = (from a in ctxData.BI_TOGEMS_SEQ1
                                where a.BI_BATCH_ID == id
                                select a).Count();
                var check312 = (from a in ctxData.BI_GEMS_BOQ_MAT
                                where a.BI_BATCH_ID == id
                                select a).Count();
                var check589 = (from a in ctxData.BI_ASSET_NO
                                where a.BI_BATCH_ID == id
                                select a).Count();
                var check591 = (from a in ctxData.BI_DECOMM_TABLE
                                where a.BI_BATCH_ID == id
                                select a).Count();

                if (checkATag > 0)
                {
                    var queryATag = (from a in ctxData.BI_ASSET_TAGGING
                                        where a.BI_BATCH_ID == id
                                        select new
                                        {
                                            a.ASSET_NO,
                                            a.CABLE_EQUIP_ID,
                                            a.TAGGING_DATE,
                                            a.LAST_INVENTORY_ON,
                                            a.EXC_ABB,
                                            a.TYPE_NAME,
                                            a.VERIFICATION_STATUS
                                        });
                    System.Diagnostics.Debug.WriteLine(queryATag);
                    if (queryATag.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryATag)
                        {
                            counter++;

                            string carry = "";
                            outputATag += lp.ASSET_NO.ToString() + "|" + lp.CABLE_EQUIP_ID.ToString() + "|" + lp.TAGGING_DATE.ToString() + "|" + lp.LAST_INVENTORY_ON.ToString() + "|" + lp.EXC_ABB + "|" + lp.TYPE_NAME + "|" + lp.VERIFICATION_STATUS;

                            outputATag += "!";
                            System.Diagnostics.Debug.WriteLine("agak:" + outputATag);
                        }
                    }
                }

                if (checkATagNis > 0)
                {
                    var queryATagNis = (from a in ctxData.BI_ASSET_TAGGING_NIS
                                    where a.BI_BATCH_ID == id
                                    select new
                                    {
                                        a.NIS_ID,
                                        a.ASSET_NO,
                                        a.TAGGING_DATE,
                                        a.LAST_INVENTORY_ON,
                                        a.EXCHANGE_CODE,
                                        a.TYPE_NAME,
                                        a.VERIFICATION_STATUS
                                    });
                    System.Diagnostics.Debug.WriteLine(queryATagNis);
                    if (queryATagNis.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryATagNis)
                        {
                            counter++;

                            string carry = "";
                            outputATagNis += lp.NIS_ID.ToString() + "|" + lp.ASSET_NO.ToString() + "|" + lp.TAGGING_DATE.ToString() + "|" + lp.LAST_INVENTORY_ON.ToString() + "|" + lp.EXCHANGE_CODE + "|" + lp.TYPE_NAME + "|" + lp.VERIFICATION_STATUS;

                            outputATagNis += "!";
                            System.Diagnostics.Debug.WriteLine("agak:" + outputATagNis);
                        }
                    }
                }

                if (check046 > 0)
                {
                    var query046 = (from a in ctxData.BI_NNW_PU_MAST
                                    where a.BI_BATCH_ID == id
                                    select new
                                    {
                                        a.PU_ID,
                                        a.PU_DESC,
                                        a.PU_INST_PR,
                                        a.PU_UOM,
                                        a.NETWORK_FLAG,
                                        a.BILL_RATE
                                    });
                    System.Diagnostics.Debug.WriteLine(query);
                    if (query046.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in query046)
                        {
                            counter++;

                            string carry = "";
                            output046 += lp.PU_ID.ToString() + "|" + lp.PU_DESC.ToString() + "|" + lp.PU_INST_PR.ToString() + "|" + lp.PU_UOM.ToString() + "|" + lp.NETWORK_FLAG + "|" + lp.BILL_RATE;

                            output046 += "!";
                            System.Diagnostics.Debug.WriteLine("agak:" + output046);
                        }
                    }
                }

                if (check047 > 0)
                {
                    var query047 = (from a in ctxData.BI_NNW_MAT_MAST
                                    where a.BI_BATCH_ID == id
                                    select new
                                    {
                                        a.MAT_ID,
                                        a.MAT_NAME,
                                        a.MAT_UOM,
                                        a.MAT_CAT,
                                        a.MAT_PRICE
                                    });
                    System.Diagnostics.Debug.WriteLine(query);
                    if (query047.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in query047)
                        {
                            counter++;

                            string carry = "";
                            output047 += lp.MAT_ID + "|" + lp.MAT_NAME + "|" + lp.MAT_UOM + "|" + lp.MAT_CAT + "|" + lp.MAT_PRICE.ToString();

                            output047 += "!";
                            System.Diagnostics.Debug.WriteLine("agak:" + output047);
                        }
                    }
                }

                if (check048 > 0)
                {
                    var query048 = (from a in ctxData.BI_NNW_PU_MAT_REF
                                    where a.BI_BATCH_ID == id
                                    select new
                                    {
                                        a.PU_ID,
                                        a.MAT_ID,
                                        a.PM_QTY
                                    });
                    System.Diagnostics.Debug.WriteLine(query);
                    if (query048.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in query048)
                        {
                            counter++;

                            string carry = "";
                            output048 += lp.PU_ID + "|" + lp.MAT_ID + "|" + lp.PM_QTY;

                            output048 += "!";
                            System.Diagnostics.Debug.WriteLine("agak:" + output048);
                        }
                    }
                }

                if (check049 > 0)
                {
                    var query049 = (from a in ctxData.BI_CONTRACT
                                    where a.BI_BATCH_ID == id
                                    select new
                                    {
                                        a.CONTRACT_NO,
                                        a.ITEM_NO,
                                        a.SHORT_TEXT,
                                        a.ITEM_CATEGORY,
                                        a.ORDER_UNIT,
                                        a.NET_PRICE
                                    });
                    System.Diagnostics.Debug.WriteLine(query);
                    if (query049.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in query049)
                        {
                            counter++;

                            string carry = "";
                            output049 += lp.CONTRACT_NO + "|" + lp.ITEM_NO + "|" + lp.SHORT_TEXT + "|" + lp.ITEM_CATEGORY + "|" + lp.ORDER_UNIT + "|" + lp.NET_PRICE.ToString();

                            output049 += "!";
                            System.Diagnostics.Debug.WriteLine("agak:" + output049);
                        }
                    }
                }

                if (check309 > 0)
                {
                    var query309 = (from a in ctxData.BI_NNW_EXT_SCH_GEM
                                    where a.BI_BATCH_ID == id
                                    select new
                                    {
                                        a.PROJ_NUM,
                                        a.NETWORK_HDR,
                                        a.RATE_INDICATOR,
                                        a.PU_ID,
                                        a.PU_QTY,
                                        a.WBS_NUM
                                    });
                    System.Diagnostics.Debug.WriteLine(query);
                    if (query309.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in query309)
                        {
                            counter++;

                            string carry = "";
                            output309 += lp.PROJ_NUM + "|" + lp.NETWORK_HDR + "|" + lp.RATE_INDICATOR + "|" + lp.PU_ID + "|" + lp.PU_QTY + "|" + lp.WBS_NUM;

                            output309 += "!";
                        }
                    }
                }
                if (check310 > 0)
                {
                    var query310 = (from a in ctxData.BI_NNW_GEM_PROJNO
                                    where a.BI_BATCH_ID == id
                                    select new
                                    {
                                        a.WBS_NUM,
                                        a.PROJ_DESC,
                                        a.WBS_DESC,
                                        a.EXC_ABB,
                                        a.AMOUNT,
                                        a.NETWORK_HDR
                                    });
                    System.Diagnostics.Debug.WriteLine(query310);
                    if (query310.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in query310)
                        {
                            counter++;

                            string carry = "";
                            output310 += lp.WBS_NUM + "|" + lp.PROJ_DESC + "|" + lp.WBS_DESC + "|" + lp.EXC_ABB + "|" + lp.AMOUNT.ToString() + "|" + lp.NETWORK_HDR;

                            output310 += "!";
                            System.Diagnostics.Debug.WriteLine("agak:" + output310);
                        }
                    }
                }

                //output310 = soutput310;

                if (check311 > 0)
                {
                    var query311 = (from a in ctxData.BI_TOGEMS_SEQ1
                                    where a.BI_BATCH_ID == id
                                    select new
                                    {
                                        a.PROJECT_NO,
                                        a.EQUIPMENT_CODE,
                                        a.EXC_ABB,
                                        a.ASSET_ADDRESS1,
                                        a.ASSET_ADDRESS2,
                                        a.ASSET_ADDRESS3,
                                        a.WBS_NO
                                    });
                    System.Diagnostics.Debug.WriteLine(query);
                    if (query311.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in query311)
                        {
                            counter++;
                            string address = lp.ASSET_ADDRESS1 + "," + lp.ASSET_ADDRESS2 + "," + lp.ASSET_ADDRESS3;

                            string carry = "";
                            output311 += lp.PROJECT_NO + "|" + lp.WBS_NO + "|" + lp.EQUIPMENT_CODE + "|" + lp.EXC_ABB + "|" + address;

                            output311 += "!";
                        }
                    }
                }
                if (check312 > 0)
                {
                    var query312 = (from a in ctxData.BI_GEMS_BOQ_MAT
                                    where a.BI_BATCH_ID == id
                                    select new
                                    {
                                        a.PROJ_NO,
                                        a.NETWORK_HEADER,
                                        a.CONTRACT_NO,
                                        a.CONTRACT_ITEM_NO,
                                        a.QUANTITY,
                                        a.WBS_NUM
                                    });
                    System.Diagnostics.Debug.WriteLine(query);
                    if (query312.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in query312)
                        {
                            counter++;

                            string carry = "";
                            output312 += lp.PROJ_NO + "|" + lp.NETWORK_HEADER + "|" + lp.CONTRACT_NO + "|" + lp.CONTRACT_ITEM_NO + "|" + lp.QUANTITY + "|" + lp.WBS_NUM;

                            output312 += "!";
                        }
                    }
                }

                if (check589 > 0)
                {
                    var query589 = (from a in ctxData.BI_ASSET_NO
                                    where a.BI_BATCH_ID == id
                                    select new
                                    {
                                        a.PROJECT_NO,
                                        a.ASSET_NO,
                                        a.ASSET_SUB_NO,
                                        a.CABLE_EQUIP_ID,
                                        a.EXC_ABB,
                                        a.ASSET_TYPE
                                    });
                    System.Diagnostics.Debug.WriteLine(query);
                    if (query589.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in query589)
                        {
                            counter++;

                            string carry = "";
                            output589 += lp.PROJECT_NO.ToString() + "|" + lp.ASSET_NO.ToString() + "|" + lp.ASSET_SUB_NO.ToString() + "|" + lp.CABLE_EQUIP_ID + "|" + lp.EXC_ABB + "|" + lp.ASSET_TYPE;

                            output589 += "!";
                            System.Diagnostics.Debug.WriteLine("agak:" + output591);
                        }
                    }
                }
                if (check591 > 0)
                {
                    var query591 = (from a in ctxData.BI_DECOMM_TABLE
                                    where a.BI_BATCH_ID == id
                                    select new
                                    {
                                        a.ASSET_NO,
                                        a.CABLE_ID,
                                        a.DECOMM_REASON,
                                        a.DECOMM_DATE
                                    });
                    System.Diagnostics.Debug.WriteLine(query);
                    if (query591.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in query591)
                        {
                            counter++;

                            string carry = "";
                            output591 += lp.ASSET_NO.ToString() + "|" + lp.CABLE_ID.ToString() + "|" + lp.DECOMM_REASON + "|" + lp.DECOMM_DATE.ToString();

                            output591 += "!";
                            System.Diagnostics.Debug.WriteLine("agak:" + output591);
                        }
                    }
                }
                
                return Json(new
                {

                    Success = true,
                    BATCH_ID = id,
                    CLASS_NAME = query.CLASS_NAME,
                    EXCEPTION = query.EXCEPTION,
                    EXCEPTION_MSG = query.EXCEPTION_MSG,
                    FILE_HAS_ERROR = query.FILE_HAS_ERROR,
                    FILENAME = query.FILENAME,
                    INSTANCE_ID = query.INSTANCE_ID,
                    SERVICE_NAME = query.SERVICE_NAME,
                    STACKTRACE = query.STACKTRACE,
                    TIME_END = query.TIME_END.ToString(),
                    TIME_START = query.TIME_START.ToString(),
                    TYPE = query.TYPE,
                    outputATag = outputATag,
                    outputATagNis = outputATagNis,
                    output046 = output046,
                    output047 = output047,
                    output048 = output048,
                    output049 = output049,
                    output309 = output309,
                    output310 = output310,
                    output311 = output311,
                    output312 = output312,
                    output589 = output589,
                    output591 = output591 
                }, JsonRequestBehavior.AllowGet); //
            }
        }
    }
}