using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data.EntityClient;
using System.Data.EntityModel;
using Oracle.DataAccess.Client;
using System.Data.Objects;
using System.Data.Common;
using System.Data;
using WebView.Models;
using WebView.Library;
using System.IO;
using PagedList;

namespace WebView.Controllers
{
    public class NetworkElementController : Controller
    {
        //
        // GET: /AND/ 

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult List(string PTT_ID, string EQUP_LOCN_TTNAME, string EQUP_EQUT_ABBREVIATION, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            if (EQUP_EQUT_ABBREVIATION == null)
            {
                EQUP_EQUT_ABBREVIATION = "Select";
            }
            WebService._base.ISPNetworkElement NetworkElementMain = new WebService._base.ISPNetworkElement();
            if (EQUP_LOCN_TTNAME != null)
            {
                if (PTT_ID.Equals("Select") && EQUP_LOCN_TTNAME.Equals("Select") && EQUP_EQUT_ABBREVIATION.Equals("Select"))
                {
                    NetworkElementMain = myWebService.GetISPNetworkElementMaintenance(0, 100000, null, null, null);
                    System.Diagnostics.Debug.WriteLine("A");
                }
                else
                {
                    NetworkElementMain = myWebService.GetISPNetworkElementMaintenance(0, 100000, PTT_ID, EQUP_LOCN_TTNAME, EQUP_EQUT_ABBREVIATION);
                    System.Diagnostics.Debug.WriteLine("b");
                }
            }
            else
            {
                NetworkElementMain = myWebService.GetISPNetworkElementMaintenance(0, 100000, null, null, null);
                System.Diagnostics.Debug.WriteLine("c");
            }

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            using (EntitiesNetworkElement ctxData = new EntitiesNetworkElement())
            {
                //filter PTT
                List<SelectListItem> list = new List<SelectListItem>();
                var queryEXC = from p in ctxData.MIG_EXC_PTT
                               select new { Text = p.PTT, Value = p.PTT };

                list.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in queryEXC.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }
                ViewBag.PTT = list;

                //filter EQUP_LOCN_TTNAME
                List<SelectListItem> list2 = new List<SelectListItem>();
                var queryEXC2 = from p in ctxData.MIG_EXC_PTT
                                where p.PTT == PTT_ID
                                select new { Text = p.EXC, Value = p.EXC };

                list2.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in queryEXC2.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list2.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }
                ViewBag.EQUP_LOCN_TTNAME = list2;
            }

            using (EntitiesNetworkElement ctxData = new EntitiesNetworkElement())
            {
                //filter EQUP_EQUT_ABBREVIATION
                List<SelectListItem> list3 = new List<SelectListItem>();
                var queryEXC3 = from p in ctxData.MIG_NIS_CHASSIS
                                join fx in ctxData.MIG_EXCHANGE on p.MIG_EXCHANGE_ID equals fx.MIG_EXCHANGE_ID
                                //where fx.EXCHANGE.Trim() == EQUP_LOCN_TTNAME.Trim()
                                select new { Text = p.EQUP_EQUT_ABBREVIATION, Value = p.EQUP_EQUT_ABBREVIATION };

                list3.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in queryEXC3.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list3.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }
                ViewBag.EQUP_EQUT_ABBREVIATION = list3;
            }          

            int pageSize = 10;
            int pageNumber = (page ?? 1);
            if (EQUP_LOCN_TTNAME == ""){
                EQUP_LOCN_TTNAME = "Select";
            }
            if (EQUP_LOCN_TTNAME == null)
            {
                EQUP_LOCN_TTNAME = "Select";
            }
            ViewBag.PTT_1 = PTT_ID;
            ViewBag.EQUP_LOCN_TTNAME_1 = EQUP_LOCN_TTNAME;
            ViewBag.EQUP_EQUT_ABBREVIATION_1 = EQUP_EQUT_ABBREVIATION;
            return View(NetworkElementMain.NetworkElementList.ToPagedList(pageNumber, pageSize));
        }

        [HttpPost]
        public ActionResult updataListData(string PTT_ID) // NIS Network Element
        {
            string PuList = "";
            //string PuList2 = "";
            using (EntitiesNetworkElement ctxData = new EntitiesNetworkElement())
            {
                //filter PTT

                List<SelectListItem> list = new List<SelectListItem>();
                var queryEXC = from p in ctxData.MIG_EXC_PTT
                               where p.PTT.Trim() == PTT_ID.Trim()
                               orderby p.EXC
                               select new { p.EXC};

                foreach (var a in queryEXC)
                {
                    PuList = PuList + a.EXC + ":  " + a.EXC + "|";
                }
                ViewBag.EQUP_LOCN_TTNAME = list;
            }

            return Json(new
            {
                Success = true,
                PuList = PuList
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult updataListData2(string PTT_ID) // NIS Network Element
        {
            string PuList = "";
            using (EntitiesNetworkElement ctxData = new EntitiesNetworkElement())
            {
                //filter PTT

                var queryEXC = from p in ctxData.MIG_NIS_CHASSIS
                               join fx in ctxData.MIG_EXCHANGE on p.MIG_EXCHANGE_ID equals fx.MIG_EXCHANGE_ID
                               where fx.EXCHANGE.Trim() == PTT_ID.Trim()
                               select new { p.EQUP_EQUT_ABBREVIATION };

                foreach (var a in queryEXC.Distinct().OrderBy(it => it.EQUP_EQUT_ABBREVIATION))
                {
                    PuList = PuList + a.EQUP_EQUT_ABBREVIATION + ":  " + a.EQUP_EQUT_ABBREVIATION + "|";
                }
                
            }

            return Json(new
            {
                Success = true,
                PuList = PuList
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult updataListData3(string PTT_ID) // Granite Network Element
        {
            string PuList = "";
            using (EntitiesNetworkElement ctxData = new EntitiesNetworkElement())
            {
                //filter PTT

                var queryEXC = from p in ctxData.MIG_GRT_CHASSIS
                               join fx in ctxData.MIG_EXCHANGE on p.MIG_EXCHANGE_ID equals fx.MIG_EXCHANGE_ID
                               where fx.EXCHANGE.Trim() == PTT_ID.Trim()
                               select new { p.EQUP_EQUT_ABBREVIATION };

                foreach (var a in queryEXC.Distinct().OrderBy(it => it.EQUP_EQUT_ABBREVIATION))
                {
                    PuList = PuList + a.EQUP_EQUT_ABBREVIATION + ":  " + a.EQUP_EQUT_ABBREVIATION + "|";
                }

            }

            return Json(new
            {
                Success = true,
                PuList = PuList
            }, JsonRequestBehavior.AllowGet); //
        } 

        [HttpPost]
        public ActionResult updataListData4(string PTT_ID) // NIS Frame Container
        {
            string PuList = "";
            using (EntitiesNetworkElement ctxData = new EntitiesNetworkElement())
            {
                //filter PTT

                var queryEXC = from c in ctxData.MIG_NIS_RACK
                               join c1 in ctxData.MIG_NIS_BAYLINE on c.MIG_NIS_BAYLINE_ID equals c1.MIG_NIS_BAYLINE_ID
                               join c2 in ctxData.MIG_NIS_PPANEL on c.MIG_NIS_RACK_ID equals c2.MIG_NIS_RACK_ID
                               join fx in ctxData.MIG_EXCHANGE on c1.MIG_EXCHANGE_ID equals fx.MIG_EXCHANGE_ID
                               where fx.EXCHANGE.Trim() == PTT_ID.Trim()
                               select new { c2.FUPT_MANR_ABBREVIATION };

                foreach (var a in queryEXC.Distinct().OrderBy(it => it.FUPT_MANR_ABBREVIATION))
                {
                    PuList = PuList + a.FUPT_MANR_ABBREVIATION + ":  " + a.FUPT_MANR_ABBREVIATION + "|";
                }

            }

            return Json(new
            {
                Success = true,
                PuList = PuList
            }, JsonRequestBehavior.AllowGet); //
        }

        public ActionResult ListGraniteNetwork(string PTT_ID, string EQUP_LOCN_TTNAME, string EQUP_EQUT_ABBREVIATION, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            if (EQUP_EQUT_ABBREVIATION == null)
            {
                EQUP_EQUT_ABBREVIATION = "Select";
            }
            WebService._base.ISPNetworkElement NetworkElementMain = new WebService._base.ISPNetworkElement();
            if (PTT_ID != null)
            {
                if (PTT_ID.Equals("Select") && EQUP_LOCN_TTNAME.Equals("Select") && EQUP_EQUT_ABBREVIATION.Equals("Select"))
                {
                    NetworkElementMain = myWebService.GetGraniteNetworkElementMaintenance(0, 100000, null, null, null);
                }
                else
                {
                    NetworkElementMain = myWebService.GetGraniteNetworkElementMaintenance(0, 100000, PTT_ID, EQUP_LOCN_TTNAME, EQUP_EQUT_ABBREVIATION);
                }
            }
            else
            {
                NetworkElementMain = myWebService.GetGraniteNetworkElementMaintenance(0, 100000, null, null, null);
            }

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            using (EntitiesNetworkElement ctxData = new EntitiesNetworkElement())
            {
                //filter PTT
                List<SelectListItem> list = new List<SelectListItem>();
                var queryEXC = from p in ctxData.MIG_EXC_PTT
                               select new { Text = p.PTT, Value = p.PTT };

                list.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in queryEXC.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }
                ViewBag.PTT = list;

                //filter EQUP_LOCN_TTNAME
                List<SelectListItem> list2 = new List<SelectListItem>();
                var queryEXC2 = from p in ctxData.MIG_EXC_PTT
                                where p.PTT == PTT_ID
                                select new { Text = p.EXC, Value = p.EXC };

                list2.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in queryEXC2.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list2.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }
                ViewBag.EQUP_LOCN_TTNAME = list2;
            }

            using (EntitiesNetworkElement ctxData = new EntitiesNetworkElement())
            {
                //filter EQUP_EQUT_ABBREVIATION
                List<SelectListItem> list3 = new List<SelectListItem>();
                var queryEXC3 = from p in ctxData.MIG_GRT_CHASSIS
                                join fx in ctxData.MIG_EXCHANGE on p.MIG_EXCHANGE_ID equals fx.MIG_EXCHANGE_ID
                                //where fx.EXCHANGE.Trim() == EQUP_LOCN_TTNAME.Trim()
                                select new { Text = p.EQUP_EQUT_ABBREVIATION, Value = p.EQUP_EQUT_ABBREVIATION };

                list3.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in queryEXC3.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list3.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }
                ViewBag.EQUP_EQUT_ABBREVIATION = list3;
            }

            int pageSize = 10;
            int pageNumber = (page ?? 1);
            if (EQUP_LOCN_TTNAME == "")
            {
                EQUP_LOCN_TTNAME = "Select";
            }
            if (EQUP_LOCN_TTNAME == null)
            {
                EQUP_LOCN_TTNAME = "Select";
            }
            ViewBag.PTT_1 = PTT_ID;
            ViewBag.EQUP_LOCN_TTNAME_1 = EQUP_LOCN_TTNAME;
            ViewBag.EQUP_EQUT_ABBREVIATION_1 = EQUP_EQUT_ABBREVIATION;
            return View(NetworkElementMain.NetworkElementList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult ListNisFrame(string PTT_ID, string EQUP_LOCN_TTNAME, string EQUP_EQUT_ABBREVIATION, string FRAU_NAME, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            if (EQUP_EQUT_ABBREVIATION == null)
            {
                EQUP_EQUT_ABBREVIATION = "Select";
            }
            if (FRAU_NAME == null)
            {
                FRAU_NAME = "Select";
            }
            WebService._base.ISPNetworkElement NetworkElementMain = new WebService._base.ISPNetworkElement();
            if (PTT_ID != null)
            {
                if (PTT_ID.Equals("Select") && EQUP_LOCN_TTNAME.Equals("Select") && EQUP_EQUT_ABBREVIATION.Equals("Select") && FRAU_NAME.Equals("Select"))
                {
                    NetworkElementMain = myWebService.GetNisFrameMaintenance(0, 100000, null, null, null, null);
                }
                else
                {
                    NetworkElementMain = myWebService.GetNisFrameMaintenance(0, 100000, PTT_ID, EQUP_LOCN_TTNAME, EQUP_EQUT_ABBREVIATION, FRAU_NAME);
                }
            }
            else
            {
                NetworkElementMain = myWebService.GetNisFrameMaintenance(0, 100000, null, null, null, null);
            }

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            using (EntitiesNetworkElement ctxData = new EntitiesNetworkElement())
            {
                //filter PTT
                List<SelectListItem> list = new List<SelectListItem>();
                var queryEXC = from p in ctxData.MIG_EXC_PTT
                               select new { Text = p.PTT, Value = p.PTT };

                list.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in queryEXC.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }
                ViewBag.PTT = list;

                //filter EQUP_LOCN_TTNAME
                List<SelectListItem> list2 = new List<SelectListItem>();
                var queryEXC2 = from p in ctxData.MIG_EXC_PTT
                                where p.PTT == PTT_ID
                                select new { Text = p.EXC, Value = p.EXC };

                list2.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in queryEXC2.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list2.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }
                ViewBag.EQUP_LOCN_TTNAME = list2;
            }

            using (EntitiesNetworkElement ctxData = new EntitiesNetworkElement())
            {
                //filter EQUP_LOCN_TTNAME
                //List<SelectListItem> list2 = new List<SelectListItem>();
                //var queryEXC2 = from p in ctxData.MIG_EXC_PTT
                //                where p.PTT == PTT_ID
                //                select new { Text = p.EXC, Value = p.EXC };

                //list2.Add(new SelectListItem() { Text = "", Value = "Select" });
                //foreach (var a in queryEXC2.Distinct().OrderBy(it => it.Value))
                //{
                //    if (a.Value != null)
                //        list2.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                //}
                //ViewBag.EQUP_LOCN_TTNAME = list2;

                //filter EQUP_EQUT_ABBREVIATION
                List<SelectListItem> list3 = new List<SelectListItem>();
                var queryEXC3 = from c in ctxData.MIG_NIS_RACK
                                join c1 in ctxData.MIG_NIS_BAYLINE on c.MIG_NIS_BAYLINE_ID equals c1.MIG_NIS_BAYLINE_ID
                                join c2 in ctxData.MIG_NIS_PPANEL on c.MIG_NIS_RACK_ID equals c2.MIG_NIS_RACK_ID
                                join fx in ctxData.MIG_EXCHANGE on c1.MIG_EXCHANGE_ID equals fx.MIG_EXCHANGE_ID
                                //where fx.EXCHANGE.Trim() == EQUP_LOCN_TTNAME.Trim()
                                select new { Text = c1.FRAC_FRAN_NAME, Value = c1.FRAC_FRAN_NAME };

                list3.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in queryEXC3.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list3.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }
                ViewBag.EQUP_EQUT_ABBREVIATION = list3;

                //filter EQUP_EQUT_ABBREVIATION
                List<SelectListItem> list4 = new List<SelectListItem>();
                var queryPOS = from c in ctxData.MIG_NIS_RACK
                               select new { Text = c.FRAU_NAME, Value = c.FRAU_NAME };

                list4.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in queryPOS.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list4.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }
                ViewBag.FRAU_NAME = list4;
            }

            int pageSize = 10;
            int pageNumber = (page ?? 1);
            if (EQUP_LOCN_TTNAME == "")
            {
                EQUP_LOCN_TTNAME = "Select";
            }
            if (EQUP_LOCN_TTNAME == null)
            {
                EQUP_LOCN_TTNAME = "Select";
            }
            ViewBag.PTT_1 = PTT_ID;
            ViewBag.EQUP_LOCN_TTNAME_1 = EQUP_LOCN_TTNAME;
            ViewBag.EQUP_EQUT_ABBREVIATION_1 = EQUP_EQUT_ABBREVIATION;
            ViewBag.FRAU_NAME_1 = FRAU_NAME;
            return View(NetworkElementMain.NetworkElementList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult TemplateList(string MAN, string MOD, string FNO, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.NepsAgTemplate AgTemplateMain = new WebService._base.NepsAgTemplate();
            if (MAN != null)
            {
                if (MAN.Equals("Select") && MOD.Equals("Select") && FNO.Equals("Select"))
                {
                    AgTemplateMain = myWebService.GetAgTemplate(0, 100000, null, null, null);
                }
                else
                {
                    AgTemplateMain = myWebService.GetAgTemplate(0, 100000, MAN, MOD, FNO);
                }
            }
            else
            {
                AgTemplateMain = myWebService.GetAgTemplate(0, 100000, null, null, null);
            }

            string input = "\\\\adsvr";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));
            ViewBag.output = output;

            using (Entities_NEPS ctxData = new Entities_NEPS())
            {
                List<SelectListItem> list = new List<SelectListItem>();
                var queryMAN = from p in ctxData.AG_MSAN_XLTEMPLATE
                               select new { Text = p.MANUFACTURER, Value = p.MANUFACTURER };

                list.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in queryMAN.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }
                ViewBag.MANUFACTURER = list;

                //filter EQUP_LOCN_TTNAME
                List<SelectListItem> list2 = new List<SelectListItem>();
                var queryMOD = from p in ctxData.AG_MSAN_XLTEMPLATE
                                select new { Text = p.MODEL, Value = p.MODEL };

                list2.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in queryMOD.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list2.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }
                ViewBag.MODEL = list2;
            
            }
            using (Entities ctxData = new Entities())
            {
             //filter EQUP_EQUT_ABBREVIATION
                List<SelectListItem> list3 = new List<SelectListItem>();
                var queryFNO = from p in ctxData.G3E_FEATURE
                                select new { p.G3E_FNO, p.G3E_USERNAME};

                list3.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in queryFNO.Distinct().OrderBy(it => it.G3E_FNO))
                {
                        list3.Add(new SelectListItem() { Text = a.G3E_FNO + " - " + a.G3E_USERNAME, Value = a.G3E_FNO.ToString() });
                }
                ViewBag.FNO = list3;

            }
            ViewBag.MAN_1 = MAN;
            ViewBag.MOD_1 = MOD;
            ViewBag.FNO_1 = FNO;
            int pageSize = 10;
            int pageNumber = (page ?? 1);
            return View(AgTemplateMain.AgTemplateList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult NewTemplate()
        {
            using (Entities ctxData = new Entities())
            {
                List<SelectListItem> list2 = new List<SelectListItem>();
                ViewBag.MANUFACTURER = list2;//.Distinct();

                List<SelectListItem> list3 = new List<SelectListItem>();
                ViewBag.MODEL = list3;

                List<SelectListItem> list4 = new List<SelectListItem>();
                var query1 = (from p in ctxData.G3E_FEATURE
                              select new { p.G3E_FNO, p.G3E_USERNAME }).Distinct();

                foreach (var a in query1.Distinct().OrderBy(it => it.G3E_FNO))
                {
                    list4.Add(new SelectListItem() { Text = a.G3E_FNO + " - " + a.G3E_USERNAME, Value = a.G3E_FNO.ToString() });
                }
                ViewBag.FNO = list4;
            }
            return View();
        }

        [HttpPost]
        public ActionResult NewTemplate(AgTemplate m)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;

            WebService._base.AgTemplate TempNewjob = new WebService._base.AgTemplate();
            TempNewjob.MANUFACTURER = m.MANUFACTURER;
            TempNewjob.TEMPLATE_MODEL = m.TEMPLATE_MODEL;
            TempNewjob.XL_FILE = m.XL_FILE;
            TempNewjob.MSAN_FNO = m.MSAN_FNO;
            success = myWebService.AddTemplate(TempNewjob);

            selected = true;


            if (ModelState.IsValid && selected)
            {
                if (success == true)
                    return RedirectToAction("TemplateList");
                else
                    return RedirectToAction("NewSaveFail"); // store to db failed.
            }

            return View(m);
        }

        [HttpPost]
        public ActionResult GetDetails(int id)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            string outputcrtNE = "";
            string outputcrtMin = "";
            string xlName = "";
            string fno = "";
            using (Entities_NEPS ctxData = new Entities_NEPS())
            {
                var queryXL = (from a in ctxData.AG_MSAN_XLTEMPLATE
                               where a.MODEL_ID == id
                               select new
                               { a.XL_FILE, a.MSAN_FNO }).Single();
                
                xlName = queryXL.XL_FILE;
                fno = queryXL.MSAN_FNO.ToString();

                var queryNE = (from a in ctxData.AG_MSAN_TEMPLATE
                               join fx in ctxData.AG_MSAN_XLTEMPLATE on a.MODEL_ID equals fx.MODEL_ID
                               where a.MODEL_ID == id
                               select new
                               {
                                   a.ID,
                                   a.RACK_NO,
                                   a.FRAME_NO,
                                   a.SLOT_NO,
                                   a.SLOT_NIS,
                                   a.CARD_TYPE,
                                   a.CARD_MODEL,
                                   a.PORT_LO,
                                   a.PORT_HI,
                                   a.MIN_MATERIAL
                               });
                System.Diagnostics.Debug.WriteLine(queryNE);
                if (queryNE.Count() > 0)
                {
                    int counter = 0;
                    foreach (var lp in queryNE)
                    {
                        counter++;
                        outputcrtNE += lp.RACK_NO + "|" + lp.FRAME_NO + "|" + lp.SLOT_NO + "|" + lp.SLOT_NIS + "|" + lp.CARD_TYPE + 
                                       "|" + lp.CARD_MODEL + "|" + lp.PORT_LO + "|" + lp.PORT_HI + "|" +lp.MIN_MATERIAL + "|" + lp.ID  ;

                        outputcrtNE += "!";
                    }
                }

                if (queryXL.MSAN_FNO == 9100)
                {
                    var queryMin = (from a in ctxData.REF_IPMSAN
                                    select new
                                    {
                                        a.MIN_MATERIAL,
                                        a.MODEL,
                                        a.MANUFACTURER
                                    });
                    foreach (var lp in queryMin)
                    {
                        outputcrtMin += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
                    }
                }
                else if (queryXL.MSAN_FNO == 9600)
                {
                    var queryMin = (from a in ctxData.REF_RT
                                    select new
                                    {
                                        a.MIN_MATERIAL,
                                        a.MODEL,
                                        a.MANUFACTURER
                                    });
                    foreach (var lp in queryMin)
                    {
                        outputcrtMin += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
                    }
                }
                else if (queryXL.MSAN_FNO == 9800)
                {
                    var queryMin = (from a in ctxData.REF_VDSL2
                                    select new
                                    {
                                        a.MIN_MATERIAL,
                                        a.MODEL,
                                        a.MANUFACTURER
                                    });
                    foreach (var lp in queryMin)
                    {
                        outputcrtMin += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
                    }
                }
                else if (queryXL.MSAN_FNO == 5200)
                {
                    var queryMin = (from a in ctxData.REF_EPE
                                    select new
                                    {
                                        a.MIN_MATERIAL,
                                        a.MODEL,
                                        a.MANUFACTURER
                                    });
                    foreach (var lp in queryMin)
                    {
                        outputcrtMin += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
                    }
                }
                else if (queryXL.MSAN_FNO == 5400)
                {
                    var queryMin = (from a in ctxData.REF_UPE
                                    select new
                                    {
                                        a.MIN_MATERIAL,
                                        a.MODEL,
                                        a.MANUFACTURER
                                    });
                    foreach (var lp in queryMin)
                    {
                        outputcrtMin += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
                    }
                }
                else if (queryXL.MSAN_FNO == 9300)
                {
                    var queryMin = (from a in ctxData.REF_DDN
                                    select new
                                    {
                                        a.MIN_MATERIAL,
                                        a.MODEL,
                                        a.MANUFACTURER
                                    });
                    foreach (var lp in queryMin)
                    {
                        outputcrtMin += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
                    }
                }
                else if (queryXL.MSAN_FNO == 9400)
                {
                    var queryMin = (from a in ctxData.REF_NDH
                                    select new
                                    {
                                        a.MIN_MATERIAL,
                                        a.MODEL,
                                        a.MANUFACTURER
                                    });
                    foreach (var lp in queryMin)
                    {
                        outputcrtMin += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
                    }
                }
            }

            using (Entities ctxData = new Entities())
            {
                if (fno == "9500")
                {
                    var queryMin = (from a in ctxData.REF_MUX
                                    select new
                                    {
                                        a.MIN_MATERIAL,
                                        a.MODEL,
                                        a.MANUFACTURER
                                    });
                    foreach (var lp in queryMin)
                    {
                        outputcrtMin += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
                    }
                }
            }

            ViewBag.ID_Temp = id;
            return Json(new
            {
                crtNE = outputcrtNE,
                crtMin = outputcrtMin,
                id = id,
                xlName = xlName
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult GetDetailsUpdate(int id)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            string outputcrtNE = "";
            string outputcrtMin = "";
            //string xlName = "";
            string ID;
            string RACK_NO;
            string FRAME_NO;
            string SLOT_NO;
            string SLOT_NIS;
            string CARD_TYPE = "";
            string CARD_MODEL = "";
            string PORT_LO = "";
            string PORT_HI = "";
            string MIN_MATERIAL = "";
            string ID_MIN = "";
            string fno = "";
            using (Entities_NEPS ctxData = new Entities_NEPS())
            {
                var queryNE = (from a in ctxData.AG_MSAN_TEMPLATE
                               join fx in ctxData.AG_MSAN_XLTEMPLATE on a.MODEL_ID equals fx.MODEL_ID
                               where a.ID == id
                               select new
                               {
                                   a.ID,
                                   a.RACK_NO,
                                   a.FRAME_NO,
                                   a.SLOT_NO,
                                   a.SLOT_NIS,
                                   a.CARD_TYPE,
                                   a.CARD_MODEL,
                                   a.PORT_LO,
                                   a.PORT_HI,
                                   a.MIN_MATERIAL,
                                   fx.MSAN_FNO
                               }).Single();
                System.Diagnostics.Debug.WriteLine(queryNE);
                ID = queryNE.ID.ToString();
                RACK_NO = queryNE.RACK_NO;
                FRAME_NO = queryNE.FRAME_NO;
                SLOT_NO = queryNE.SLOT_NO;
                SLOT_NIS = queryNE.SLOT_NIS;
                CARD_TYPE = queryNE.CARD_TYPE;
                CARD_MODEL = queryNE.CARD_MODEL;
                PORT_LO = queryNE.PORT_LO;
                PORT_HI = queryNE.PORT_HI;
                MIN_MATERIAL = queryNE.MIN_MATERIAL;

                fno = queryNE.MSAN_FNO.ToString();
                if (queryNE.MSAN_FNO == 9100)
                {
                    var queryMin = (from a in ctxData.REF_IPMSAN
                                    select new
                                    {
                                        a.MIN_MATERIAL,
                                        a.MODEL,
                                        a.MANUFACTURER
                                    });
                    foreach (var lp in queryMin)
                    {
                        outputcrtMin += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
                    }
                }
                else if (queryNE.MSAN_FNO == 9600)
                {
                    var queryMin = (from a in ctxData.REF_RT
                                    select new
                                    {
                                        a.MIN_MATERIAL,
                                        a.MODEL,
                                        a.MANUFACTURER
                                    });
                    foreach (var lp in queryMin)
                    {
                        outputcrtMin += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
                    }
                }
                else if (queryNE.MSAN_FNO == 9800)
                {
                    var queryMin = (from a in ctxData.REF_VDSL2
                                    select new
                                    {
                                        a.MIN_MATERIAL,
                                        a.MODEL,
                                        a.MANUFACTURER
                                    });
                    foreach (var lp in queryMin)
                    {
                        outputcrtMin += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
                    }
                }
                else if (queryNE.MSAN_FNO == 5200)
                {
                    var queryMin = (from a in ctxData.REF_EPE
                                    select new
                                    {
                                        a.MIN_MATERIAL,
                                        a.MODEL,
                                        a.MANUFACTURER
                                    });
                    foreach (var lp in queryMin)
                    {
                        outputcrtMin += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
                    }
                }
                else if (queryNE.MSAN_FNO == 5400)
                {
                    var queryMin = (from a in ctxData.REF_UPE
                                    select new
                                    {
                                        a.MIN_MATERIAL,
                                        a.MODEL,
                                        a.MANUFACTURER
                                    });
                    foreach (var lp in queryMin)
                    {
                        outputcrtMin += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
                    }
                }
                else if (queryNE.MSAN_FNO == 9300)
                {
                    var queryMin = (from a in ctxData.REF_DDN
                                    select new
                                    {
                                        a.MIN_MATERIAL,
                                        a.MODEL,
                                        a.MANUFACTURER
                                    });
                    foreach (var lp in queryMin)
                    {
                        outputcrtMin += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
                    }
                }
                else if (queryNE.MSAN_FNO == 9400)
                {
                    var queryMin = (from a in ctxData.REF_NDH
                                    select new
                                    {
                                        a.MIN_MATERIAL,
                                        a.MODEL,
                                        a.MANUFACTURER
                                    });
                    foreach (var lp in queryMin)
                    {
                        outputcrtMin += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
                    }
                }
            }

            using (Entities ctxData = new Entities())
            {
                if (fno == "9500")
                {
                    var queryMin = (from a in ctxData.REF_MUX
                                    select new
                                    {
                                        a.MIN_MATERIAL,
                                        a.MODEL,
                                        a.MANUFACTURER
                                    });
                    foreach (var lp in queryMin)
                    {
                        outputcrtMin += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
                    }
                }
            }

            ViewBag.ID_Temp = id;
            return Json(new
            {
                crtNE = outputcrtNE,
                id = id,
                RACK_NO = RACK_NO,
                FRAME_NO = FRAME_NO,
                SLOT_NO = SLOT_NO,
                SLOT_NIS = SLOT_NIS,
                CARD_TYPE = CARD_TYPE,
                CARD_MODEL = CARD_MODEL,
                PORT_LO = PORT_LO,
                PORT_HI = PORT_HI,
                MIN_MATERIAL = MIN_MATERIAL,
                crtMin = outputcrtMin,
                ID_MIN = ID_MIN
            }, JsonRequestBehavior.AllowGet);
        }

        //public ActionResult GetMinMat(int id)
        //{
        //    WebView.WebService._base myWebService;
        //    myWebService = new WebService._base();

        //    string fno = "";
        //    string MIN_MATERIAL = "";
        //    using (Entities_NEPS ctxData = new Entities_NEPS())
        //    {
        //        var queryXL = (from a in ctxData.AG_MSAN_XLTEMPLATE
        //                       where a.MODEL_ID == id
        //                       select new { a.XL_FILE, a.MSAN_FNO }).Single();

        //        fno = queryXL.MSAN_FNO.ToString();

        //        if (queryXL.MSAN_FNO == 9100)
        //        {
        //            var queryMin = (from a in ctxData.REF_IPMSAN
        //                            select new
        //                            {
        //                                a.MIN_MATERIAL,
        //                                a.MODEL,
        //                                a.MANUFACTURER
        //                            });
        //            foreach (var lp in queryMin)
        //            {
        //                MIN_MATERIAL += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
        //            }
        //        }
        //        else if (queryXL.MSAN_FNO == 9600)
        //        {
        //            var queryMin = (from a in ctxData.REF_RT
        //                            select new
        //                            {
        //                                a.MIN_MATERIAL,
        //                                a.MODEL,
        //                                a.MANUFACTURER
        //                            });
        //            foreach (var lp in queryMin)
        //            {
        //                MIN_MATERIAL += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
        //            }
        //        }
        //        else if (queryXL.MSAN_FNO == 9800)
        //        {
        //            var queryMin = (from a in ctxData.REF_VDSL2
        //                            select new
        //                            {
        //                                a.MIN_MATERIAL,
        //                                a.MODEL,
        //                                a.MANUFACTURER
        //                            });
        //            foreach (var lp in queryMin)
        //            {
        //                MIN_MATERIAL += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
        //            }
        //        }
        //        else if (queryXL.MSAN_FNO == 5200)
        //        {
        //            var queryMin = (from a in ctxData.REF_EPE
        //                            select new
        //                            {
        //                                a.MIN_MATERIAL,
        //                                a.MODEL,
        //                                a.MANUFACTURER
        //                            });
        //            foreach (var lp in queryMin)
        //            {
        //                MIN_MATERIAL += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
        //            }
        //        }
        //        else if (queryXL.MSAN_FNO == 5400)
        //        {
        //            var queryMin = (from a in ctxData.REF_UPE
        //                            select new
        //                            {
        //                                a.MIN_MATERIAL,
        //                                a.MODEL,
        //                                a.MANUFACTURER
        //                            });
        //            foreach (var lp in queryMin)
        //            {
        //                MIN_MATERIAL += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
        //            }
        //        }
        //        else if (queryXL.MSAN_FNO == 9300)
        //        {
        //            var queryMin = (from a in ctxData.REF_DDN
        //                            select new
        //                            {
        //                                a.MIN_MATERIAL,
        //                                a.MODEL,
        //                                a.MANUFACTURER
        //                            });
        //            foreach (var lp in queryMin)
        //            {
        //                MIN_MATERIAL += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
        //            }
        //        }
        //        else if (queryXL.MSAN_FNO == 9400)
        //        {
        //            var queryMin = (from a in ctxData.REF_NDH
        //                            select new
        //                            {
        //                                a.MIN_MATERIAL,
        //                                a.MODEL,
        //                                a.MANUFACTURER
        //                            });
        //            foreach (var lp in queryMin)
        //            {
        //                MIN_MATERIAL += lp.MIN_MATERIAL + "|" + lp.MODEL + "|" + lp.MANUFACTURER + "!";
        //            }
        //        }
        //    }
        //    return Json(new
        //    {
        //        MIN_MATERIAL = MIN_MATERIAL
        //    }, JsonRequestBehavior.AllowGet);
        //}

        public ActionResult AddDetails(string id, string RACK_NO, string FRAME_NO, string SLOT_NO, string SLOT_NIS, string CARD_TYPE, string CARD_MODEL, string PORT_LO, string PORT_HI, string MIN_MATERIAL)
        {
            Tools tool = new Tools();
            bool success = true;
            string id_min = "";

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.AgTemplate tempDet = new WebService._base.AgTemplate();
            System.Diagnostics.Debug.WriteLine("ADD!! :" + RACK_NO + " " + FRAME_NO);
            tempDet.MODEL_ID = id;
            tempDet.RACK_NO = RACK_NO;
            tempDet.FRAME_NO = FRAME_NO;
            tempDet.SLOT_NO = SLOT_NO;
            tempDet.SLOT_NIS = SLOT_NIS;
            tempDet.CARD_TYPE = CARD_TYPE;
            tempDet.CARD_MODEL = CARD_MODEL;
            tempDet.PORT_LO = PORT_LO;
            tempDet.PORT_HI = PORT_HI;
            tempDet.MIN_MATERIAL = MIN_MATERIAL;

            success = myWebService.AddDetailsTemp(tempDet, id);
            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateDetails(string id, string RACK_NO, string FRAME_NO, string SLOT_NO, string SLOT_NIS, string CARD_TYPE, string CARD_MODEL, string PORT_LO, string PORT_HI, string MIN_MATERIAL)
        {
            Tools tool = new Tools();
            bool success = true;
            string id_min = "";

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            //using (Entities_NEPS ctxData = new Entities_NEPS())
            //{
            //    var SingleMin = (from p in ctxData.REF_CARD_MIN_MAT
            //                     where p.ID == MIN_MATERIAL
            //                     select p).Single();

            //    id_min = SingleMin.MIN_MATERIAL;
            //}
            WebService._base.AgTemplate tempDet = new WebService._base.AgTemplate();
            System.Diagnostics.Debug.WriteLine("UPDATE!! :" + RACK_NO);
            tempDet.ID_CHILD = id;
            tempDet.RACK_NO = RACK_NO;
            tempDet.FRAME_NO = FRAME_NO;
            tempDet.SLOT_NO = SLOT_NO;
            tempDet.SLOT_NIS = SLOT_NIS;
            tempDet.CARD_TYPE = CARD_TYPE;
            tempDet.CARD_MODEL = CARD_MODEL;
            tempDet.PORT_LO = PORT_LO;
            tempDet.PORT_HI = PORT_HI;
            tempDet.MIN_MATERIAL = MIN_MATERIAL;

            success = myWebService.UpdateDetailsTemp(tempDet, id);
            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult CopyDetails(string id, string xlfile, string MANUFACTURER, string MODEL, string FNO)
        {
            Tools tool = new Tools();
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.CopyDetailsTemp(id, xlfile, MANUFACTURER, MODEL, FNO);
            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult GetDetailsDelete(int id)
        {
            Tools tool = new Tools();
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.DeleteDetailsTemp(id);
            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult updataListManufacturer(string MANUFACTURER) // NIS Network Element
        {
            string ModList = "";
            using (Entities_NEPS ctxData = new Entities_NEPS())
            {
                if (MANUFACTURER == "IPMSAN")
                {
                    List<SelectListItem> list = new List<SelectListItem>();
                    var queryMan = (from p in ctxData.REF_IPMSAN orderby p.MANUFACTURER select new { p.MANUFACTURER });

                    foreach (var a in queryMan.Distinct().OrderBy(x => x.MANUFACTURER)) {
                        ModList = ModList + a.MANUFACTURER + ":  " + a.MANUFACTURER + "|"; }
                }
                else if (MANUFACTURER == "RT")
                {
                    List<SelectListItem> list = new List<SelectListItem>();
                    var queryMan = (from p in ctxData.REF_RT orderby p.MANUFACTURER select new { p.MANUFACTURER });

                    foreach (var a in queryMan.Distinct().OrderBy(x => x.MANUFACTURER))
                    {
                        ModList = ModList + a.MANUFACTURER + ":  " + a.MANUFACTURER + "|";
                    }
                }
                else if (MANUFACTURER == "VDSL2")
                {
                    List<SelectListItem> list = new List<SelectListItem>();
                    var queryMan = (from p in ctxData.REF_VDSL2 orderby p.MANUFACTURER select new { p.MANUFACTURER });

                    foreach (var a in queryMan.Distinct().OrderBy(x => x.MANUFACTURER))
                    {
                        ModList = ModList + a.MANUFACTURER + ":  " + a.MANUFACTURER + "|";
                    }
                }
                else if (MANUFACTURER == "EPE")
                {
                    List<SelectListItem> list = new List<SelectListItem>();
                    var queryMan = (from p in ctxData.REF_EPE orderby p.MANUFACTURER select new { p.MANUFACTURER });

                    foreach (var a in queryMan.Distinct().OrderBy(x => x.MANUFACTURER))
                    {
                        ModList = ModList + a.MANUFACTURER + ":  " + a.MANUFACTURER + "|";
                    }
                }
                else if (MANUFACTURER == "UPE")
                {
                    List<SelectListItem> list = new List<SelectListItem>();
                    var queryMan = (from p in ctxData.REF_UPE orderby p.MANUFACTURER select new { p.MANUFACTURER });

                    foreach (var a in queryMan.Distinct().OrderBy(x => x.MANUFACTURER))
                    {
                        ModList = ModList + a.MANUFACTURER + ":  " + a.MANUFACTURER + "|";
                    }
                }
                else if (MANUFACTURER == "DDN")
                {
                    List<SelectListItem> list = new List<SelectListItem>();
                    var queryMan = (from p in ctxData.REF_DDN orderby p.MANUFACTURER select new { p.MANUFACTURER });

                    foreach (var a in queryMan.Distinct().OrderBy(x => x.MANUFACTURER))
                    {
                        ModList = ModList + a.MANUFACTURER + ":  " + a.MANUFACTURER + "|";
                    }
                }
                else if (MANUFACTURER == "NDH")
                {
                    List<SelectListItem> list = new List<SelectListItem>();
                    var queryMan = (from p in ctxData.REF_NDH orderby p.MANUFACTURER select new { p.MANUFACTURER });

                    foreach (var a in queryMan.Distinct().OrderBy(x => x.MANUFACTURER))
                    {
                        ModList = ModList + a.MANUFACTURER + ":  " + a.MANUFACTURER + "|";
                    }
                }
            }
            using (Entities ctxData = new Entities())
            {
                if (MANUFACTURER == "MUX")
                {
                    List<SelectListItem> list = new List<SelectListItem>();
                    var queryMan = (from p in ctxData.REF_MUX  orderby p.MANUFACTURER select new { p.MANUFACTURER });

                    foreach (var a in queryMan.Distinct().OrderBy(x => x.MANUFACTURER))
                    {
                        ModList = ModList + a.MANUFACTURER + ":  " + a.MANUFACTURER + "|";
                    }
                }
            }

            return Json(new
            {
                Success = true,
                ModList = ModList
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult updataListModel(string MANUFACTURER, string TYPE) // NIS Network Element
        {
            string ModList = "";
            using (Entities_NEPS ctxData = new Entities_NEPS())
            {
                if (TYPE == "IPMSAN")
                {
                    List<SelectListItem> list = new List<SelectListItem>();
                    var queryMan = (from p in ctxData.REF_IPMSAN where p.MANUFACTURER == MANUFACTURER orderby p.MODEL select new { p.MODEL });

                    foreach (var a in queryMan.Distinct().OrderBy(x => x.MODEL))
                    {
                        ModList = ModList + a.MODEL + ":  " + a.MODEL + "|";
                    }
                }
                else if (TYPE == "RT")
                {
                    List<SelectListItem> list = new List<SelectListItem>();
                    var queryMan = (from p in ctxData.REF_RT where p.MANUFACTURER == MANUFACTURER orderby p.MODEL select new { p.MODEL });

                    foreach (var a in queryMan.Distinct().OrderBy(x => x.MODEL))
                    {
                        ModList = ModList + a.MODEL + ":  " + a.MODEL + "|";
                    }
                }
                else if (TYPE == "VDSL2")
                {
                    List<SelectListItem> list = new List<SelectListItem>();
                    var queryMan = (from p in ctxData.REF_VDSL2 where p.MANUFACTURER == MANUFACTURER orderby p.MODEL select new { p.MODEL });

                    foreach (var a in queryMan.Distinct().OrderBy(x => x.MODEL))
                    {
                        ModList = ModList + a.MODEL + ":  " + a.MODEL + "|";
                    }
                }
                else if (TYPE == "EPE")
                {
                    List<SelectListItem> list = new List<SelectListItem>();
                    var queryMan = (from p in ctxData.REF_EPE where p.MANUFACTURER == MANUFACTURER orderby p.MODEL select new { p.MODEL });

                    foreach (var a in queryMan.Distinct().OrderBy(x => x.MODEL))
                    {
                        ModList = ModList + a.MODEL + ":  " + a.MODEL + "|";
                    }
                }
                else if (TYPE == "UPE")
                {
                    List<SelectListItem> list = new List<SelectListItem>();
                    var queryMan = (from p in ctxData.REF_UPE where p.MANUFACTURER == MANUFACTURER orderby p.MODEL select new { p.MODEL });

                    foreach (var a in queryMan.Distinct().OrderBy(x => x.MODEL))
                    {
                        ModList = ModList + a.MODEL + ":  " + a.MODEL + "|";
                    }
                }
                else if (TYPE == "DDN")
                {
                    List<SelectListItem> list = new List<SelectListItem>();
                    var queryMan = (from p in ctxData.REF_DDN where p.MANUFACTURER == MANUFACTURER orderby p.MODEL select new { p.MODEL });

                    foreach (var a in queryMan.Distinct().OrderBy(x => x.MODEL))
                    {
                        ModList = ModList + a.MODEL + ":  " + a.MODEL + "|";
                    }
                }
                else if (TYPE == "NDH")
                {
                    List<SelectListItem> list = new List<SelectListItem>();
                    var queryMan = (from p in ctxData.REF_NDH where p.MANUFACTURER == MANUFACTURER orderby p.MODEL select new { p.MODEL });

                    foreach (var a in queryMan.Distinct().OrderBy(x => x.MODEL))
                    {
                        ModList = ModList + a.MODEL + ":  " + a.MODEL + "|";
                    }
                }
            }

            using (Entities ctxData = new Entities())
            {
                if (TYPE == "MUX")
                {
                    List<SelectListItem> list = new List<SelectListItem>();
                    var queryMan = (from p in ctxData.REF_MUX where p.MANUFACTURER == MANUFACTURER orderby p.MODEL select new { p.MODEL }).Distinct();

                    foreach (var a in queryMan.Distinct().OrderBy(x => x.MODEL))
                    {
                        ModList = ModList + a.MODEL + ":  " + a.MODEL + "|";
                    }
                }
            }
            System.Diagnostics.Debug.WriteLine("MANUFACTURER : " + MANUFACTURER);
            System.Diagnostics.Debug.WriteLine("MODEL : " + ModList);
            return Json(new
            {
                Success = true,
                ModList = ModList
            }, JsonRequestBehavior.AllowGet); //
        }

        //[HttpPost]
        //public ActionResult updataListMinMat(string MANUFACTURER) // NIS Network Element
        //{
        //    string ModList = "";
        //    using (Entities_NEPS ctxData = new Entities_NEPS())
        //    {
        //        if (MANUFACTURER == "IPMSAN")
        //        {
        //            List<SelectListItem> list = new List<SelectListItem>();
        //            var queryMan = (from p in ctxData.REF_IPMSAN where p.MODEL == MANUFACTURER orderby p.MIN_MATERIAL select new { p.MIN_MATERIAL });

        //            foreach (var a in queryMan.Distinct().OrderBy(x => x.MIN_MATERIAL))
        //            {
        //                ModList = ModList + a.MIN_MATERIAL + ":  " + a.MIN_MATERIAL + "|";
        //            }
        //        }
        //        else if (MANUFACTURER == "RT")
        //        {
        //            List<SelectListItem> list = new List<SelectListItem>();
        //            var queryMan = (from p in ctxData.REF_RT where p.MODEL == MANUFACTURER orderby p.MIN_MATERIAL select new { p.MIN_MATERIAL });

        //            foreach (var a in queryMan.Distinct().OrderBy(x => x.MIN_MATERIAL))
        //            {
        //                ModList = ModList + a.MIN_MATERIAL + ":  " + a.MIN_MATERIAL + "|";
        //            }
        //        }
        //        else if (MANUFACTURER == "VDSL2")
        //        {
        //            List<SelectListItem> list = new List<SelectListItem>();
        //            var queryMan = (from p in ctxData.REF_VDSL2 where p.MODEL == MANUFACTURER orderby p.MIN_MATERIAL select new { p.MIN_MATERIAL });

        //            foreach (var a in queryMan.Distinct().OrderBy(x => x.MIN_MATERIAL))
        //            {
        //                ModList = ModList + a.MIN_MATERIAL + ":  " + a.MIN_MATERIAL + "|";
        //            }
        //        }
        //        else if (MANUFACTURER == "EPE")
        //        {
        //            List<SelectListItem> list = new List<SelectListItem>();
        //            var queryMan = (from p in ctxData.REF_EPE where p.MODEL == MANUFACTURER orderby p.MIN_MATERIAL select new { p.MIN_MATERIAL });

        //            foreach (var a in queryMan.Distinct().OrderBy(x => x.MIN_MATERIAL))
        //            {
        //                ModList = ModList + a.MIN_MATERIAL + ":  " + a.MIN_MATERIAL + "|";
        //            }
        //        }
        //        else if (MANUFACTURER == "UPE")
        //        {
        //            List<SelectListItem> list = new List<SelectListItem>();
        //            var queryMan = (from p in ctxData.REF_UPE where p.MODEL == MANUFACTURER orderby p.MIN_MATERIAL select new { p.MIN_MATERIAL });

        //            foreach (var a in queryMan.Distinct().OrderBy(x => x.MIN_MATERIAL))
        //            {
        //                ModList = ModList + a.MIN_MATERIAL + ":  " + a.MIN_MATERIAL + "|";
        //            }
        //        }
        //        else if (MANUFACTURER == "DDN")
        //        {
        //            List<SelectListItem> list = new List<SelectListItem>();
        //            var queryMan = (from p in ctxData.REF_DDN where p.MODEL == MANUFACTURER orderby p.MIN_MATERIAL select new { p.MIN_MATERIAL });

        //            foreach (var a in queryMan.Distinct().OrderBy(x => x.MIN_MATERIAL))
        //            {
        //                ModList = ModList + a.MIN_MATERIAL + ":  " + a.MIN_MATERIAL + "|";
        //            }
        //        }
        //        else if (MANUFACTURER == "NDH")
        //        {
        //            List<SelectListItem> list = new List<SelectListItem>();
        //            var queryMan = (from p in ctxData.REF_NDH where p.MODEL == MANUFACTURER orderby p.MIN_MATERIAL select new { p.MIN_MATERIAL });

        //            foreach (var a in queryMan.Distinct().OrderBy(x => x.MIN_MATERIAL))
        //            {
        //                ModList = ModList + a.MIN_MATERIAL + ":  " + a.MIN_MATERIAL + "|";
        //            }
        //        }
        //    }

        //    using (Entities ctxData = new Entities())
        //    {
        //        if (MANUFACTURER == "MUX")
        //        {
        //            List<SelectListItem> list = new List<SelectListItem>();
        //            var queryMan = (from p in ctxData.REF_MUX where p.MODEL == MANUFACTURER orderby p.MIN_MATERIAL select new { p.MIN_MATERIAL });

        //            foreach (var a in queryMan.Distinct().OrderBy(x => x.MIN_MATERIAL))
        //            {
        //                ModList = ModList + a.MIN_MATERIAL + ":  " + a.MIN_MATERIAL + "|";
        //            }
        //        }
        //    }

        //    return Json(new
        //    {
        //        Success = true,
        //        ModList = ModList
        //    }, JsonRequestBehavior.AllowGet); //
        //}

        [HttpPost]
        public ActionResult NewEquipment(string equipment, string minmat, string desc, string model, string manufacturer, string capasity, string e1, string protection, string numcore)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;
            minmat = minmat.ToUpper();
            int cap = 0;
            int core3 = 0;
            if (capasity == ""){
                cap = 0;
            } 
            if (numcore == ""){
                core3 = 0;
            }

            System.Diagnostics.Debug.WriteLine("cap :" + cap);

            if (equipment != "MUX")
            {
                using (Entities_NEPS ctxData = new Entities_NEPS())
                {
                    if (equipment == "IPMSAN")
                    {
                        var queryIPMSAN = (from a in ctxData.REF_IPMSAN
                                           where a.MIN_MATERIAL.Trim() == minmat.Trim() && a.MODEL == model && a.MANUFACTURER == manufacturer && a.SYSTEM_CAPACITY == cap
                                           select a).Count();
                        if (queryIPMSAN == 0)
                        {
                            success = myWebService.AddEquipmentTemp(equipment, minmat, desc, model, manufacturer, capasity, e1, protection, numcore);
                        }
                        else
                        {
                            success = false;
                        }
                    }
                    else if (equipment == "RT")
                    {
                        var queryRT = (from a in ctxData.REF_RT
                                       where a.MIN_MATERIAL.Trim() == minmat.Trim() && a.MODEL == model && a.MANUFACTURER == manufacturer && a.SYSTEM_CAPACITY == cap
                                       select a).Count();
                        System.Diagnostics.Debug.WriteLine("queryRT :" + queryRT);
                        if (queryRT == 0)
                        {
                            success = myWebService.AddEquipmentTemp(equipment, minmat, desc, model, manufacturer, capasity, e1, protection, numcore);
                        }
                        else
                        {
                            success = false;
                        }
                    }
                    else if (equipment == "VDSL2")
                    {
                        var queryVDSL2 = (from a in ctxData.REF_VDSL2
                                          where a.MIN_MATERIAL.Trim() == minmat.Trim() && a.MODEL == model && a.MANUFACTURER == manufacturer && a.SYSTEM_CAPACITY == cap
                                          select a).Count();
                        if (queryVDSL2 == 0)
                        {
                            success = myWebService.AddEquipmentTemp(equipment, minmat, desc, model, manufacturer, capasity, e1, protection, numcore);
                        }
                        else
                        {
                            success = false;
                        }
                    }
                    else if (equipment == "EPE")
                    {
                        var queryEPE = (from a in ctxData.REF_EPE
                                        where a.MIN_MATERIAL.Trim() == minmat.Trim() && a.MODEL == model && a.MANUFACTURER == manufacturer && a.SYSTEM_CAPACITY == cap
                                        select a).Count();
                        if (queryEPE == 0)
                        {
                            success = myWebService.AddEquipmentTemp(equipment, minmat, desc, model, manufacturer, capasity, e1, protection, numcore);
                        }
                        else
                        {
                            success = false;
                        }
                    }
                    else if (equipment == "UPE")
                    {
                        var queryUPE = (from a in ctxData.REF_UPE
                                        where a.MIN_MATERIAL.Trim() == minmat.Trim() && a.MODEL == model && a.MANUFACTURER == manufacturer && a.SYSTEM_CAPACITY == cap
                                        select a).Count();
                        if (queryUPE == 0)
                        {
                            success = myWebService.AddEquipmentTemp(equipment, minmat, desc, model, manufacturer, capasity, e1, protection, numcore);
                        }
                        else
                        {
                            success = false;
                        }
                    }
                    else if (equipment == "DDN")
                    {
                        var queryDDN = (from a in ctxData.REF_DDN
                                        where a.MIN_MATERIAL.Trim() == minmat.Trim() && a.MODEL == model && a.MANUFACTURER == manufacturer && a.SYSTEM_CAPACITY == cap
                                        select a).Count();
                        if (queryDDN == 0)
                        {
                            success = myWebService.AddEquipmentTemp(equipment, minmat, desc, model, manufacturer, capasity, e1, protection, numcore);
                        }
                        else
                        {
                            success = false;
                        }
                    }
                    else if (equipment == "NDH")
                    {
                        var queryNDH = (from a in ctxData.REF_NDH
                                        where a.MIN_MATERIAL.Trim() == minmat.Trim() && a.MODEL == model && a.MANUFACTURER == manufacturer && a.SYSTEM_CAPACITY == cap
                                        select a).Count();
                        if (queryNDH == 0)
                        {
                            success = myWebService.AddEquipmentTemp(equipment, minmat, desc, model, manufacturer, capasity, e1, protection, numcore);
                        }
                        else
                        {
                            success = false;
                        }
                    }
                }
            }
            else
            {
                using (Entities ctxData = new Entities())
                {
                    var queryMUX = (from a in ctxData.REF_MUX
                                    where a.MIN_MATERIAL.Trim() == minmat.Trim() && a.MODEL == model && a.MANUFACTURER == manufacturer && a.NO_E1 == e1 && a.PROTECTION == protection && a.NUM_CORE == core3
                                    select a).Count();
                    if (queryMUX == 0)
                    {
                        success = myWebService.AddEquipmentTemp(equipment, minmat, desc, model, manufacturer, capasity, e1, protection, numcore);
                    }
                    else
                    {
                        success = false;
                    }
                }
            }

            selected = true;

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }
    }
}
