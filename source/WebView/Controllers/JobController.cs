using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Security;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Web;
using System.Web.Mvc;
using PagedList;
using WebView.Library;
using WebView.Models;
using WebView.NRMWebService;
using WebView.OSPHandoverServices;
using WebView.ISPHandoverServices;
using WebView.GraniteISP;
using WebView.GraniteOSP;
using System.Configuration;
using WebView.NepsLoadSiteService;
using WebView.NEPS_NIS;
using System.Xml;
using Oracle.DataAccess.Client;
using System.Data;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;
namespace WebView.Controllers
{

    public class JobController : Controller
    {
        //
        // GET: /Job/

        private string connString = ConfigurationManager.AppSettings.Get("connString");
        // region Mubin - CR74-20180330
        private string connString2 = ConfigurationManager.AppSettings.Get("connString2");
        // endRegion
        // region Fatihin - NEPSFLASH
        // private string connStringFlash = ConfigurationManager.AppSettings.Get("connStringFlash");
        // endRegion
        private string visionael = ConfigurationManager.AppSettings.Get("VISIONAEL_URL");

        public ActionResult ViewFeatures(string scheme_name)
        {
            using (Entities ctxData = new Entities())
            {
                // dr gc_neteelm get FNO from g3e_feature
                List<G3E_FEATURE> dataG3EFeature = new List<G3E_FEATURE>();
                var dataQuery = (from fxx in ctxData.GC_NETELEM
                                 join fx in ctxData.G3E_FEATURE on fxx.G3E_FNO equals fx.G3E_FNO
                                 where fxx.SCHEME_NAME == scheme_name
                                 group fx.G3E_USERNAME by fx.G3E_USERNAME into g
                                 select new { G3E_USERNAME = g.Key, Bilangan = g.Count() }).ToList();
                foreach (var a in dataQuery)
                {

                    G3E_FEATURE singleData = new G3E_FEATURE();
                    singleData.G3E_NAME = a.Bilangan.ToString();
                    singleData.G3E_USERNAME = a.G3E_USERNAME;

                    dataG3EFeature.Add(singleData);
                }

                return View(dataG3EFeature);
            }
        }

        public ActionResult Index()
        {
            return RedirectToAction("List", "Job");
        }

        public ActionResult List(string searchKey, int? page, string jobExc, string jobScheme, string jobYear, string jobState)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            WebService._base.OSPJob jobs = new WebService._base.OSPJob();

            if (searchKey != null)
            {
                System.Diagnostics.Debug.WriteLine("A");
                if (searchKey.Equals("") && jobExc.Equals("Select") && jobScheme.Equals("Select") && jobYear.Equals("0") && jobState.Equals("Select"))
                {
                    jobs = myWebService.GetOSPJob(User.Identity.Name, 0, 1000000, null, null, null, null, null);
                    ViewBag.searchKey = searchKey;
                    ViewBag.excabb2 = jobExc;
                    ViewBag.schemeType2 = jobScheme;
                    ViewBag.jobyear2 = jobYear;
                    ViewBag.jobState2 = jobState;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("C");
                    jobs = myWebService.GetOSPJob(User.Identity.Name, 0, 1000000, searchKey, jobExc, jobScheme, jobYear, jobState);
                    if (searchKey != "")
                    {
                        ViewBag.searchKey = searchKey;
                    }
                    else
                    {
                        ViewBag.searchKey = "";
                    }
                    ViewBag.excabb2 = jobExc;
                    ViewBag.schemeType2 = jobScheme;
                    ViewBag.jobyear2 = jobYear;
                    ViewBag.jobState2 = jobState;
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("D");
                jobs = myWebService.GetOSPJob(User.Identity.Name, 0, 1000000, null, null, null, null, null);

                ViewBag.searchKey2 = null;
                ViewBag.excabb2 = "Select";
                ViewBag.schemeType2 = "Select";
                ViewBag.jobState2 = "Select";
                ViewBag.jobyear2 = 0;
                ViewBag.jobState2 = "Select";
            }

            string input = "\\\\adsvr";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            using (Entities ctxData = new Entities())
            {
                string UserGrp = "";

                var queryUser = (from d in ctxData.WV_USER
                                 where d.USERNAME.ToUpper() == User.Identity.Name.ToUpper() || d.USERNAME.ToLower() == User.Identity.Name.ToLower()
                                 select new { d.PTT_STATE, d.EXC, d.GROUPID });

                foreach (var a in queryUser)
                {
                    UserGrp = a.GROUPID;
                }

                var query = from p in ctxData.REF_FEATURE_STATE
                            select new { Text = p.FEATURE_STATE, Value = p.FEATURE_STATE };

                string jobStateList = "";
                foreach (var a in query)
                {
                    jobStateList = jobStateList + a.Text + "|";
                }

                ViewBag.jobstate = jobStateList.Substring(0, jobStateList.Length - 1);


                if (UserGrp == "12")
                {
                    //filter exchange
                    List<SelectListItem> list = new List<SelectListItem>();
                    var queryEXC = from p in ctxData.G3E_JOB
                                   where p.WEBVIEW == 1
                                   select new { Text = p.EXC_ABB, Value = p.EXC_ABB };

                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in queryEXC.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                    }

                    //filter scheme
                    List<SelectListItem> list2 = new List<SelectListItem>();
                    var querySCHEME = from p in ctxData.G3E_JOB
                                      where p.WEBVIEW == 1
                                      select new { Text = p.JOB_TYPE, Value = p.JOB_TYPE };

                    list2.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in querySCHEME.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list2.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                    }

                    //filter year
                    List<SelectListItem> list3 = new List<SelectListItem>();
                    var queryYEAR = from p in ctxData.G3E_JOB
                                    where p.WEBVIEW == 1
                                    select new { Text = p.YEAR, Value = p.YEAR };

                    list3.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in queryYEAR.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list3.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                    }
                    List<SelectListItem> list4 = new List<SelectListItem>();
                    var queryState = from p in ctxData.G3E_JOB
                                     where p.WEBVIEW == 1
                                     select new { Text = p.JOB_STATE, Value = p.JOB_STATE };

                    list4.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in queryState.Distinct().OrderBy(it => it.Value))
                    {

                        if (a.Value != null)
                            list4.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                    }
                    ViewBag.jobState = list4;
                    ViewBag.excabb = list;
                    ViewBag.schemeType = list2;
                    ViewBag.jobyear = list3;

                }
                else
                {
                    //filter exchange
                    List<SelectListItem> list = new List<SelectListItem>();
                    var queryEXC = from p in ctxData.G3E_JOB
                                   where p.G3E_OWNER.ToUpper() == User.Identity.Name.ToUpper() || p.G3E_OWNER.ToLower() == User.Identity.Name.ToLower()
                                   select new { Text = p.EXC_ABB, Value = p.EXC_ABB };

                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in queryEXC.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                    }

                    //filter scheme
                    List<SelectListItem> list2 = new List<SelectListItem>();
                    var querySCHEME = from p in ctxData.G3E_JOB
                                      where p.G3E_OWNER.ToUpper() == User.Identity.Name.ToUpper() || p.G3E_OWNER.ToLower() == User.Identity.Name.ToLower()
                                      select new { Text = p.JOB_TYPE, Value = p.JOB_TYPE };

                    list2.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in querySCHEME.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list2.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                    }

                    //filter year
                    List<SelectListItem> list3 = new List<SelectListItem>();
                    var queryYEAR = from p in ctxData.G3E_JOB
                                    where p.G3E_OWNER.ToUpper() == User.Identity.Name.ToUpper() || p.G3E_OWNER.ToLower() == User.Identity.Name.ToLower()
                                    select new { Text = p.YEAR, Value = p.YEAR };

                    list3.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in queryYEAR.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list3.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                    }
                    List<SelectListItem> list4 = new List<SelectListItem>();
                    var queryState = from p in ctxData.G3E_JOB
                                     where p.G3E_OWNER.ToUpper() == User.Identity.Name.ToUpper() || p.G3E_OWNER.ToLower() == User.Identity.Name.ToLower()
                                     select new { Text = p.JOB_STATE, Value = p.JOB_STATE };

                    list4.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in queryState.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list4.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                    }
                    ViewBag.jobState = list4;
                    ViewBag.excabb = list;
                    ViewBag.schemeType = list2;
                    ViewBag.jobyear = list3;
                }

                #region Mubin CR59/CR60 29082018
                List<SelectListItem> list5 = new List<SelectListItem>();
                var query5 = from p in ctxData.WV_PTT_MAST
                             select new { p.PTT_ID };
                list5.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query5.OrderBy(it => it.PTT_ID))
                {
                    list5.Add(new SelectListItem() { Text = a.PTT_ID, Value = a.PTT_ID });
                }
                ViewBag.Ptt = list5;

                List<SelectListItem> list6 = new List<SelectListItem>();
                //var query6 = from p in ctxData.WV_EXC_MAST
                //             select new { p.EXC_ABB};
                list6.Add(new SelectListItem() { Text = "", Value = "" });
                //foreach (var a in query6.OrderBy(it => it.EXC_ABB))
                //{
                //    list6.Add(new SelectListItem() { Text = a.EXC_ABB, Value = a.EXC_ABB });
                //}
                ViewBag.ExcForQnd = list6;

                List<SelectListItem> list8 = new List<SelectListItem>();
                list8.Add(new SelectListItem() { Text = "", Value = "" });
                ViewBag.qndequipmentid = list8;
            }

            using (Entities9 ctxData2 = new Entities9())
            {
                List<SelectListItem> list7 = new List<SelectListItem>();
                var query7 = from p in ctxData2.REF_GRNQND
                             select new { p.DESCRIPTION };
                list7.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query7.OrderBy(it => it.DESCRIPTION))
                {
                    list7.Add(new SelectListItem() { Text = a.DESCRIPTION, Value = a.DESCRIPTION });
                }
                ViewBag.FeatureType = list7;
            }
                #endregion

            //int page = 1;
            int pageSize = 10;
            int pageNumber = (page ?? 1);
            return View(jobs.JobList.ToPagedList(pageNumber, pageSize));
        }

        [HttpPost]
        public ActionResult updatePttQND(string ptt) // Update list port
        {
            string listData = "";
            using (Entities ctxData = new Entities())
            {
                var query6 = from p in ctxData.WV_EXC_MAST
                             where p.PTT_ID == ptt
                             select new { p.EXC_ABB };
                foreach (var a in query6.Distinct().OrderBy(it => it.EXC_ABB))
                {
                    listData = listData + a.EXC_ABB + "|";
                }
            }
            return Json(new
            {
                Success = true,
                listData = listData
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult updateEquipQND(string exc, string type) // Update list port
        {
            string listData = "";
            using (Entities9 ctxData2 = new Entities9())
            {
                if (type.Contains("Granite QND FDC"))
                {
                    var query6 = from p in ctxData2.BI_GRNQND_FDC
                                 where p.EXC_ABB == exc
                                 select new { p.FDC_CODE };
                    foreach (var a in query6.Distinct().OrderBy(it => it.FDC_CODE))
                    {
                        listData = listData + a.FDC_CODE + "|";
                    }
                }
                else if (type.Contains("Granite QND MSAN"))
                {
                    var query6 = from p in ctxData2.BI_GRNQND_MSAN
                                 where p.EXC_ABB == exc
                                 select new { p.MSAN_CODE };
                    foreach (var a in query6.Distinct().OrderBy(it => it.MSAN_CODE))
                    {
                        listData = listData + a.MSAN_CODE + "|";
                    }
                }
                else if (type.Contains("Granite QND ODP"))
                {
                    var query6 = from p in ctxData2.BI_GRNQND_ODP
                                 where p.EXC_ABB == exc
                                 select new { p.ODP_CODE };
                    foreach (var a in query6.Distinct().OrderBy(it => it.ODP_CODE))
                    {
                        listData = listData + a.ODP_CODE + "|";
                    }
                }
                else
                {
                    var query6 = from p in ctxData2.BI_GRNQND_OLT
                                 where p.EXC_ABB == exc
                                 select new { p.OLT_CODE };
                    foreach (var a in query6.Distinct().OrderBy(it => it.OLT_CODE))
                    {
                        listData = listData + a.OLT_CODE + "|";
                    }
                }

            }
            return Json(new
            {
                Success = true,
                listData = listData
            }, JsonRequestBehavior.AllowGet); //
        }

        public ActionResult ISPList(string searchKey, int? page, string jobExc, string jobScheme, string jobYear, string jobState)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            WebService._base.OSPJob jobs = new WebService._base.OSPJob();

            if (searchKey != null)
            {
                System.Diagnostics.Debug.WriteLine("A");
                if (searchKey.Equals("") && jobExc.Equals("Select") && jobScheme.Equals("Select") && jobYear.Equals("0") && jobState.Equals("Select"))
                {
                    System.Diagnostics.Debug.WriteLine("B");
                    jobs = myWebService.GetISPJob(User.Identity.Name, 0, 1000000, null, null, null, null, null);
                    ViewBag.searchKey = searchKey;
                    ViewBag.excabb2 = jobExc;
                    ViewBag.schemeType2 = jobScheme;
                    ViewBag.jobyear2 = jobYear;
                    ViewBag.jobState2 = jobState;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("C");
                    jobs = myWebService.GetISPJob(User.Identity.Name, 0, 1000000, searchKey, jobExc, jobScheme, jobYear, jobState);
                    if (searchKey != "")
                    {
                        ViewBag.searchKey = searchKey;
                    }
                    else
                    {
                        ViewBag.searchKey = "";
                    }
                    System.Diagnostics.Debug.WriteLine("  " + searchKey + "   " + jobExc + "   " + jobScheme + "   " + jobYear + "   " + jobState);
                    ViewBag.excabb2 = jobExc;
                    ViewBag.schemeType2 = jobScheme;
                    ViewBag.jobyear2 = jobYear;
                    ViewBag.jobState2 = jobState;
                    System.Diagnostics.Debug.WriteLine("C");
                }
            }
            else
            {

                System.Diagnostics.Debug.WriteLine("D");
                jobs = myWebService.GetISPJob(User.Identity.Name, 0, 1000000, null, null, null, null, null);

                ViewBag.searchKey2 = null;
                ViewBag.excabb2 = "Select";
                ViewBag.schemeType2 = "Select";
                ViewBag.jobyear2 = 0;
                ViewBag.jobState2 = "Select";
            }


            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            using (Entities ctxData = new Entities())
            {
                string UserGrp = "";

                var queryUser = (from d in ctxData.WV_USER
                                 where d.USERNAME.ToUpper() == User.Identity.Name.ToUpper() || d.USERNAME.ToLower() == User.Identity.Name.ToLower()
                                 select new { d.PTT_STATE, d.EXC, d.GROUPID });

                foreach (var a in queryUser)
                {
                    UserGrp = a.GROUPID;
                }

                var query = from p in ctxData.REF_FEATURE_STATE
                            select new { Text = p.FEATURE_STATE, Value = p.FEATURE_STATE };

                string jobStateList = "";
                foreach (var a in query)
                {
                    jobStateList = jobStateList + a.Text + "|";
                }

                ViewBag.jobstate = jobStateList.Substring(0, jobStateList.Length - 1);

                if (UserGrp == "12")
                {
                    //filter exchange
                    List<SelectListItem> list = new List<SelectListItem>();
                    var queryEXC = from p in ctxData.WV_ISP_JOB
                                   where p.WEBVIEW == 1
                                   select new { Text = p.EXC_ABB, Value = p.EXC_ABB };

                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in queryEXC.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                    }

                    //filter scheme
                    List<SelectListItem> list2 = new List<SelectListItem>();
                    var querySCHEME = from p in ctxData.WV_ISP_JOB
                                      where p.WEBVIEW == 1
                                      select new { Text = p.JOB_TYPE, Value = p.JOB_TYPE };

                    list2.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in querySCHEME.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list2.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                    }

                    //filter year
                    List<SelectListItem> list3 = new List<SelectListItem>();
                    var queryYEAR = from p in ctxData.WV_ISP_JOB
                                    where p.WEBVIEW == 1
                                    select new { Text = p.YEAR, Value = p.YEAR };

                    list3.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in queryYEAR.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list3.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                    }
                    List<SelectListItem> list4 = new List<SelectListItem>();
                    var queryState = from p in ctxData.G3E_JOB
                                     where p.WEBVIEW == 1
                                     select new { Text = p.JOB_STATE, Value = p.JOB_STATE };

                    list4.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in queryState.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list4.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                    }
                    ViewBag.jobState = list4;
                    ViewBag.excabb = list;
                    ViewBag.schemeType = list2;
                    ViewBag.jobyear = list3;

                }
                else
                {
                    //filter exchange
                    List<SelectListItem> list = new List<SelectListItem>();
                    var queryEXC = from p in ctxData.WV_ISP_JOB
                                   where p.G3E_OWNER.ToUpper() == User.Identity.Name.ToUpper() || p.G3E_OWNER.ToLower() == User.Identity.Name.ToLower()
                                   select new { Text = p.EXC_ABB, Value = p.EXC_ABB };

                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in queryEXC.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                    }

                    //filter scheme
                    List<SelectListItem> list2 = new List<SelectListItem>();
                    var querySCHEME = from p in ctxData.WV_ISP_JOB
                                      where p.G3E_OWNER.ToUpper() == User.Identity.Name.ToUpper() || p.G3E_OWNER.ToLower() == User.Identity.Name.ToLower()
                                      select new { Text = p.JOB_TYPE, Value = p.JOB_TYPE };

                    list2.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in querySCHEME.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list2.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                    }

                    //filter year
                    List<SelectListItem> list3 = new List<SelectListItem>();
                    var queryYEAR = from p in ctxData.WV_ISP_JOB
                                    where p.G3E_OWNER.ToUpper() == User.Identity.Name.ToUpper() || p.G3E_OWNER.ToLower() == User.Identity.Name.ToLower()
                                    select new { Text = p.YEAR, Value = p.YEAR };

                    list3.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in queryYEAR.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list3.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                    } List<SelectListItem> list4 = new List<SelectListItem>();
                    var queryState = from p in ctxData.G3E_JOB
                                     where p.G3E_OWNER.ToUpper() == User.Identity.Name.ToUpper() || p.G3E_OWNER.ToLower() == User.Identity.Name.ToLower()
                                     select new { Text = p.JOB_STATE, Value = p.JOB_STATE };

                    list4.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in queryState.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list4.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                    }
                    ViewBag.jobState = list4;
                    ViewBag.excabb = list;
                    ViewBag.schemeType = list2;
                    ViewBag.jobyear = list3;
                }
            }

            //int page = 1;
            int pageSize = 10;
            int pageNumber = (page ?? 1);
            //ViewBag.boq = boqCheck;
            return View(jobs.JobList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult RNOList(string searchKey, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            WebService._base.OSPJob jobs = new WebService._base.OSPJob();

            if (searchKey != null)
            {
                if (searchKey.Equals(""))
                {
                    jobs = myWebService.GetRNOJob(User.Identity.Name, 0, 1000000, null);
                }
                else
                {
                    jobs = myWebService.GetRNOJob(User.Identity.Name, 0, 1000000, searchKey);
                    ViewBag.searchKey = searchKey;
                }
            }
            else
            {
                jobs = myWebService.GetRNOJob(User.Identity.Name, 0, 1000000, null);
                ViewBag.searchKey = "";
            }

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.REF_FEATURE_STATE
                            select new { Text = p.FEATURE_STATE, Value = p.FEATURE_STATE };

                string jobStateList = "";
                foreach (var a in query)
                {
                    jobStateList = jobStateList + a.Text + "|";
                }

                ViewBag.jobstate = jobStateList.Substring(0, jobStateList.Length - 1);
            }

            //int page = 1;
            int pageSize = 10;
            int pageNumber = (page ?? 1);
            return View(jobs.JobList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult APPList(string searchKey, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            WebService._base.OSPJob jobs = new WebService._base.OSPJob();

            if (searchKey != null)
            {
                if (searchKey.Equals(""))
                {
                    jobs = myWebService.GetAPPJob(User.Identity.Name, 0, 1000000, null);
                }
                else
                {
                    jobs = myWebService.GetAPPJob(User.Identity.Name, 0, 1000000, searchKey);
                    ViewBag.searchKey = searchKey;
                }
            }
            else
            {
                jobs = myWebService.GetAPPJob(User.Identity.Name, 0, 1000000, null);
                ViewBag.searchKey = "";
            }

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.REF_FEATURE_STATE
                            select new { Text = p.FEATURE_STATE, Value = p.FEATURE_STATE };

                string jobStateList = "";
                foreach (var a in query)
                {
                    jobStateList = jobStateList + a.Text + "|";
                }

                ViewBag.jobstate = jobStateList.Substring(0, jobStateList.Length - 1);
            }

            //int page = 1;
            int pageSize = 10;
            int pageNumber = (page ?? 1);
            return View(jobs.JobList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult LNMSList(string searchKey, int? page, string jobExc, string jobScheme, string jobYear)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            WebService._base.OSPJob jobs = new WebService._base.OSPJob();

            if (searchKey != null)
            {
                System.Diagnostics.Debug.WriteLine("A");
                if (searchKey.Equals("") && jobExc.Equals("Select") && jobScheme.Equals("Select") && jobYear.Equals("0"))
                {
                    System.Diagnostics.Debug.WriteLine("B");
                    jobs = myWebService.GetLNMSJob(User.Identity.Name, 0, 1000, null, null, null, null);
                    ViewBag.searchKey = searchKey;
                    ViewBag.excabb2 = jobExc;
                    ViewBag.schemeType2 = jobScheme;
                    ViewBag.jobyear2 = jobYear;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("C");
                    jobs = myWebService.GetLNMSJob(User.Identity.Name, 0, 1000, searchKey, jobExc, jobScheme, jobYear);
                    if (searchKey != "")
                    {
                        ViewBag.searchKey = searchKey;
                    }
                    else
                    {
                        ViewBag.searchKey = "";
                    }
                    ViewBag.excabb2 = jobExc;
                    ViewBag.schemeType2 = jobScheme;
                    ViewBag.jobyear2 = jobYear;
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("D");
                jobs = myWebService.GetLNMSJob(User.Identity.Name, 0, 1000, null, null, null, null);

                ViewBag.searchKey2 = null;
                ViewBag.excabb2 = "Select";
                ViewBag.schemeType2 = "Select";
                ViewBag.jobyear2 = 0;
            }

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.REF_FEATURE_STATE
                            select new { Text = p.FEATURE_STATE, Value = p.FEATURE_STATE };

                string jobStateList = "";
                foreach (var a in query)
                {
                    jobStateList = jobStateList + a.Text + "|";
                }

                ViewBag.jobstate = jobStateList.Substring(0, jobStateList.Length - 1);

                //filter exchange
                List<SelectListItem> list = new List<SelectListItem>();
                var queryEXC = from p in ctxData.G3E_JOB
                               join fx in ctxData.GC_NETELEM on p.G3E_IDENTIFIER equals fx.JOB_ID
                               where p.WEBVIEW == null && p.G3E_IDENTIFIER != "109693"
                               select new { Text = p.G3E_IDENTIFIER, Value = p.G3E_IDENTIFIER };

                list.Add(new SelectListItem() { Text = "", Value = "Select" });

                string listDis = "";
                foreach (var a in queryEXC.Distinct().OrderBy(it => it.Value))
                {
                    string[] exc = a.Text.Split('-');
                    listDis = listDis + exc[0] + ",";
                }
                System.Diagnostics.Debug.WriteLine(listDis);
                string[] listText = listDis.Split(',');

                foreach (string a in listText.Distinct())
                {
                    list.Add(new SelectListItem() { Text = a.ToString(), Value = a.ToString() });
                }
                ViewBag.excabb = list;

                //filter scheme
                List<SelectListItem> list2 = new List<SelectListItem>();
                var querySCHEME = from p in ctxData.G3E_JOB
                                  join fx in ctxData.GC_NETELEM on p.G3E_IDENTIFIER equals fx.JOB_ID
                                  where p.WEBVIEW == null
                                  select new { Text = p.JOB_TYPE, Value = p.JOB_TYPE };

                list2.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in querySCHEME.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list2.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }

                ViewBag.schemeType = list2;

                //filter year
                List<SelectListItem> list3 = new List<SelectListItem>();
                var queryYEAR = from p in ctxData.G3E_JOB
                                join fx in ctxData.GC_NETELEM on p.G3E_IDENTIFIER equals fx.JOB_ID
                                where p.WEBVIEW == null
                                select new { Text = p.YEAR_INSTALL, Value = p.YEAR_INSTALL };

                list3.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in queryYEAR.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list3.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }

                ViewBag.jobyear = list3;
            }

            //int page = 1;
            int pageSize = 10;
            int pageNumber = (page ?? 1);
            return View(jobs.JobList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult NewJob(string job)
        {
            List<SelectListItem> list = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.REF_FEATURE_STATE
                            select new { Text = p.FEATURE_STATE, Value = p.FEATURE_STATE };

                foreach (var a in query)
                {
                    list.Add(new SelectListItem() { Text = a.Text, Value = a.Value.ToString() });
                }

                ViewBag.jobstate = list;
            }

            List<SelectListItem> list3 = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                var queryUsr = (from up in ctxData.WV_USER
                                where up.USERNAME.ToUpper() == User.Identity.Name.ToUpper() || up.USERNAME.ToLower() == User.Identity.Name.ToLower()
                                select up).Single();

                var query = from p in ctxData.WV_EXC_MAST
                            orderby p.EXC_NAME ascending, p.EXC_ABB ascending
                            where p.PTT_ID.Trim() == queryUsr.PTT_STATE.Trim()
                            select new { Text = p.EXC_NAME + " (" + p.EXC_ABB.Trim() + ")", Value = p.EXC_ABB };

                foreach (var a in query)
                {
                    list3.Add(new SelectListItem() { Text = a.Text, Value = a.Value.ToString() });
                }

                ViewBag.excabb = list3;
            }

            List<SelectListItem> list4 = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                var queryUsr = (from up in ctxData.WV_USER
                                where up.USERNAME.ToUpper() == User.Identity.Name.ToUpper() || up.USERNAME.ToLower() == User.Identity.Name.ToLower()
                                select up).Single();
                if (queryUsr.PTT_STATE.Trim() == "ALL" || queryUsr.GROUPID == "7")
                {
                    System.Diagnostics.Debug.WriteLine("ALL");
                    var query = from p in ctxData.WV_EXC_MAST
                                orderby p.PTT_ID ascending
                                select new { p.PTT_ID };

                    foreach (var a in query.Distinct().OrderBy(it => it.PTT_ID))
                    {
                        list4.Add(new SelectListItem() { Text = a.PTT_ID, Value = a.PTT_ID });
                    }
                    ViewBag.checkEXC = "TRUE";
                    ViewBag.pttID = list4;
                }
                else if (queryUsr.GROUPID == "6")
                {
                    System.Diagnostics.Debug.WriteLine("ALL  6");
                    if (queryUsr.PTT_STATE.Trim() == "ALL")
                    {

                        var queryPttMasts = from p in ctxData.WV_EXC_MAST
                                            orderby p.PTT_ID ascending
                                            //where p.PTT_ID.Trim() == queryUsr.PTT_STATE.Trim()
                                            select new { p.PTT_ID };

                        System.Diagnostics.Debug.WriteLine("ALL " + queryUsr.PTT_STATE.Trim());

                        list4.Clear();
                        foreach (var a in queryPttMasts.Distinct().OrderBy(it => it.PTT_ID))
                        {
                            list4.Add(new SelectListItem() { Text = a.PTT_ID, Value = a.PTT_ID });
                        }
                        ViewBag.checkEXC = "TRUE";
                        ViewBag.pttID = list4;
                    }
                    var queryPttMast1 = (from f in ctxData.WV_PTT_MAST
                                         join h in ctxData.WV_EXC_MAST on f.PTT_ID equals h.PTT_ID
                                         join g in ctxData.WV_REGION_MAST on f.REGION_ID equals g.REGION_ID
                                         where g.NATIONWIDE_GRP.Trim() == queryUsr.PTT_STATE.Trim()
                                         select new { h.SEGMENT });

                    System.Diagnostics.Debug.WriteLine("ALL " + queryUsr.PTT_STATE.Trim());

                    list4.Clear();
                    foreach (var a in queryPttMast1.Distinct().OrderBy(it => it.SEGMENT))
                    {
                        list4.Add(new SelectListItem() { Text = a.SEGMENT, Value = a.SEGMENT });
                    }
                    ViewBag.checkEXC = "TRUE";
                    ViewBag.pttID = list4;

                }
                else if (queryUsr.GROUPID == "5")
                {
                    System.Diagnostics.Debug.WriteLine("ALL  5");
                    string region = "";

                    var queryPttMast = (from f in ctxData.WV_PTT_MAST
                                        join h in ctxData.WV_EXC_MAST on f.PTT_ID equals h.PTT_ID
                                        join g in ctxData.WV_REGION_MAST on f.REGION_ID equals g.REGION_ID
                                        where g.REGION_ID.Trim() == queryUsr.PTT_STATE.Trim()
                                        select new { h.SEGMENT });



                    list4.Clear();
                    foreach (var a in queryPttMast.Distinct().OrderBy(it => it.SEGMENT))
                    {
                        list4.Add(new SelectListItem() { Text = a.SEGMENT, Value = a.SEGMENT });
                    }
                    ViewBag.checkEXC = "TRUE";
                    ViewBag.pttID = list4;
                }
                else
                {
                    list4.Add(new SelectListItem() { Text = "SELECT", Value = "SELECT" });
                    ViewBag.pttID = list4;
                    ViewBag.checkEXC = "";
                }

                List<SelectListItem> list2 = new List<SelectListItem>();

                var queryJob = from p in ctxData.REF_JOB_TYPES
                               orderby p.JOB_TYPE ascending
                               select new { Text = p.JOB_TYPE, Value = p.JOB_TYPE };

                //if (queryUsr.GROUPID == "7")
                //{
                //    list2.Add(new SelectListItem() { Text = "CID", Value = "CID" });
                //}
                //else
                //{
                list2.Add(new SelectListItem() { Text = "Select Scheme Type", Value = "" });
                //}

                foreach (var a in queryJob)
                {
                    list2.Add(new SelectListItem() { Text = a.Text, Value = a.Value.ToString() });
                }

                ViewBag.schemetype = list2;

            }


            return View(new WebView.Models.NewJobModel());
        }

        [HttpPost]
        public ActionResult updataListData(string PTT_ID) // FOR USER ALL PTT
        {
            string PuList = "";
            string Mast = "";
            //string PuList2 = "";
            using (Entities ctxData = new Entities())
            {
                List<SelectListItem> list = new List<SelectListItem>();
                var queryEXC = from p in ctxData.WV_EXC_MAST
                               where p.PTT_ID.Trim() == PTT_ID.Trim()
                               orderby p.EXC_ABB
                               select new { p.EXC_ABB, p.EXC_NAME };

                foreach (var a in queryEXC.Distinct().OrderBy(it => it.EXC_NAME))
                {
                    PuList = PuList + a.EXC_NAME + " (" + a.EXC_ABB.Trim() + ") :  " + a.EXC_ABB + "|";
                }
                ViewBag.excabb = list;
                ViewBag.Region = "";
            }

            //return View();
            return Json(new
            {
                Success = true,
                PuList = PuList
            }, JsonRequestBehavior.AllowGet); //
        }

        public ActionResult NewJobRNO(string job)
        {
            List<SelectListItem> list1 = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.GC_ITFACE
                             orderby p.ITFACE_CODE ascending
                             select new { Text = p.CABLE_CODE, Value = p.CABLE_CODE }).ToList();

                foreach (var a in query.Distinct())
                {
                    list1.Add(new SelectListItem() { Text = a.Text, Value = a.Value });
                }

                ViewBag.cabletype = list1;
            }

            List<SelectListItem> list2 = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.GC_ITFACE
                             where p.ITFACE_CLASS == "CABINET"
                             orderby p.ITFACE_CODE ascending
                             select new { Text = p.ITFACE_CODE, Value = p.ITFACE_CODE }).ToList();

                foreach (var a in query.Distinct())
                {
                    list2.Add(new SelectListItem() { Text = a.Text, Value = a.Value });
                }

                ViewBag.schemetype = list2;
            }

            List<SelectListItem> list3 = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                var queryUsr = (from up in ctxData.WV_USER
                                where up.USERNAME.ToUpper() == User.Identity.Name.ToUpper() || up.USERNAME.ToLower() == User.Identity.Name.ToLower()
                                select up).Single();

                var query = from p in ctxData.WV_EXC_MAST
                            where p.PTT_ID.Trim() == queryUsr.PTT_STATE.Trim()
                            select new { Text = p.EXC_NAME + " (" + p.EXC_ABB + ")", Value = p.EXC_ABB };

                foreach (var a in query)
                {
                    list3.Add(new SelectListItem() { Text = a.Text, Value = a.Value.ToString() });
                }

                ViewBag.excabb = list3;
            }
            List<SelectListItem> list4 = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.REF_JOB_TYPES
                            orderby p.JOB_TYPE ascending
                            select new { Text = p.JOB_TYPE, Value = p.JOB_TYPE };

                foreach (var a in query.Distinct())
                {
                    list4.Add(new SelectListItem() { Text = a.Text, Value = a.Value });
                }

                ViewBag.jobtype = list4;
            }

            return View(new WebView.Models.NewJobModel());
        }

        private string GetSchemeType(string c)
        {
            System.Diagnostics.Debug.WriteLine(c);
            string[] sch = new string[13];
            sch[1] = "Fibre";
            sch[2] = "E-Side";
            sch[3] = "D-Side";
            sch[4] = "Civil";
            sch[5] = "Other";
            sch[6] = "Civil CT-AP";
            sch[7] = "Civil CT-IP";
            sch[8] = "Secure";
            sch[9] = "ADSL";
            sch[10] = "HSBB";

            return sch[Convert.ToInt32(c)];
        }

        [HttpPost]
        public ActionResult NewJob(NewJobModel model, IEnumerable<HttpPostedFileBase> files)
        {
            bool success = true;
            string job_types = "";
            string result = "";
            string resultEmail = "";

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;
            using (Entities ctxData = new Entities())
            {
                var query = (from up in ctxData.REF_JOB_TYPES
                             where up.JOB_TYPE.Trim() == model.SchemeType.Trim()
                             select up).Single();

                job_types = query.ISPOSP;
            }

            // create job in OSP (SOAP)
            if (job_types == "OSP")
            {
                selected = true;

                WebService._base.Job newjob = new WebService._base.Job();
                newjob.G3E_DESCRIPTION = model.Description;
                newjob.G3E_IDENTIFIER = "0"; // Stored procedure handle this.
                newjob.WebViewName = "NEPS"; //-----------------------------------------------------------------------------------------
                newjob.G3E_STATE = "PROPOSED";  //model.JobState;
                newjob.G3E_STATUS = "Posted";//
                newjob.EXC_ABB = model.EXC_ABB.Trim();
                newjob.Scheme_Type = model.SchemeType.Trim();
                newjob.G3E_DESCRIPTION_2 = model.Description_2;
                if (model.PlanEndDate < model.PlanStartDate)
                {
                    ModelState.AddModelError("", "Invalid Date Range");
                }
                else
                {
                    newjob.G3E_PlanStartDate = model.PlanStartDate;
                    newjob.G3E_PlanEndDate = model.PlanEndDate;
                    result = myWebService.AddJobOSP(newjob, User.Identity.Name);
                }

                if (result == "fail")
                    success = false;
                else
                {
                    success = true;

                    // Handling Files.                   
                    bool folderCreated = false;
                    foreach (var file in files)
                    {
                        if (file != null)
                        {
                            if (!folderCreated)
                            {
                                //create job directory
                                string newPath = System.IO.Path.Combine(Server.MapPath("~/App_Data/uploads"), "job" + result);
                                System.IO.Directory.CreateDirectory(newPath);
                                System.Diagnostics.Debug.WriteLine(newPath);

                                folderCreated = true;
                            }

                            if (file.ContentLength > 0)
                            {
                                var fileName = Path.GetFileName(file.FileName);
                                DateTime thisDay = DateTime.Now;
                                string extension = Path.GetExtension(file.FileName);
                                string schemeName = "";

                                if (extension == ".xml")
                                {
                                    using (Entities ctxData = new Entities())
                                    {
                                        var query = (from p in ctxData.G3E_JOB
                                                     where p.G3E_IDENTIFIER == result
                                                     select p).Single();

                                        schemeName = query.SCHEME_NAME;
                                    }

                                    var path = Path.Combine(Server.MapPath("~/App_Data/uploads/job" + result + "/" + schemeName + thisDay.ToString("MMdd") + "_" + thisDay.ToString("HH") + "" + thisDay.ToString("mm") + extension));
                                    file.SaveAs(path);
                                }
                                else
                                {
                                    var path = Path.Combine(Server.MapPath("~/App_Data/uploads/job" + result), fileName);
                                    file.SaveAs(path);
                                }
                            }
                        }
                    }
                }
            }

            // create job in ISPOSP (SOAP)
            if (job_types == "ISPOSP")
            {
                selected = true;

                WebService._base.Job newjob = new WebService._base.Job();
                newjob.G3E_DESCRIPTION = model.Description;
                newjob.G3E_IDENTIFIER = "0"; // Stored procedure handle this.
                newjob.WebViewName = "NEPS"; //-----------------------------------------------------------------------------------------
                newjob.G3E_STATE = "PROPOSED";  //model.JobState;
                newjob.G3E_STATUS = "Posted";//
                newjob.EXC_ABB = model.EXC_ABB.Trim();
                newjob.Scheme_Type = model.SchemeType.Trim();
                newjob.G3E_DESCRIPTION_2 = model.Description_2;
                if (model.PlanEndDate < model.PlanStartDate)
                {
                    ModelState.AddModelError("", "Invalid Date Range");
                }
                else
                {
                    newjob.G3E_PlanStartDate = model.PlanStartDate;
                    newjob.G3E_PlanEndDate = model.PlanEndDate;
                    result = myWebService.AddJobISPOSP(newjob, User.Identity.Name);
                }

                if (result == "fail")
                    success = false;
                else
                {
                    success = true;
                    string ispScheme = "";
                    string ispDesc = "";
                    string ispOwner = "";
                    using (Entities ctxData = new Entities())
                    {
                        var query = from p in ctxData.WV_ISP_JOB
                                    where p.G3E_IDENTIFIER == result
                                    select new { p.SCHEME_NAME, p.G3E_DESCRIPTION, p.G3E_OWNER, p.JOB_STATE };


                        foreach (var a in query)
                        {
                            ispScheme = a.SCHEME_NAME;
                            ispDesc = a.G3E_DESCRIPTION;
                            ispOwner = a.G3E_OWNER;
                        }

                    }
                    try
                    {
                        // call ISP web service (using restful)
                        // Test NRM Web service
                        System.Diagnostics.Debug.WriteLine(ispScheme);
                        NrmServiceInterfaceClient testWS = new NrmServiceInterfaceClient();
                        testWS.CreateProject(ispScheme, ispDesc, ispScheme, ispOwner);

                        //System.Diagnostics.Debug.WriteLine(NRM);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(ex.ToString());
                    }
                    // Handling Files.                   
                    bool folderCreated = false;
                    foreach (var file in files)
                    {
                        if (file != null)
                        {
                            if (!folderCreated)
                            {
                                //create job directory
                                string newPath = System.IO.Path.Combine(Server.MapPath("~/App_Data/uploads"), "job" + result);
                                System.IO.Directory.CreateDirectory(newPath);
                                System.Diagnostics.Debug.WriteLine(newPath);

                                folderCreated = true;
                            }

                            if (file.ContentLength > 0)
                            {
                                var fileName = Path.GetFileName(file.FileName);
                                DateTime thisDay = DateTime.Now;
                                string extension = Path.GetExtension(file.FileName);
                                string schemeName = "";

                                if (extension == ".xml")
                                {
                                    using (Entities ctxData = new Entities())
                                    {
                                        var query = (from p in ctxData.G3E_JOB
                                                     where p.G3E_IDENTIFIER == result
                                                     select p).Single();

                                        schemeName = query.SCHEME_NAME;
                                    }

                                    var path = Path.Combine(Server.MapPath("~/App_Data/uploads/job" + result + "/" + schemeName + thisDay.ToString("MMdd") + "_" + thisDay.ToString("HH") + "" + thisDay.ToString("mm") + extension));
                                    file.SaveAs(path);
                                }
                                else
                                {
                                    var path = Path.Combine(Server.MapPath("~/App_Data/uploads/job" + result), fileName);
                                    file.SaveAs(path);
                                }
                            }
                        }
                    }
                }
            }



            // Create job isp
            if (job_types == "ISP")
            {
                selected = true;

                WebService._base.Job newjob = new WebService._base.Job();
                newjob.G3E_DESCRIPTION = model.Description;
                newjob.G3E_IDENTIFIER = "0"; // Stored procedure handle this.
                newjob.WebViewName = "NEPS"; //-----------------------------------------------------------------------------------------
                newjob.G3E_STATE = "PROPOSED";  //model.JobState;
                newjob.G3E_STATUS = "Posted";//
                newjob.EXC_ABB = model.EXC_ABB.Trim();
                newjob.Scheme_Type = model.SchemeType.Trim();
                newjob.G3E_DESCRIPTION_2 = model.Description_2;
                if (model.PlanEndDate < model.PlanStartDate)
                {
                    ModelState.AddModelError("", "Invalid Date Range");
                }
                else
                {
                    newjob.G3E_PlanStartDate = model.PlanStartDate;
                    newjob.G3E_PlanEndDate = model.PlanEndDate;
                    result = myWebService.AddJobISP(newjob, User.Identity.Name);
                }

                //result = myWebService.AddJobISP(newjob, User.Identity.Name);

                string ispScheme = "";
                string ispDesc = "";
                string ispOwner = "";

                if (result == "fail")
                {
                    success = false;
                }
                else
                {
                    success = true;

                    using (Entities ctxData = new Entities())
                    {
                        var query = from p in ctxData.WV_ISP_JOB
                                    where p.G3E_IDENTIFIER == result
                                    select new { p.SCHEME_NAME, p.G3E_DESCRIPTION, p.G3E_OWNER, p.JOB_STATE };


                        foreach (var a in query)
                        {
                            ispScheme = a.SCHEME_NAME;
                            ispDesc = a.G3E_DESCRIPTION;
                            ispOwner = a.G3E_OWNER;
                        }

                    }

                    // call ISP web service (using restful)
                    // Test NRM Web service
                    try
                    {
                        System.Diagnostics.Debug.WriteLine(ispScheme);
                        NrmServiceInterfaceClient testWS = new NrmServiceInterfaceClient();
                        testWS.CreateProject(ispScheme, ispDesc, ispScheme, ispOwner);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(ex.ToString());
                    }

                    // Handling Files.                   
                    bool folderCreated = false;
                    foreach (var file in files)
                    {
                        if (file != null)
                        {
                            if (!folderCreated)
                            {
                                //create job directory
                                string newPath = System.IO.Path.Combine(Server.MapPath("~/App_Data/uploads"), "job" + result);
                                System.IO.Directory.CreateDirectory(newPath);
                                System.Diagnostics.Debug.WriteLine(newPath);

                                folderCreated = true;
                            }

                            if (file.ContentLength > 0)
                            {
                                var fileName = Path.GetFileName(file.FileName);
                                DateTime thisDay = DateTime.Now;
                                string extension = Path.GetExtension(file.FileName);
                                string schemeName = "";

                                if (extension == ".xml")
                                {
                                    using (Entities ctxData = new Entities())
                                    {
                                        var query = (from p in ctxData.WV_ISP_JOB
                                                     where p.G3E_IDENTIFIER == result
                                                     select p).Single();

                                        schemeName = query.SCHEME_NAME;
                                    }

                                    var path = Path.Combine(Server.MapPath("~/App_Data/uploads/job" + result + "/" + schemeName + thisDay.ToString("MMdd") + "_" + thisDay.ToString("HH") + "" + thisDay.ToString("mm") + extension));
                                    file.SaveAs(path);
                                }
                                else
                                {
                                    var path = Path.Combine(Server.MapPath("~/App_Data/uploads/job" + result), fileName);
                                    file.SaveAs(path);
                                }
                                //var path = Path.Combine(Server.MapPath("~/App_Data/uploads"), fileName);
                                //var path = Path.Combine(Server.MapPath("~/App_Data/uploads/job" + result), fileName);
                                //file.SaveAs(path);
                            }
                        }
                    }
                }
            }

            if (ModelState.IsValid && selected)
            {
                if (success == true)
                {
                    string uemail;
                    using (Entities ctxData = new Entities())
                    {
                        var queryUsr = (from p in ctxData.WV_USER
                                        where p.USERNAME == User.Identity.Name
                                        select new { p.EMAIL }).Single();

                        uemail = queryUsr.EMAIL;
                    }

                    //try
                    //{
                    //    MailMessage msg = new MailMessage();
                    //    msg.IsBodyHtml = true;
                    //    msg.From = new MailAddress("neps@tm.com.my", "NEPS");
                    //    msg.To.Add(uemail);
                    //    msg.Subject = "NEPS - JOB CREATED ";
                    //    msg.Body = "<h1>FILES DETAILS</h1>SCHEME NAME	: <br/><br/>DESCRIPTION : <br/><br/>REDMARK FILE NAME: .xml ";
                    //    msg.Body += "<br/><br/> <h1>RNO DETAILS</h1> <br>";
                    //    msg.Body += "RNO ID : " + User.Identity.Name + "<br/><br/>RNO EMAIL	: <br/><br/>RNO PHONE NUMBER: <br/><br/>Please log in to <a href='http://10.41.101.168/'>NEPS WEBVIEW  </a>to download the file.";
                    //    msg.IsBodyHtml = true;
                    //    SmtpClient emailClient = new SmtpClient("smtp.tm.com.my");
                    //    emailClient.UseDefaultCredentials = true;
                    //    emailClient.Port = 25;
                    //    emailClient.EnableSsl = false;
                    //    //emailClient.UseDefaultCredentials = false;
                    //    emailClient.Send(msg);
                    resultEmail = "OK";
                    //}
                    //catch (Exception ex)
                    //{
                    //    resultEmail = ex.ToString();
                    //}
                    //return RedirectToAction("NewSave?res=" + result);
                    return RedirectToAction("NewSave", new { res = result, email = resultEmail });
                }
                else
                    return RedirectToAction("NewSaveFail"); // store to db failed.
            }

            List<SelectListItem> list = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.REF_FEATURE_STATE
                            select new { Text = p.FEATURE_STATE, Value = p.FEATURE_STATE };

                foreach (var a in query)
                {
                    list.Add(new SelectListItem() { Text = a.Text, Value = a.Value.ToString() });
                }

                ViewBag.jobstate = list;
            }

            List<SelectListItem> list2 = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.REF_JOB_TYPES
                            orderby p.JOB_TYPE ascending
                            select new { Text = p.JOB_TYPE, Value = p.JOB_TYPE };

                foreach (var a in query)
                {
                    list2.Add(new SelectListItem() { Text = a.Text, Value = a.Value.ToString() });
                }

                ViewBag.schemetype = list2;
            }

            List<SelectListItem> list3 = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.WV_EXC_MAST
                            orderby p.EXC_NAME ascending
                            //where p.PTT_ID 
                            select new { Text = p.EXC_NAME, Value = p.EXC_ABB };

                foreach (var a in query)
                {
                    list3.Add(new SelectListItem() { Text = a.Text, Value = a.Value.ToString() });
                }

                ViewBag.excabb = list3;
            }

            return View(model);
        }

        [HttpPost]
        public ActionResult NewJobRNO(NewJobModel model, IEnumerable<HttpPostedFileBase> files)
        {
            bool success = true;
            string result = "";
            string schemeName = "";
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;

            // create job in RNO
            selected = true;
            WebService._base.Job newjob = new WebService._base.Job();
            newjob.G3E_DESCRIPTION = model.Description;
            newjob.G3E_IDENTIFIER = "0"; // Stored procedure handle this.
            newjob.WebViewName = "NEPS"; //-----------------------------------------------------------------------------------------
            newjob.G3E_STATE = "PENDING";  //model.JobState;
            newjob.G3E_STATUS = "PENDING";//
            if (model.EXC_ABB == "M*" || model.EXC_ABB == "***")
            {
                newjob.EXC_ABB = model.EXC_ABB.Trim();
            }
            else
            {
                newjob.EXC_ABB = model.EXC_ABB;
            }
            newjob.G3E_WORK_ORDER_ID = model.JobState.Trim();
            newjob.Scheme_Type = model.SchemeType.Trim();


            result = myWebService.AddJobRNO(newjob, User.Identity.Name);

            DateTime thisDay = DateTime.Now;
            string resultRedmark = "";
            //string emailList = "";
            if (result == "fail")
                success = false;
            else
            {
                success = true;

                // Handling Files.                   
                bool folderCreated = false;
                foreach (var file in files)
                {
                    if (file != null)
                    {
                        if (!folderCreated)
                        {
                            //create job directory
                            string newPath = System.IO.Path.Combine(Server.MapPath("~/App_Data/uploads"), "job" + result);
                            System.IO.Directory.CreateDirectory(newPath);
                            System.Diagnostics.Debug.WriteLine(newPath);

                            folderCreated = true;
                        }

                        if (file.ContentLength > 0)
                        {
                            var fileName = Path.GetFileName(file.FileName);

                            string extension = Path.GetExtension(file.FileName);

                            if (extension == ".xml")
                            {
                                string EmailListStr = "";
                                using (Entities ctxData = new Entities())
                                {
                                    var queryJOB = (from p in ctxData.WV_NONNETWORK_JOB
                                                    where p.G3E_IDENTIFIER == result
                                                    select p).Single();

                                    schemeName = queryJOB.SCHEME_NAME;

                                    var queryUsr = (from up in ctxData.WV_USER
                                                    where up.USERNAME == User.Identity.Name
                                                    select up).Single();

                                    var query = (from up in ctxData.WV_USER
                                                 where up.GROUPID == "16" && up.PTT_STATE.Trim() == queryUsr.PTT_STATE.Trim()
                                                 select up);

                                    int i = 0;

                                    EmailListStr = queryUsr.EMAIL + ",";

                                    foreach (var emailAND in query)
                                    {
                                        EmailListStr += emailAND.EMAIL + ",";
                                    }
                                    System.Diagnostics.Debug.WriteLine(EmailListStr);

                                    EmailListStr = EmailListStr.TrimEnd(',');

                                    System.Diagnostics.Debug.WriteLine(EmailListStr);
                                    try
                                    {
                                        MailMessage msg = new MailMessage();
                                        msg.IsBodyHtml = true;
                                        msg.From = new MailAddress("neps@tm.com.my", "NEPS");
                                        msg.To.Add(EmailListStr);
                                        msg.Subject = "New Redmark File for Scheme " + schemeName;
                                        msg.Body = "<h1>FILES DETAILS</h1>SCHEME NAME	: " + schemeName + "<br/><br/>JOB NO : " + result + "<br/><br/>DESCRIPTION : " + newjob.G3E_DESCRIPTION + " <br/><br/>REDMARK FILE NAME: " + schemeName + "-" + thisDay.ToString("HH") + "" + thisDay.ToString("mm") + ".xml ";
                                        msg.Body += "<br/><br/> <h1>RNO DETAILS</h1> <br>";
                                        msg.Body += "RNO ID : " + User.Identity.Name + "<br/><br/>RNO NAME	: " + queryUsr.FULL_NAME + "<br/><br/>RNO EMAIL	: " + queryUsr.EMAIL + "<br/><br/>RNO PHONE NUMBER: " + queryUsr.NO_TEL + "<br/><br/>Please log in to <a href='http://10.41.101.168/'>NEPS WEBVIEW  </a>to download the file.";
                                        msg.IsBodyHtml = true;
                                        SmtpClient emailClient = new SmtpClient("smtp.tm.com.my", 25);
                                        emailClient.UseDefaultCredentials = false;
                                        emailClient.Credentials = new NetworkCredential("neps", "nepsadmin", "tmmaster");
                                        emailClient.EnableSsl = false;
                                        emailClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                                        emailClient.Send(msg);

                                        System.Diagnostics.Debug.WriteLine("OK :" + EmailListStr);
                                        resultRedmark = "ok";
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine(ex.ToString());
                                        resultRedmark = ex.ToString();
                                    }
                                }
                                schemeName = schemeName.Replace("/", "");
                                schemeName = schemeName.Replace("*", "");
                                var path = Path.Combine(Server.MapPath("~/App_Data/uploads/job" + result + "/" + schemeName + "-" + thisDay.ToString("HH") + "" + thisDay.ToString("mm") + extension));
                                file.SaveAs(path);
                            }
                            else
                            {
                                var path = Path.Combine(Server.MapPath("~/App_Data/uploads/job" + result), fileName);
                                file.SaveAs(path);
                            }
                        }
                    }
                }
            }

            if (ModelState.IsValid && selected)
            {
                if (success == true)
                    //return RedirectToAction("NewSave?res=" + result);
                    return RedirectToAction("NewSave", new { res = result, email = resultRedmark });
                else
                    return RedirectToAction("NewSaveFail"); // store to db failed.
            }

            return View(model);
        }

        public ActionResult Test()
        {
            Tools tool = new Tools();

            string a = tool.ExecuteStr("Data Source =NEPSTRN; User Id =neps; Password =nepstrn; ", "SELECT * FROM G3E_JOB");

            ViewBag.num = a;

            return View();
        }

        [HttpPost]
        public ActionResult New(JobModel model)
        {
            // move to web service

            using (Entities ctxData = new Entities())
            {
                Tools tool = new Tools();
                tool.ExecuteSql(ctxData, "INSERT INTO WV_SCHEME_COUNTER (YEAR, COUNTER) VALUES (2012, 0) ");
            }
            return RedirectToAction("NewSave");
        }

        public ActionResult NewSave(string res, string email)
        {
            ViewBag.jobId = res;
            ViewBag.email = email;
            return View();
        }

        [HttpPost]
        public ActionResult Edit(JobModel model)
        {
            using (Entities ctxData = new Entities())
            {
                ctxData.SaveChanges();
            }

            return RedirectToAction("List");
        }

        //[Authorize]
        public ActionResult Notification()
        {
            /*
            WebView.WebService._base myWebService;
            WebView.WebService._base.BookShop shop;
            WebView.WebService._base.Book[] books;

            myWebService = new WebService._base();
            shop = myWebService.GetBookShop();
            books = shop.booklist;

            foreach (WebView.WebService._base.Book a in books)
            {
                System.Diagnostics.Debug.WriteLine("Book name: " + a.name);
                System.Diagnostics.Debug.WriteLine("Author: " + a.author);
            }

            myWebService.GetOSPJob(5, 5, null);

            // to test sent
            WebService._base.Book book = new WebService._base.Book();
            book.author = "Halim";
            book.name = "Business";
            myWebService.AddBook(book);
            */

            /*
            using (Entities ctxData = new Entities())
            {
                var query = (from d in ctxData.G3E_JOB
                             orderby d.G3E_ID
                             select new { d.G3E_IDENTIFIER, d.G3E_ID }).Skip(3).Take(5);
             * /

                /*
                var row =
                    query
                    .Select((obj, Index) => new { obj.G3E_IDENTIFIER, obj.G3E_ID, Index = Index });
                */
            /*
                foreach (var a in query)
                {
                    System.Diagnostics.Debug.WriteLine(a.G3E_ID + " : " + a.G3E_IDENTIFIER);
                }
            }

            string sqlcmd = "select * from  (select xyz.*, rownum rnum from ";
            sqlcmd += "(SELECT X.G3E_IDENTIFIER, X.WORK_ORDER_ID, X.G3E_ID FROM G3E_Job X ORDER bY X.G3E_IDENTIFIER) ";
            sqlcmd += "xyz where rownum <=10) where rnum >= 5";
            Execute("Data Source =nova; User Id =NEPS; Password =NEPS; ", sqlcmd);
            */

            return View();
        }

        // Load Job details
        [HttpPost]
        public ActionResult GetDetails(string id)
        {
            System.Diagnostics.Debug.WriteLine(id);
            System.Diagnostics.Debug.WriteLine("test!!");
            Tools tool = new Tools();
            string path = Server.MapPath("~/App_Data/uploads");
            string[] fileList = tool.GetFileList(path + "/job" + id);
            string fileListStr = "";

            if (fileList.Count() > 0)
            {
                for (int i = 0; i < fileList.Count(); i++)
                {
                    fileListStr += fileList[i] + "|";
                }
            }

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            string ptt;
            string grpId;
            string hndover;
            int outputHandover = 0;
            int outputHandover2 = 0;
            string checkStatusJob;
            #region Fatihin - CR80 - 2018 - 29 Jan 2018
            int checkGRN = 0;
            int checkNIS = 0;
            int checkGRNISP = 0;
            int checkNISISP = 0;
            #endregion

            //using (Entities ctxData = new Entities())
            //{
            //    var query = (from p in ctxData.WV_USER
            //                 where p.USERNAME.ToUpper() == User.Identity.Name.ToUpper() || p.USERNAME.ToLower() == User.Identity.Name.ToLower()
            //                 select p).Single();
            //    ptt = query.PTT_STATE;
            //    grpId = query.GROUPID;

            //    //var query2 = (from p in ctxData.G3E_JOB
            //    //             where p.G3E_IDENTIFIER == id
            //    //             select new { p.JOB_STATE }).Single(); // check status
            //    //checkStatusJob = query2.JOB_STATE;
            //}

            using (EntitiesNetworkElement ctxHandover = new EntitiesNetworkElement())
            {
                var queryHandover = (from a in ctxHandover.BI_PROCESS
                                     where a.NEPS_JOB_ID.Trim() == id.Trim() && a.STATUS == "SUCCESS"
                                     select a).Count();


                var queryHandover2 = (from a in ctxHandover.BI_PROCESS
                                      where a.NEPS_JOB_ID.Trim() == id.Trim()
                                      select a).Count();

                outputHandover = queryHandover; // update status by BI_PROSESS
                outputHandover2 = queryHandover2;
                if (outputHandover2 == 0)// if BI_PROCESS = 0 check BI_PROC_GRN_ISP
                {
                    var queryHandoverGRNs = (from a in ctxHandover.BI_PROC_GRN_ISP
                                             where a.NEPS_JOB_ID.Trim() == id.Trim() && a.STATUS == "SUCCESS"
                                             select a).Count(); //check status

                    outputHandover = queryHandoverGRNs; // update status by GRANITE

                    var queryHandoverGRN = (from a in ctxHandover.BI_PROC_GRN_ISP
                                            where a.NEPS_JOB_ID.Trim() == id.Trim()
                                            select a).Count();
                    outputHandover2 = queryHandoverGRN;

                    if (outputHandover2 == 0) // if BI_PROC_GRN_ISP = 0 check BI_PROC_GRN_OSP
                    {
                        var queryHandoverGRN1s = (from a in ctxHandover.BI_PROC_GRN_OSP
                                                  where a.NEPS_JOB_ID.Trim() == id.Trim() && a.STATUS == "SUCCESS"
                                                  select a).Count(); //check status

                        outputHandover = queryHandoverGRN1s; // update status by GRANITE

                        var queryHandoverGRN1 = (from a in ctxHandover.BI_PROC_GRN_OSP
                                                 where a.NEPS_JOB_ID.Trim() == id.Trim()
                                                 select a).Count();

                        outputHandover2 = queryHandoverGRN1;
                    }
                }
            }

            System.Diagnostics.Debug.WriteLine("TEST HANDOVER" + outputHandover2);

            using (Entities ctxData = new Entities())
            {
                var queryUsr = (from p in ctxData.WV_USER
                                where p.USERNAME.ToUpper() == User.Identity.Name.ToUpper() || p.USERNAME.ToLower() == User.Identity.Name.ToLower()
                                select p).Single();
                ptt = queryUsr.PTT_STATE;
                grpId = queryUsr.GROUPID;

                string res = myWebService.GetOwnerList(ptt, grpId);

                var query = (from p in ctxData.G3E_JOB
                             where p.G3E_IDENTIFIER == id
                             select p).Single();

                //var queryUsr = (from p in ctxData.WV_USER
                //                where p.USERNAME.ToUpper() == query.G3E_OWNER.ToUpper() || p.USERNAME.ToLower() == query.G3E_OWNER.ToLower()
                //                select p).Single();
                hndover = queryUsr.HANDOVER;


                int checkfeature = 1;
                string dts = Convert.ToDateTime(query.PLAN_START_DATE).ToString("dd-MMMM-yyyy");
                string dtsEnd = Convert.ToDateTime(query.PLAN_END_DATE).ToString("dd-MMMM-yyyy");

                if (query.JOB_STATE != "COMMISSIONED" && outputHandover >= 1)
                {
                    myWebService.AutoApproval(id);
                }
                #region Fatihin - CR80 - 2018 - 29 Jan 2018
                using (EntitiesNetworkElement ctxHandover = new EntitiesNetworkElement())
                {
                    checkGRN = (from a in ctxHandover.BI_PROC_GRN_OSP
                                where a.NEPS_JOB_ID.Trim() == id.Trim() && (a.STATUS == "NEW" || a.STATUS == "IN PROGRESS")
                                select a).Count();

                    checkNIS = (from a in ctxHandover.BI_PROCESS
                                where a.NEPS_JOB_ID.Trim() == id.Trim() && (a.STATUS == "NEW" || a.STATUS == "IN PROGRESS")
                                select a).Count();

                    checkGRNISP = (from a in ctxHandover.BI_PROC_GRN_ISP
                                   where a.NEPS_JOB_ID.Trim() == id.Trim() && (a.STATUS == "NEW" || a.STATUS == "IN PROGRESS")
                                   select a).Count();

                    checkNISISP = (from a in ctxHandover.BI_PROCESS_ISP
                                   where a.NEPS_JOB_ID.Trim() == id.Trim() && (a.STATUS == "NEW" || a.STATUS == "IN PROGRESS")
                                   select a).Count();
                }
                #endregion
                return Json(new
                {
                    Success = true,
                    jobid = id,
                    record = res,
                    name = query.WEBVIEW_NAME,
                    description = query.G3E_DESCRIPTION,
                    description2 = query.G3E_DESCRIPTION_2,
                    jobstate = query.JOB_STATE,
                    jobstatus = query.G3E_STATUS,
                    project_no = query.PROJECT_NO,
                    scheme_name = query.SCHEME_NAME,
                    excabb = query.EXC_ABB,
                    user = query.G3E_OWNER,
                    handover = hndover,
                    userRole = queryUsr.USER_ROLES,
                    scheme_type = query.SCHEME_TYPE,
                    outputHandover = outputHandover2,
                    grpId = grpId,
                    checkfeature = checkfeature,
                    StartDate = dts,
                    EndDate = dtsEnd,
                    #region Fatihin - CR80 - 2018 - 29 Jan 2018
                    checkGRN = checkGRN,
                    checkNIS = checkNIS,
                    checkGRNISP = checkGRNISP,
                    checkNISISP = checkNISISP,
                    #endregion
                    files_list = new
                    {
                        files = fileListStr
                    }
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        #region Fatihin - CR80 - 2018 - 14 April 2018
        [HttpPost]
        public ActionResult GetStatus(string id)
        {
            int checkGRN = 0;
            int checkNIS = 0;
            int checkGRNISP = 0;
            int checkNISISP = 0;

            using (EntitiesNetworkElement ctxHandover = new EntitiesNetworkElement())
            {
                checkGRN = (from a in ctxHandover.BI_PROC_GRN_OSP
                            where a.NEPS_JOB_ID.Trim() == id.Trim() && (a.STATUS == "NEW" || a.STATUS == "IN PROGRESS")
                            select a).Count();

                checkNIS = (from a in ctxHandover.BI_PROCESS
                            where a.NEPS_JOB_ID.Trim() == id.Trim() && (a.STATUS == "NEW" || a.STATUS == "IN PROGRESS")
                            select a).Count();

                checkGRNISP = (from a in ctxHandover.BI_PROC_GRN_ISP
                               where a.NEPS_JOB_ID.Trim() == id.Trim() && (a.STATUS == "NEW" || a.STATUS == "IN PROGRESS")
                               select a).Count();

                checkNISISP = (from a in ctxHandover.BI_PROCESS_ISP
                               where a.NEPS_JOB_ID.Trim() == id.Trim() && (a.STATUS == "NEW" || a.STATUS == "IN PROGRESS")
                               select a).Count();
            }
            return Json(new
            {
                checkGRN = checkGRN,
                checkNIS = checkNIS,
                checkGRNISP = checkGRNISP,
                checkNISISP = checkNISISP
            }, JsonRequestBehavior.AllowGet);
        }
        #endregion

        [HttpPost]
        public ActionResult ISPGetDetails(string id)
        {
            Tools tool = new Tools();
            string path = Server.MapPath("~/App_Data/uploads");
            string[] fileList = tool.GetFileList(path + "/job" + id);
            string fileListStr = "";
            int outputHandover = 0;
            int outputHandover2 = 0;
            #region Fatihin - CR80 - 2018 - 29 Jan 2018
            /*int checkGRNISP = 0;
            int checkNISISP = 0;*/
            #endregion

            if (fileList.Count() > 0)
            {
                for (int i = 0; i < fileList.Count(); i++)
                {
                    fileListStr += fileList[i] + "|";
                }
            }

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            string ptt;
            string grpId;
            string hndover;
            //string ProjectNRMID;
            string exc_abb;
            string checkStatusJob;
            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_USER
                             where p.USERNAME.ToUpper() == User.Identity.Name.ToUpper() || p.USERNAME.ToLower() == User.Identity.Name.ToLower()
                             select p).Single();
                ptt = query.PTT_STATE;
                grpId = query.GROUPID;

                var queryNRM = (from p in ctxData.WV_ISP_JOB
                                where p.G3E_IDENTIFIER == id
                                select p).Single();
                exc_abb = queryNRM.EXC_ABB;

                checkStatusJob = queryNRM.JOB_STATE;
            }
            //if (checkStatusJob == "COMPLETED" || checkStatusJob == "COMMISSIONED") // make loading on certain job only
            //{
            using (EntitiesNetworkElement ctxHandover = new EntitiesNetworkElement())
            {
                var queryHandover = (from a in ctxHandover.BI_PROCESS_ISP
                                     where a.NEPS_JOB_ID.Trim() == id.Trim() && a.STATUS == "SUCCESS"
                                     select a).Count();

                var queryHandover2 = (from a in ctxHandover.BI_PROCESS_ISP
                                      where a.NEPS_JOB_ID.Trim() == id.Trim()
                                      select a).Count();

                outputHandover = queryHandover;
                outputHandover2 = queryHandover2;
                if (outputHandover2 == 0) // kalau process_isp 0, teruskan dengan grn_isp
                {
                    var queryHandoverGRNs = (from a in ctxHandover.BI_PROC_GRN_ISP
                                             where a.NEPS_JOB_ID.Trim() == id.Trim() && a.STATUS == "SUCCESS"
                                             select a).Count(); //check status

                    outputHandover = queryHandoverGRNs; // update status by GRANITE

                    var queryHandoverGRN = (from a in ctxHandover.BI_PROC_GRN_ISP
                                            where a.NEPS_JOB_ID.Trim() == id.Trim()
                                            select a).Count();
                    outputHandover2 = queryHandoverGRN;

                    if (outputHandover2 == 0) // if BI_PROC_GRN_ISP = 0 check BI_PROC_GRN_OSP
                    {
                        var queryHandoverGRN1s = (from a in ctxHandover.BI_PROC_GRN_OSP
                                                  where a.NEPS_JOB_ID.Trim() == id.Trim() && a.STATUS == "SUCCESS"
                                                  select a).Count(); //check status

                        outputHandover = queryHandoverGRN1s; // update status by GRANITE

                        var queryHandoverGRN1 = (from a in ctxHandover.BI_PROC_GRN_OSP
                                                 where a.NEPS_JOB_ID.Trim() == id.Trim()
                                                 select a).Count();

                        outputHandover2 = queryHandoverGRN1;
                    }
                }
            }
            #region Fatihin - CR80 - 2018 - 29 Jan 2018
            /*using (EntitiesNetworkElement ctxHandover = new EntitiesNetworkElement())
            {
                checkGRNISP = (from a in ctxHandover.BI_PROC_GRN_ISP
                               where a.NEPS_JOB_ID.Trim() == id.Trim() &&  (a.STATUS == "NEW" || a.STATUS == "IN PROGRESS")
                            select a).Count();

                checkNISISP = (from a in ctxHandover.BI_PROCESS_ISP
                               where a.NEPS_JOB_ID.Trim() == id.Trim() && (a.STATUS == "NEW" || a.STATUS == "IN PROGRESS")
                            select a).Count();
            }*/
            #endregion
            //}
            System.Diagnostics.Debug.WriteLine("TEST" + outputHandover);

            string res = myWebService.GetOwnerList(ptt, grpId);


            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_ISP_JOB
                             where p.G3E_IDENTIFIER == id
                             select p).Single();

                var queryUsr = (from p in ctxData.WV_USER
                                where p.USERNAME == query.G3E_OWNER
                                select p).Single();
                hndover = queryUsr.HANDOVER;

                if (query.JOB_STATE != "COMMISSIONED" && outputHandover >= 1)
                {
                    myWebService.ISPAutoApproval(id);
                }

                string dts = Convert.ToDateTime(query.PLAN_START_DATE).ToString("dd-MMMM-yyyy");
                string dtsEnd = Convert.ToDateTime(query.PLAN_END_DATE).ToString("dd-MMMM-yyyy");

                return Json(new
                {
                    Success = true,
                    jobid = id,
                    record = res,
                    name = query.WEBVIEW_NAME,
                    description = query.G3E_DESCRIPTION,
                    description2 = query.G3E_DESCRIPTION_2,
                    jobstate = query.JOB_STATE,
                    jobstatus = query.G3E_STATUS,
                    project_no = query.PROJECT_NO,
                    scheme_name = query.SCHEME_NAME,
                    excabb = query.EXC_ABB,
                    user = query.G3E_OWNER,
                    handover = hndover,
                    userRole = queryUsr.USER_ROLES,
                    scheme_type = query.SCHEME_TYPE,
                    outputHandover = outputHandover2,
                    startDate = dts,
                    endDate = dtsEnd,
                    grpId = grpId,
                    #region Fatihin - CR80 - 2018 - 29 Jan 2018
                    /*checkGRNISP = checkGRNISP,
                    checkNISISP = checkNISISP,*/
                    #endregion
                    files_list = new
                    {
                        files = fileListStr
                    }
                }, JsonRequestBehavior.AllowGet); //
            }

        }

        [HttpPost]
        public ActionResult GetDetailsAPP(string id)
        {
            System.Diagnostics.Debug.WriteLine(id);

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            string ptt;
            string grpId;
            string hndover;
            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_USER
                             where p.USERNAME.ToUpper() == User.Identity.Name.ToUpper() || p.USERNAME.ToLower() == User.Identity.Name.ToLower()
                             select p).Single();
                ptt = query.PTT_STATE;
                grpId = query.GROUPID;
            }

            System.Diagnostics.Debug.WriteLine("PTT :" + ptt + "  GRPID : " + grpId);
            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_USER_APPROVE
                             where p.JOB_ID == id && (p.APPROVAL_USER.ToUpper() == User.Identity.Name.ToUpper() || p.APPROVAL_USER.ToLower() == User.Identity.Name.ToLower())
                             select p).Single();

                return Json(new
                {
                    Success = true,
                    jobid = id,
                    description = query.REMARKS,
                    jobstate = query.STATUS,
                    jobstatus = query.GLOBAL_STATUS,
                    user = query.APPROVAL_USER
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        [HttpPost]
        public ActionResult GetDetails_G3E_JOB(string id)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            string schemeName = "";
            string userdet;
            string excabb = "";

            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_USER
                             where p.USERNAME.ToUpper() == User.Identity.Name.ToUpper() || p.USERNAME.ToLower() == User.Identity.Name.ToLower()
                             select p).Single();
                userdet = query.USERNAME;

                //var queryNetelem = from p in ctxData.GC_NETELEM
                //               where p.SCHEME_NAME == id
                //               select new { p.EXC_ABB };

                //foreach (var a in queryNetelem.Distinct())
                //{
                //    excabb = a.EXC_ABB;
                //}
                string[] aa = id.Split('-');
                string aaa = aa[0];
                var queryJob = from p in ctxData.G3E_JOB
                               where p.G3E_OWNER == userdet && p.WEBVIEW == 1 && p.SCHEME_NAME.Contains(aaa)
                               orderby p.G3E_IDENTIFIER
                               select new { p.G3E_IDENTIFIER };

                foreach (var a in queryJob)
                {
                    schemeName = schemeName + a.G3E_IDENTIFIER + ":" + a.G3E_IDENTIFIER + "|";
                }
                ViewBag.jobid = id;
                return Json(new
                {
                    Success = true,
                    jobid = id,
                    schemeName = schemeName
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        public class Crypto
        {
            private static byte[] _salt = Encoding.ASCII.GetBytes("o6806642kbM7c5");
            private const string sharedSecret = "NEPS";

            /// <summary>
            /// Encrypt the given string using AES.  The string can be decrypted using 
            /// DecryptStringAES().  The sharedSecret parameters must match.
            /// </summary>
            /// <param name="plainText">The text to encrypt.</param>
            /// <param name="sharedSecret">A password used to generate a key for encryption.</param>
            public static string EncryptStringAES(string plainText)
            {
                if (string.IsNullOrEmpty(plainText))
                    return "";

                string outStr = null;                       // Encrypted string to return
                RijndaelManaged aesAlg = null;              // RijndaelManaged object used to encrypt the data.

                try
                {
                    // generate the key from the shared secret and the salt
                    Rfc2898DeriveBytes key = new Rfc2898DeriveBytes(sharedSecret, _salt);

                    // Create a RijndaelManaged object
                    // with the specified key and IV.
                    aesAlg = new RijndaelManaged();
                    aesAlg.Key = key.GetBytes(aesAlg.KeySize / 8);
                    aesAlg.IV = key.GetBytes(aesAlg.BlockSize / 8);

                    // Create a decrytor to perform the stream transform.
                    ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);

                    // Create the streams used for encryption.
                    using (MemoryStream msEncrypt = new MemoryStream())
                    {
                        using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                        {
                            using (StreamWriter swEncrypt = new StreamWriter(csEncrypt))
                            {

                                //Write all data to the stream.
                                swEncrypt.Write(plainText);
                            }
                        }
                        outStr = Convert.ToBase64String(msEncrypt.ToArray());
                    }
                }
                finally
                {
                    // Clear the RijndaelManaged object.
                    if (aesAlg != null)
                        aesAlg.Clear();
                }

                // Return the encrypted bytes from the memory stream.
                return outStr;
            }

            /// <summary>
            /// Decrypt the given string.  Assumes the string was encrypted using 
            /// EncryptStringAES(), using an identical sharedSecret.
            /// </summary>
            /// <param name="cipherText">The text to decrypt.</param>
            /// <param name="sharedSecret">A password used to generate a key for decryption.</param>
            public static string DecryptStringAES(string cipherText)
            {
                if (string.IsNullOrEmpty(cipherText))
                    return "";

                // Declare the RijndaelManaged object
                // used to decrypt the data.
                RijndaelManaged aesAlg = null;

                // Declare the string used to hold
                // the decrypted text.
                string plaintext = null;

                try
                {
                    // generate the key from the shared secret and the salt
                    Rfc2898DeriveBytes key = new Rfc2898DeriveBytes(sharedSecret, _salt);

                    // Create a RijndaelManaged object
                    // with the specified key and IV.
                    aesAlg = new RijndaelManaged();
                    aesAlg.Key = key.GetBytes(aesAlg.KeySize / 8);
                    aesAlg.IV = key.GetBytes(aesAlg.BlockSize / 8);

                    // Create a decrytor to perform the stream transform.
                    ICryptoTransform decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);
                    // Create the streams used for decryption.                
                    byte[] bytes = Convert.FromBase64String(cipherText);
                    using (MemoryStream msDecrypt = new MemoryStream(bytes))
                    {
                        using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                        {
                            using (StreamReader srDecrypt = new StreamReader(csDecrypt))

                                // Read the decrypted bytes from the decrypting stream
                                // and place them in a string.
                                plaintext = srDecrypt.ReadToEnd();
                        }
                    }
                }
                finally
                {
                    // Clear the RijndaelManaged object.
                    if (aesAlg != null)
                        aesAlg.Clear();
                }

                return plaintext;
            }
        }

        class Encryptor
        {
            private static byte[] CreateKeyBytes(string passPhrase)
            {
                char[] keyChars = passPhrase.ToCharArray();
                byte[] keyByte = new byte[16];
                for (int i = 0; i < 16; i++)
                {
                    if (keyChars.Length > i)
                        keyByte[i] = (byte)keyChars[i];
                }

                if (keyChars.Length < 16)
                {
                    for (int i = keyChars.Length; i < 16; i++)
                        keyByte[i] = 0x00;
                }

                return keyByte;

            }

            public static string Encrypt(Encoding encoding, string strtoencrypt, string key, string iv, CipherMode mode, PaddingMode padding, int blocksize)
            {

                var mstream = new MemoryStream();
                using (var aes = new AesManaged())
                {
                    var keybytes = CreateKeyBytes(key);
                    aes.BlockSize = blocksize;
                    aes.KeySize = keybytes.Length * 8;
                    aes.Key = keybytes;

                    aes.IV = CreateKeyBytes(iv);
                    aes.Mode = mode;
                    aes.Padding = padding;


                    using (var cstream = new CryptoStream(mstream, aes.CreateEncryptor(aes.Key, aes.IV), CryptoStreamMode.Write))
                    {
                        var bytesToEncrypt = encoding.GetBytes(strtoencrypt);
                        cstream.Write(bytesToEncrypt, 0, bytesToEncrypt.Length);
                        cstream.FlushFinalBlock();
                    }

                }

                var encrypted = mstream.ToArray();
                return Convert.ToBase64String(encrypted);
            }


            public static string GetURL(string passPhrase, string hostname, string port, string username, List<string> roles)
            {

                string output = "";


                string rolesCommaSeparated = ListToCommaSeparated(roles);
                var key = passPhrase;
                var iv = "";

                string encryptedUsername = Encrypt(Encoding.ASCII, username, key, iv, CipherMode.CBC, PaddingMode.PKCS7, 128);
                string encryptedRoles = Encrypt(Encoding.ASCII, rolesCommaSeparated, key, iv, CipherMode.CBC, PaddingMode.PKCS7, 128);


                //output = string.Format("http://{0}:{1}/nrm?user={2}&roles={3}", hostname, port,
                //    HttpUtility.UrlEncode(encryptedUsername), HttpUtility.UrlEncode(encryptedRoles));

                output = string.Format("?user={2}&roles={3}", hostname, port,
                    HttpUtility.UrlEncode(encryptedUsername), HttpUtility.UrlEncode(encryptedRoles));

                return output;
            }

            private static string ListToCommaSeparated(List<string> roles)
            {
                string output = "";

                int i = 0;
                foreach (string role in roles)
                {
                    if (i++ > 0)
                        output += ",";
                    output += role;
                }
                return output;
            }




        }

        [HttpPost]
        public ActionResult ISPLaunchNRM(string id)
        {
            string exc_abb = "";
            string scheme_name = "";
            string desc_job = "";
            string ProjectNRMID = "";
            string role = "";
            string username = "";
            List<string> roles = new List<string>();
            Tools tool = new Tools();

            using (Entities ctxData = new Entities())
            {
                var queryNRM = (from p in ctxData.WV_ISP_JOB
                                where p.G3E_IDENTIFIER == id
                                select p).SingleOrDefault();
                exc_abb = queryNRM.EXC_ABB;
                scheme_name = queryNRM.SCHEME_NAME;
                desc_job = queryNRM.G3E_DESCRIPTION;

               var queryUser = (from p in ctxData.WV_USER
                                 where p.USERNAME.ToUpper() == User.Identity.Name.ToUpper() || p.USERNAME.ToLower() == User.Identity.Name.ToLower()
                                select p).SingleOrDefault();

                var queryUserRole = (from fx in ctxData.WV_GROUP
                                     join fxx in ctxData.WV_GRP_ROLE on fx.GRPNAME equals fxx.GRPNAME
                                     where fx.GRP_ID == queryUser.GROUPID
                                     select fxx).SingleOrDefault();

                role = queryUserRole.ROLENAME;

                username = queryUser.USERNAME;

            }
            using (Entities_NRM b = new Entities_NRM())
            {
                // check in NRM Project
                var queryCheckProject = (from p in b.PROJECTs
                                         where p.NAME.Trim() == scheme_name.Trim()
                                         select p).Count();
                System.Diagnostics.Debug.WriteLine("QUERY " + queryCheckProject);
                if (queryCheckProject == 0)
                {
                    try
                    {
                        // call ISP web service (add project in NRM)
                        NrmServiceInterfaceClient testWS = new NrmServiceInterfaceClient();
                        string NRM = testWS.CreateProject(scheme_name, desc_job, scheme_name, username);

                        testWS.UpdateProject(scheme_name, scheme_name, scheme_name);

                        System.Diagnostics.Debug.WriteLine("BETUL KE " + NRM);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(ex.ToString());
                    }
                }

                var queryCount = (from p in b.VFDITEMs
                                  where p.USERPARTIALNAME.Trim() == exc_abb.Trim() && p.DTYPE == "Plan"
                                  select p).Count();
                if (queryCount > 0)
                {
                    System.Diagnostics.Debug.WriteLine("count " + queryCount);
                   var query = (from p in b.VFDITEMs
                                 where p.USERPARTIALNAME.Trim() == exc_abb.Trim() && p.DTYPE == "Plan"
                                 select p).Single();
                    ProjectNRMID = query.ID.ToString();
                    //ProjectNRMID = "109787";
                }
                else
                {
                    var query = (from p in b.VFDITEMs
                                 where p.USERDISPLAYNAME.Trim() == exc_abb.Trim() && p.DTYPE == "Plan"
                                 select p).Single();
                    ProjectNRMID = query.ID.ToString();
                    //ProjectNRMID = "Fail";
                }
            }

            roles.Add(role);
            System.Diagnostics.Debug.WriteLine(ProjectNRMID);
            string passPhrase = "preAuthpassword";
            string url = Encryptor.GetURL(passPhrase, "10.41.61.177", "8080", username, roles);

            System.Diagnostics.Debug.WriteLine("DATA LINK : " + visionael);
            return Json(new
            {
                Success = true,
                ProjectNRMID = ProjectNRMID,
                url = url,
                visionael = visionael
            }, JsonRequestBehavior.AllowGet); //
        }

        //[HttpPost]
        public ActionResult ISP(string ispid, string username)
        {
            string role = "";
            List<string> roles = new List<string>();
            Tools tool = new Tools();

            using (Entities ctxData = new Entities())
            {
                var queryUser = (from p in ctxData.WV_USER
                                 where p.USERNAME.ToUpper() == username.ToUpper() || p.USERNAME.ToLower() == username.ToLower()
                                 select p).Single();

                var queryUserRole = (from fx in ctxData.WV_GROUP
                                     join fxx in ctxData.WV_GRP_ROLE on fx.GRPNAME equals fxx.GRPNAME
                                     where fx.GRP_ID == queryUser.GROUPID
                                     select fxx).Single();

                role = queryUserRole.GRPNAME + "," + queryUserRole.ROLENAME;
                username = queryUser.USERNAME;
            }

            roles.Add(role);

            string passPhrase = "preAuthpassword";
            string url = Encryptor.GetURL(passPhrase, "10.41.61.177", "8080", username, roles);

            return Redirect("http://10.41.61.177:8080/nrm/FacilitiesDesigner/vfd_item." + ispid + "/Graphic/ActionsTab" + url);
        }

        [HttpPost]
        public ActionResult OSPLaunchGTech(string id)
        {
            string usrPassword;
            string username = "";
            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_USER
                             where p.USERNAME.ToUpper() == User.Identity.Name.ToUpper() || p.USERNAME.ToLower() == User.Identity.Name.ToLower()
                             select p).Single();
                usrPassword = query.PASSWORD;
                username = query.USERNAME;
            }

            // to fetch from directory/database for the current user
            string passwordClear = usrPassword; // user password
            // other parameters that need to be passed
            string config = "NEPS";
            string jobID = id;
            string FIDToZoom = ""; // g3e_fid = filter FNO AND Scheme Name
            string exc_abb = "";
            string schemeType = "";

            using (Entities ctxData = new Entities())
            {
                var queryJOB = (from p in ctxData.G3E_JOB
                                where p.G3E_IDENTIFIER == id
                                select p).Single();

                string scheme_name = queryJOB.SCHEME_NAME;
                exc_abb = queryJOB.EXC_ABB;
                schemeType = queryJOB.SCHEME_TYPE;
                System.Diagnostics.Debug.WriteLine(scheme_name);

                var queryCountNetelem = (from p in ctxData.GC_NETELEM
                                         where p.SCHEME_NAME == scheme_name
                                         select p).Count();
                if (queryCountNetelem > 0)
                {
                    var query = (from p in ctxData.GC_NETELEM
                                 where p.SCHEME_NAME == scheme_name
                                 select p).Single();

                    FIDToZoom = query.G3E_FID.ToString();
                }
                else
                {
                    var Countquery = (from p in ctxData.GC_NETELEM
                                      where p.EXC_ABB.Trim() == exc_abb.Trim() && p.G3E_FNO == 6000
                                      select p).Count();
                    if (Countquery != 1)
                    {
                        FIDToZoom = "1032536";
                    }
                    else
                    {
                        var query = (from p in ctxData.GC_NETELEM
                                     where p.EXC_ABB.Trim() == exc_abb.Trim() && p.G3E_FNO == 6000
                                     select p).Single();
                        FIDToZoom = query.G3E_FID.ToString();
                    }
                }
            }
            //FIDToZoom = "1032536";
            System.Diagnostics.Debug.WriteLine(username);
            //System.Diagnostics.Debug.WriteLine(passwordClear);
            System.Diagnostics.Debug.WriteLine("FID : " + FIDToZoom);


            // encrypt the password
            string password = Crypto.EncryptStringAES(passwordClear);
            //if (schemeType == "D/Side" || schemeType == "E/Side" || schemeType == "Fiber E/Side") // , 
            //{
            //    jobID = jobID.Replace("/", "%2F");
            //}

            return Json(new
            {
                Success = true,
                username = username,
                password = password,
                config = config,
                jobID = jobID,
                FIDToZoom = FIDToZoom
            }, JsonRequestBehavior.AllowGet); //
        }

        public ActionResult UpdateData(string jobIdVal, string txtDescription, string txtDescription2, string txtJobState, string txtEndDate, string txtStartDate)
        {
            System.Diagnostics.Debug.WriteLine("UPDATE!!");
            Tools tool = new Tools();
            bool success = true;
            bool successISP = true;
            //string sqlCmd = "";
            string JobType = "";
            string schemeName = "";
            string checkGems = "";
            string checkGemsISP = "";
            string jobProposed = "";
            string G3EJobType = "";

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.Job newjob = new WebService._base.Job();
            System.Diagnostics.Debug.WriteLine("UPDATE!! :" + txtJobState);
            newjob.G3E_DESCRIPTION = txtDescription;
            newjob.G3E_DESCRIPTION_2 = txtDescription2;
            newjob.G3E_STATE = txtJobState;

            using (Entities ctxData = new Entities())
            {
                var query = (from d in ctxData.G3E_JOB
                             where d.G3E_IDENTIFIER == jobIdVal
                             select d).Single();

                var queryCariISP = (from d in ctxData.WV_ISP_JOB
                                    where d.SCHEME_NAME == jobIdVal
                                    select d).Count();

                if (queryCariISP > 0)
                {
                    var query2 = (from d in ctxData.WV_ISP_JOB
                                  where d.SCHEME_NAME == jobIdVal
                                  select d).Single();

                    schemeName = query2.G3E_IDENTIFIER;
                    checkGemsISP = query2.GEMS_TABLE;
                }
                var queryJobType = (from d in ctxData.REF_JOB_TYPES
                                    where d.JOB_TYPE == query.JOB_TYPE
                                    select d).Single();
                JobType = queryJobType.ISPOSP;
                checkGems = query.GEMS_TABLE;
                jobProposed = query.JOB_STATE;
                G3EJobType = query.JOB_TYPE;
            }

            //HAIDAR TUTUP KEJAP
            //if (txtJobState == "UN_CONSTRUCT")
            //{
            //    using (Entities ctxData = new Entities())
            //    {
            //        if (jobProposed == "PROPOSED" && G3EJobType != "HSBB (Equip)") //tukar kejap
            //        {
            //            var record = ctxData.GC_NETELEM.Where(q => q.JOB_ID == jobIdVal);
            //            if (record.Count() != 0)
            //            {
            //                var getfeature = (from p in ctxData.GC_NETELEM
            //                                  where p.JOB_ID == jobIdVal && (p.FEATURE_STATE != "PAD" && p.FEATURE_STATE != "MOD" && p.FEATURE_STATE != "PDR" && p.FEATURE_STATE != "PPF")
            //                                  select p);

            //                if (getfeature.Count() == 0)
            //                {
            //                    if (JobType == "ISPOSP")
            //                    {
            //                        successISP = myWebService.UpdateJobISP(newjob, schemeName);
            //                    }

            //                    success = myWebService.UpdateJob(newjob, jobIdVal);
            //                }
            //                else
            //                {
            //                    success = false;
            //                }
            //            }
            //            else
            //            {
            //                success = false;
            //            }
            //        }
            //        else
            //        {
            //            if (JobType == "ISPOSP")
            //            {
            //                successISP = myWebService.UpdateJobISP(newjob, schemeName);
            //            }
            //            success = myWebService.UpdateJob(newjob, jobIdVal);
            //        }

            //    }
            //}
            //else
            //{
            if (JobType == "ISPOSP")
            {
                successISP = myWebService.UpdateJobISP(newjob, schemeName, txtStartDate, txtEndDate);
            }
            success = myWebService.UpdateJob(newjob, jobIdVal);
            //}
            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateDataISP(string jobIdVal, string txtDescription, string txtDescription2, string txtJobState, string txtEndDate, string txtStartDate)
        {
            System.Diagnostics.Debug.WriteLine("UPDATE!!");
            Tools tool = new Tools();
            bool success = true;
            bool successOSP = true;
            //string sqlCmd = "";
            string JobType = "";
            string schemeName = "";
            string checkGems = "";
            string checkGemsISP = "";
            string jobProposed = "";
            string G3EJobType = "";

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.Job newjob = new WebService._base.Job();
            System.Diagnostics.Debug.WriteLine("UPDATE!! :" + txtJobState);
            newjob.G3E_DESCRIPTION = txtDescription;
            newjob.G3E_DESCRIPTION_2 = txtDescription2;
            newjob.G3E_STATE = txtJobState;
            //dapatkan job type ISP
            using (Entities ctxData = new Entities())
            {
                var query = (from d in ctxData.WV_ISP_JOB
                             where d.G3E_IDENTIFIER == jobIdVal
                             select d).Single();

                var queryCariOSP = (from d in ctxData.G3E_JOB
                                    where d.G3E_IDENTIFIER == query.SCHEME_NAME
                                    select d).Count();

                if (queryCariOSP > 0)
                {
                    var query2 = (from d in ctxData.G3E_JOB
                                  where d.G3E_IDENTIFIER == query.SCHEME_NAME
                                  select d).Single();

                    schemeName = query2.G3E_IDENTIFIER;
                    checkGems = query2.GEMS_TABLE;
                }
                var queryJobType = (from d in ctxData.REF_JOB_TYPES
                                    where d.JOB_TYPE == query.JOB_TYPE
                                    select d).Single();
                JobType = queryJobType.ISPOSP;
                checkGemsISP = query.GEMS_TABLE;
                jobProposed = query.JOB_STATE;
                G3EJobType = query.JOB_TYPE;
            }

            //Haidar tutup kejap
            //if (txtJobState == "UN_CONSTRUCT")
            //{
            //    using (Entities ctxData = new Entities())
            //    {
            //        if (JobType == "ISPOSP")
            //        {
            //            if (jobProposed == "PROPOSED" && G3EJobType != "HSBB (Equip)") //tukar kejap
            //            {
            //                var record = ctxData.GC_NETELEM.Where(q => q.JOB_ID == jobIdVal);
            //                if (record.Count() != 0)
            //                {
            //                    var getfeature = (from p in ctxData.GC_NETELEM
            //                                      where p.JOB_ID == jobIdVal && (p.FEATURE_STATE != "PAD" && p.FEATURE_STATE != "MOD" && p.FEATURE_STATE != "PDR" && p.FEATURE_STATE != "PPF")
            //                                      select p);

            //                    if (getfeature.Count() == 0)
            //                    {
            //                        successOSP = myWebService.UpdateJob(newjob, schemeName);
            //                        success = myWebService.UpdateJobISP(newjob, jobIdVal);
            //                    }
            //                }
            //                else
            //                {
            //                    success = false;
            //                }
            //            }
            //            else
            //            {
            //                successOSP = myWebService.UpdateJob(newjob, schemeName);
            //                success = myWebService.UpdateJobISP(newjob, jobIdVal);
            //            }
            //        }
            //        else
            //        {
            //            success = myWebService.UpdateJobISP(newjob, jobIdVal);
            //        }
            //    }
            //}
            //else
            //{
            if (JobType == "ISPOSP")
            {
                successOSP = myWebService.UpdateJob(newjob, schemeName);
            }
            success = myWebService.UpdateJobISP(newjob, jobIdVal, txtStartDate, txtEndDate);
            //}

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateDataApprove(string jobIdVal, string txtDescription)
        {
            System.Diagnostics.Debug.WriteLine("UPDATE!!");

            Tools tool = new Tools();
            bool success = true;
            bool successglobal = true;
            bool successg3e = true;
            string sqlCmd = "";
            string sqlCmdGlobal = "";
            string sqlCmdG3E = "";
            string sqlCmdWV = "";
            int count;
            int countApproved;
            string username = "";
            using (Entities ctxData = new Entities())
            {
                var queryUsr = (from c in ctxData.WV_USER
                                where c.USERNAME.ToUpper() == User.Identity.Name.ToUpper() || c.USERNAME.ToLower() == User.Identity.Name.ToLower()
                                select c).Single();
                username = queryUsr.USERNAME;
            }

            sqlCmd = "UPDATE WV_USER_APPROVE SET REMARKS = '" + txtDescription + "', STATUS='APPROVED' WHERE JOB_ID ='" + jobIdVal + "' AND APPROVAL_USER='" + username + "'";
            System.Diagnostics.Debug.WriteLine("SQL Statement :" + sqlCmd);
            using (Entities ctxData = new Entities())
            {
                success = tool.ExecuteSql(ctxData, sqlCmd);
            }
            System.Diagnostics.Debug.WriteLine("bool " + success);

            using (Entities ctxData = new Entities())
            {
                count = (from c in ctxData.WV_USER_APPROVE
                         where c.JOB_ID == jobIdVal
                         select c).Count();

                countApproved = (from c in ctxData.WV_USER_APPROVE
                                 where c.JOB_ID == jobIdVal && c.STATUS == "APPROVED"
                                 select c).Count();
                if (count == countApproved)
                {
                    int countosp = (from c in ctxData.G3E_JOB
                                    where c.G3E_IDENTIFIER == jobIdVal
                                    select c).Count();

                    int countisp = (from c in ctxData.WV_ISP_JOB
                                    where c.G3E_IDENTIFIER == jobIdVal
                                    select c).Count();

                    if (countosp > 0)
                    {
                        sqlCmdGlobal = "UPDATE WV_USER_APPROVE SET GLOBAL_STATUS = 'COMPLETED' WHERE JOB_ID ='" + jobIdVal + "'";
                        sqlCmdG3E = "UPDATE G3E_JOB SET JOB_STATE = 'APPROVED' WHERE G3E_IDENTIFIER ='" + jobIdVal + "'";
                        System.Diagnostics.Debug.WriteLine("SQL Statement :" + sqlCmd);
                    }
                    if (countisp > 0)
                    {
                        sqlCmdGlobal = "UPDATE WV_USER_APPROVE SET GLOBAL_STATUS = 'COMPLETED' WHERE JOB_ID ='" + jobIdVal + "'";
                        sqlCmdWV = "UPDATE WV_ISP_JOB SET JOB_STATE = 'APPROVED' WHERE G3E_IDENTIFIER ='" + jobIdVal + "'";
                        System.Diagnostics.Debug.WriteLine("SQL Statement :" + sqlCmd);
                    }
                }
            }

            using (Entities ctxData = new Entities())
            {
                successglobal = tool.ExecuteSql(ctxData, sqlCmdGlobal);
                successg3e = tool.ExecuteSql(ctxData, sqlCmdG3E);
                tool.ExecuteSql(ctxData, sqlCmdWV);
            }

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateDataReject(string jobIdVal, string txtDescription)
        {
            System.Diagnostics.Debug.WriteLine("UPDATE!!");

            Tools tool = new Tools();
            bool success = true;
            string sqlCmd = "";
            if (txtDescription != "")
            {

                sqlCmd = "UPDATE WV_USER_APPROVE SET REMARKS = '" + txtDescription + "', STATUS='REJECTED' WHERE JOB_ID ='" + jobIdVal + "'";
                string sqlCmd2 = "UPDATE G3E_JOB SET G3E_DESCRIPTION_2 = '" + txtDescription + "', JOB_STATE='REJECTED' WHERE G3E_IDENTIFIER ='" + jobIdVal + "'";
                string sqlCmd3 = "UPDATE WV_ISP_JOB SET G3E_DESCRIPTION_2 = '" + txtDescription + "', JOB_STATE='REJECTED' WHERE G3E_IDENTIFIER ='" + jobIdVal + "'";
                System.Diagnostics.Debug.WriteLine("SQL Statement :" + sqlCmd);

                using (Entities ctxData = new Entities())
                {
                    tool.ExecuteSql(ctxData, sqlCmd2);
                    tool.ExecuteSql(ctxData, sqlCmd3);
                    success = tool.ExecuteSql(ctxData, sqlCmd);

                }
            }
            else
            {
                success = false;
            }
            System.Diagnostics.Debug.WriteLine("bool " + success);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult ChangeOwner(string jobIdVal, string username)
        {
            System.Diagnostics.Debug.WriteLine("change owner!!");
            bool success = true;


            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.Job newjob = new WebService._base.Job();
            System.Diagnostics.Debug.WriteLine("UPDATE!! :" + jobIdVal);
            newjob.G3E_OWNER = username;
            success = myWebService.UpdateOwner(newjob, jobIdVal);


            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult MigratedLNMS(string lmnsjob, string nepsjob)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            System.Diagnostics.Debug.WriteLine(lmnsjob + "-!!!!!-" + nepsjob);
            success = myWebService.migratedLmns(lmnsjob, nepsjob);
            System.Diagnostics.Debug.WriteLine("-ok-");
            return Json(new
            {
                Success = true
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult DeleteData(string targetJob)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            using (Entities ctxData = new Entities())
            {
                var queryOSP = (from p in ctxData.G3E_JOB
                                where p.G3E_IDENTIFIER == targetJob
                                select p).Count();

                var queryISP = (from p in ctxData.WV_ISP_JOB
                                where p.SCHEME_NAME == targetJob
                                select p).Count();

                if (queryOSP > 0)
                {
                    success = myWebService.DeleteJob(targetJob);
                }
                if (queryISP > 0)
                {
                    success = myWebService.DeleteJobISP(targetJob);
                }
            }
            //success = myWebService.DeleteJob(targetJob);

            return Json(new
            {
                Success = true
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult DeleteDataISP(string targetJob)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            using (Entities ctxData = new Entities())
            {
                var queryISP = (from p in ctxData.WV_ISP_JOB
                                where p.G3E_IDENTIFIER == targetJob
                                select p).Count();

                var queryDATAISP = (from p in ctxData.WV_ISP_JOB
                                    where p.G3E_IDENTIFIER == targetJob
                                    select p).Single();

                var queryOSP = (from p in ctxData.G3E_JOB
                                where p.G3E_IDENTIFIER == queryDATAISP.SCHEME_NAME
                                select p).Count();

                if (queryISP > 0)
                {
                    success = myWebService.DeleteJobISP(queryDATAISP.SCHEME_NAME);
                }
                if (queryOSP > 0)
                {
                    var queryDATAOSP = (from p in ctxData.G3E_JOB
                                        where p.G3E_IDENTIFIER == queryDATAISP.SCHEME_NAME
                                        select p).Single();
                    success = myWebService.DeleteJob(queryDATAOSP.G3E_IDENTIFIER);
                }
            }
            //success = myWebService.DeleteJobISP(targetJob);

            return Json(new
            {
                Success = true
            }, JsonRequestBehavior.AllowGet);
        }


        [HttpPost]
        public ActionResult DeleteDataRedmark(string id)
        {
            bool success;

            Tools tool = new Tools();
            string sqlCmdGlobal = "DELETE FROM WV_NONNETWORK_JOB  WHERE G3E_IDENTIFIER ='" + id + "'";
            using (Entities ctxData = new Entities())
            {
                success = tool.ExecuteSql(ctxData, sqlCmdGlobal);
            }

            //return RedirectToAction("RNOList");
            return Json(new
            {
                Success = true
            }, JsonRequestBehavior.AllowGet);
        }


        [HttpPost]
        public ActionResult FindProjectNo(string projectNo, string exchange)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            DateTime MyDateTime1;
            DateTime MyDateTime2;
            string output = "";
            string output2 = "";
            if (exchange.Trim() == "BJL01" || exchange.Trim() == "BJL02" || exchange.Trim() == "BJL03" || exchange.Trim() == "BJL04" || exchange.Trim() == "BJL05" ||
                exchange.Trim() == "BJL06" || exchange.Trim() == "BJL07" || exchange.Trim() == "BJL08" || exchange.Trim() == "BJL09" || exchange.Trim() == "BJL10" ||
                exchange.Trim() == "BJL11" || exchange.Trim() == "BJL12" || exchange.Trim() == "BJL13" || exchange.Trim() == "BJL14" || exchange.Trim() == "BJL15" ||
                exchange.Trim() == "BJL16" || exchange.Trim() == "BJL17" || exchange.Trim() == "BJL18" || exchange.Trim() == "BJL19" || exchange.Trim() == "BJL20" ||
                exchange.Trim() == "BJL21" || exchange.Trim() == "BJL22" || exchange.Trim() == "BJL23" || exchange.Trim() == "BJL24" || exchange.Trim() == "BJL25" ||
                exchange.Trim() == "BJL26" || exchange.Trim() == "BJL27" || exchange.Trim() == "BJL28" || exchange.Trim() == "BJL29" || exchange.Trim() == "BJL30" ||
                exchange.Trim() == "BJL31")
            {
                exchange = "BJL";
            }
            using (Entities ctxData = new Entities())
            {
                System.Diagnostics.Debug.WriteLine("projectNo : " + projectNo);
                if (projectNo != "" && projectNo != null)
                {
                    var query = (from a in ctxData.WV_GEM_PROJNO
                                 //join fx in ctxData.G3E_JOB on a.WBS_NUM equals fx.WBS_NUM into test
                                 //from fx1 in test.DefaultIfEmpty()
                                 //join fxx in ctxData.WV_ISP_JOB on a.WBS_NUM equals fxx.WBS_NUM into test2
                                 //from fx2 in test2.DefaultIfEmpty()
                                 where //a.EXC_ABB.Trim() == exchange.Trim() && 
                                     //(fx1 == null && fx2 == null) 
                                 !ctxData.G3E_JOB.Any(r2 => a.WBS_NUM.Trim() == r2.WBS_NUM.Trim())
                                 && !ctxData.WV_ISP_JOB.Any(r2 => a.WBS_NUM.Trim() == r2.WBS_NUM.Trim())
                                 && (a.PROJECT_NO.Contains(projectNo.ToUpper()) || a.PROJECT_NO.Contains(projectNo.ToLower()) || a.PROJECT_NO.Contains(projectNo)) //&& a.EXC_ABB == exchange
                                 orderby a.PROJECT_NO, a.WBS_NUM
                                 select new
                                 {
                                     a.PROJECT_NO,
                                     a.WBS_NUM,
                                     a.WBS_DESC,
                                     a.START_DATE,
                                     a.END_DATE,
                                     a.PROJ_DESC,
                                     a.AMOUNT
                                 }).Take(100);

                    if (query.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in query)
                        {
                            counter++;
                            MyDateTime1 = new DateTime();
                            MyDateTime2 = new DateTime();

                            MyDateTime1 = DateTime.Parse(lp.START_DATE.ToString());
                            MyDateTime2 = DateTime.Parse(lp.END_DATE.ToString());

                            string carry = "";
                            output += lp.PROJECT_NO + "|" + lp.PROJ_DESC + "|" + lp.WBS_NUM + "|" + lp.WBS_DESC + "|" + MyDateTime1.ToString("dd-MM-yyyy") +
                                "|" + MyDateTime2.ToString("dd-MM-yyyy") + "|" + lp.AMOUNT;

                            output += "!";
                            //System.Diagnostics.Debug.WriteLine(output);
                            output2 = "Cari2";
                        }
                    }
                    //}
                }
                else
                {
                    var query = (from a in ctxData.WV_GEM_PROJNO
                                 //join fx in ctxData.G3E_JOB on a.WBS_NUM equals fx.WBS_NUM into test
                                 //from fx1 in test.DefaultIfEmpty()
                                 //join fxx in ctxData.WV_ISP_JOB on a.WBS_NUM equals fxx.WBS_NUM into test2
                                 //from fx2 in test2.DefaultIfEmpty()
                                 where !ctxData.G3E_JOB.Any(r2 => a.WBS_NUM.Trim() == r2.WBS_NUM.Trim())
                                 && !ctxData.WV_ISP_JOB.Any(r2 => a.WBS_NUM.Trim() == r2.WBS_NUM.Trim())
                                 orderby a.PROJECT_NO, a.WBS_NUM
                                 select new
                                 {
                                     a.PROJECT_NO,
                                     a.WBS_NUM,
                                     a.WBS_DESC,
                                     a.START_DATE,
                                     a.END_DATE,
                                     a.PROJ_DESC,
                                     a.AMOUNT
                                 }).Take(100);

                    //var DistinctItems = query.GroupBy(x => x.WBS_NUM).Select(y => y.First());

                    if (query.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in query)
                        {
                            counter++;
                            MyDateTime1 = new DateTime();
                            MyDateTime2 = new DateTime();

                            MyDateTime1 = DateTime.Parse(lp.START_DATE.ToString());
                            MyDateTime2 = DateTime.Parse(lp.END_DATE.ToString());

                            string carry = "";
                            output += lp.PROJECT_NO + "|" + lp.PROJ_DESC + "|" + lp.WBS_NUM + "|" + lp.WBS_DESC + "|" + MyDateTime1.ToString("dd-MM-yyyy") +
                                "|" + MyDateTime2.ToString("dd-MM-yyyy") + "|" + lp.AMOUNT;

                            output += "!";
                            //System.Diagnostics.Debug.WriteLine(output);
                            output2 = "Cari2";
                        }
                    }
                    //}
                }
            }
            //string res = myWebService.GetProject(projectNo);

            return Json(new
            {
                record = output,
                jenisCarian = output2,
                user = User.Identity.Name
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult UserList(string userRole)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            System.Diagnostics.Debug.WriteLine("USERROLE : " + userRole);
            string res = myWebService.GetUserList(userRole, User.Identity.Name);

            return Json(new
            {
                record = res,
                user = User.Identity.Name
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult GetListData(int id, string jobID)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            string outputcrtCardR = "";
            string outputcrtFC = "";
            string outputcrtFU = "";
            string outputcrtNE = "";
            string outputcrtLE = "";
            string outputcrtLPC = "";
            string outputcrtLPNC = "";
            string outputcrtLOSITE = "";
            string outputcrtGSB = "";
            string outputcrtDM = "";

            using (EntitiesNetworkElement ctxData = new EntitiesNetworkElement())
            {
                var queryHandover = (from a in ctxData.BI_PROCESS_ISP
                                     where a.NEPS_JOB_ID.Trim() == jobID.Trim() && a.PROC_ID == id
                                     select a).Count();

                var queryHandoverGRN = (from a in ctxData.BI_PROC_GRN_ISP
                                        where a.NEPS_JOB_ID.Trim() == jobID.Trim() && a.PROC_ID == id
                                        select a).Count();

                var queryHandoverGRNOSP = (from a in ctxData.BI_PROC_GRN_OSP
                                           where a.NEPS_JOB_ID.Trim() == jobID.Trim() && a.PROC_ID == id
                                           select a).Count();

                if (queryHandover > 0)
                {
                    var query = (from a in ctxData.BI_CREATECARD_REQUEST
                                 join fx in ctxData.BI_CREATECARD_REPLY on a.REQUEST_ID equals fx.REQUEST_ID
                                 where a.PROC_ID == id
                                 //orderby a.PROJECT_NO, a.WBS_NUM
                                 select new
                                 {
                                     a.EQUPID,
                                     a.CARDSLOT,
                                     a.CARDNAME,
                                     a.CARDMODEL,
                                     a.CARDCOUNTPORT,
                                     a.PORTSTARTNUM,
                                     a.CARDSTATUS,
                                     a.TIME_RETURNED,
                                     fx.ERRORMSG
                                 });
                    System.Diagnostics.Debug.WriteLine(query);
                    if (query.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in query.OrderBy(it => it.CARDSLOT))
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.ERRORMSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.ERRORMSG;
                            }
                            string carry = "";
                            outputcrtCardR += lp.EQUPID + "|" + lp.CARDSLOT + "|" + lp.CARDNAME + "|" + lp.CARDMODEL + "|" + lp.CARDCOUNTPORT + "|" + lp.PORTSTARTNUM + "|" + lp.CARDSTATUS + "|" + lp.TIME_RETURNED + "|" + errorMsg;

                            outputcrtCardR += "!";
                        }
                    }

                    var queryFC = (from a in ctxData.BI_CREATEFC_REQUEST
                                   join fx in ctxData.BI_CREATEFC_REPLY on a.REQUEST_ID equals fx.REQUEST_ID
                                   where a.PROC_ID == id
                                   //orderby a.PROJECT_NO, a.WBS_NUM
                                   select new
                                   {
                                       a.LOCNTTNAME,
                                       a.FRANNAME,
                                       a.INDEXX,
                                       a.LOCATIONDETAIL,
                                       a.STATUS,
                                       a.TIME_RETURNED,
                                       fx.ERRORMSG
                                   });
                    System.Diagnostics.Debug.WriteLine(queryFC);
                    if (queryFC.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryFC)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.ERRORMSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.ERRORMSG;
                            }
                            string carry = "";
                            outputcrtFC += lp.FRANNAME + "|" + lp.LOCNTTNAME + "|" + lp.INDEXX + "|" + lp.LOCATIONDETAIL + "|" + lp.STATUS + "|" + lp.TIME_RETURNED + "|" + errorMsg;

                            outputcrtFC += "!";
                        }
                    }

                    var queryFU = (from a in ctxData.BI_CREATEFU_REQUEST
                                   join fx in ctxData.BI_CREATEFU_REPLY on a.REQUEST_ID equals fx.REQUEST_ID
                                   where a.PROC_ID == id
                                   //orderby a.PROJECT_NO, a.WBS_NUM
                                   select new
                                   {
                                       a.FRAUPOSITION,
                                       a.FRAUNAME,
                                       a.FUPTMANRABBREAVIATION,
                                       a.PRODUCTTYPE,
                                       a.TERMINATIONTYPE,
                                       a.COUNTPAIR,
                                       a.TIME_RETURNED,
                                       fx.ERRORMSG
                                   });
                    System.Diagnostics.Debug.WriteLine(queryFU);
                    if (queryFU.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryFU)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.ERRORMSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.ERRORMSG;
                            }
                            string carry = "";
                            outputcrtFU += lp.FRAUNAME + "|" + lp.FRAUPOSITION + "|" + lp.FUPTMANRABBREAVIATION + "|" + lp.PRODUCTTYPE + "|" + lp.TERMINATIONTYPE + "|" + lp.COUNTPAIR + "|" + lp.TIME_RETURNED + "|" + errorMsg;

                            outputcrtFU += "!";
                        }
                    }

                    var queryNE = (from a in ctxData.BI_CREATENE_REQUEST
                                   join fx in ctxData.BI_CREATENE_REPLY on a.REQUEST_ID equals fx.REQUEST_ID
                                   where a.PROC_ID == id
                                   select new
                                   {
                                       a.EQUPLOCNTTNAME,
                                       a.EQUPEQUTABBREAVIATION,
                                       a.EQUPINDEX,
                                       a.EQUPSTATUS,
                                       a.EQUPMANRABBREAVIATION,
                                       a.EQUPEQUMMODEL,
                                       a.TIME_RETURNED,
                                       fx.ERRORMSG
                                   });
                    System.Diagnostics.Debug.WriteLine(queryNE);
                    if (queryNE.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryNE)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.ERRORMSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.ERRORMSG;
                            }
                            string carry = "";
                            outputcrtNE += lp.EQUPLOCNTTNAME + "|" + lp.EQUPEQUTABBREAVIATION + "|" + lp.EQUPINDEX + "|" + lp.EQUPEQUMMODEL + "|" + lp.EQUPMANRABBREAVIATION + "|" + lp.EQUPSTATUS + "|" + lp.TIME_RETURNED + "|" + errorMsg;

                            outputcrtNE += "!";
                        }
                    }
                }
                if (queryHandoverGRN > 0)
                {

                    var queryLE = (from a in ctxData.BI_GRNDLOADEQUIP_REQ
                                   join fx in ctxData.BI_GRNDLOADEQUIP_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                   where a.PROC_ID == id
                                   select new
                                   {
                                       a.EQUIP_EQUIPID,
                                       a.EQUIP_EQUIPCAT,
                                       a.EQUIP_EQUIPVEND,
                                       a.EQUIP_EQUIPMODEL,
                                       a.EQUIPUDA_TAGGING,
                                       a.EQUIPUDA_OUTINDOORTAG,
                                       a.TIME_SENT,
                                       fx.CALLSTATUS_MSG
                                   });
                    System.Diagnostics.Debug.WriteLine(queryLE);
                    if (queryLE.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryLE)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.CALLSTATUS_MSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.CALLSTATUS_MSG;
                            }
                            string carry = "";
                            outputcrtLE += lp.EQUIP_EQUIPID + "|" + lp.EQUIP_EQUIPCAT + "|" + lp.EQUIP_EQUIPVEND + "|" + lp.EQUIP_EQUIPMODEL + "|" + lp.EQUIPUDA_TAGGING + "|" + lp.EQUIPUDA_OUTINDOORTAG + "|" + lp.TIME_SENT + "|" + errorMsg;

                            outputcrtLE += "!";
                        }
                    }

                    var queryLPNC = (from a in ctxData.BI_GRNDLDPATHNONCONS_REQ
                                     join fx in ctxData.BI_GRNDLDPATHNONCONS_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                     where a.PROC_ID == id
                                     select new
                                     {
                                         a.ANAME,
                                         a.ASITE,
                                         a.ATYPE,
                                         a.ASLOT,
                                         a.ACARD,
                                         a.APORT,
                                         a.ZNAME,
                                         a.ZSITE,
                                         a.ZTYPE,
                                         a.ZSLOT,
                                         a.ZCARD,
                                         a.ZPORT,
                                         a.PRIMARYSECONDARY,
                                         a.PATHBW,
                                         a.TIME_SENT,
                                         fx.CALLSTATUS_MSG
                                     });
                    System.Diagnostics.Debug.WriteLine(queryLPNC);
                    if (queryLPNC.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryLPNC)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.CALLSTATUS_MSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.CALLSTATUS_MSG;
                            }
                            string carry = "";
                            outputcrtLPNC += lp.ANAME + "|" + lp.ASITE + "|" + lp.ATYPE + "|" + lp.ASLOT + "|";
                            outputcrtLPNC += lp.ACARD + "|" + lp.APORT + "|" + lp.ZNAME + "|" + lp.ZSITE + "|";
                            outputcrtLPNC += lp.ZTYPE + "|" + lp.ZSLOT + "|" + lp.ZCARD + "|" + lp.ZPORT + "|";
                            outputcrtLPNC += lp.PRIMARYSECONDARY + "|" + lp.PATHBW + "|" + lp.TIME_SENT + "|" + errorMsg;

                            outputcrtLPNC += "!";
                        }
                    }

                    var queryLPC = (from a in ctxData.BI_GRNDLDPATHCONS_REQ
                                    join fx in ctxData.BI_GRNDLDPATHCONS_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                    where a.PROC_ID == id
                                    select new
                                    {
                                        a.ANAME,
                                        a.ASITE,
                                        a.ATYPE,
                                        a.ACARD2,
                                        a.APORT2,
                                        a.ACARD3,
                                        a.APORT3,
                                        a.ZNAME,
                                        a.ZSITE,
                                        a.ZTYPE,
                                        a.ZCARD,
                                        a.ZPORT,
                                        a.DPNAME,
                                        a.DPSITE,
                                        a.TIME_SENT,
                                        fx.CALLSTATUS_MSG
                                    });
                    System.Diagnostics.Debug.WriteLine(queryLPC);
                    if (queryLPC.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryLPC)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.CALLSTATUS_MSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.CALLSTATUS_MSG;
                            }
                            string carry = "";
                            outputcrtLPC += lp.ANAME + "|" + lp.ASITE + "|" + lp.ATYPE + "|" + lp.ACARD2 + "|";
                            outputcrtLPC += lp.APORT2 + "|" + lp.ACARD3 + "|" + lp.APORT3 + "|" + lp.ZNAME + "|";
                            outputcrtLPC += lp.ZSITE + "|" + lp.ZTYPE + "|" + lp.ZCARD + "|" + lp.ZPORT + "|";
                            outputcrtLPC += lp.DPNAME + "|" + lp.DPSITE + "|" + lp.TIME_SENT + "|" + errorMsg;

                            outputcrtLPC += "!";
                        }
                    }

                    var queryDM = (from a in ctxData.BI_GRNDATAMATCH_REQ
                                   join fx in ctxData.BI_GRNDATAMATCH_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                   where a.PROC_ID == id
                                   select new
                                   {
                                       a.EVENTNAME,
                                       a.EQUIP_EQUIPID,
                                       a.EQUIP_EXCHDESC,
                                       a.EQUIP_SITE,
                                       a.TIME_SENT,
                                       fx.CALLSTATUS_MSG
                                   });
                    System.Diagnostics.Debug.WriteLine(queryDM);
                    if (queryDM.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryDM)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.CALLSTATUS_MSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.CALLSTATUS_MSG;
                            }
                            string carry = "";
                            outputcrtDM += lp.EVENTNAME + "|" + lp.EQUIP_EQUIPID + "|" + lp.EQUIP_EXCHDESC + "|" + lp.EQUIP_SITE + "|";
                            outputcrtDM += lp.TIME_SENT + "|" + errorMsg;

                            outputcrtDM += "!";
                        }
                    }
                }
                if (queryHandoverGRNOSP > 0) // granite OSP
                {
                    var queryLE = (from a in ctxData.BI_GRNOSPDLOADEQUIP_REQ
                                   join fx in ctxData.BI_GRNOSPDLOADEQUIP_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                   where a.PROC_ID == id
                                   select new
                                   {
                                       a.EQUIP_EQUIPID,
                                       a.EQUIP_EQUIPCAT,
                                       a.EQUIP_EQUIPVEND,
                                       a.EQUIP_EQUIPMODEL,
                                       a.EQUIPUDA_TAGGING,
                                       a.EQUIPUDA_OUTINDOORTAG,
                                       a.TIME_SENT,
                                       fx.CALLSTATUS_MSG
                                   });
                    System.Diagnostics.Debug.WriteLine(queryLE);
                    if (queryLE.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryLE)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.CALLSTATUS_MSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.CALLSTATUS_MSG;
                            }
                            string carry = "";
                            outputcrtLE += lp.EQUIP_EQUIPID + "|" + lp.EQUIP_EQUIPCAT + "|" + lp.EQUIP_EQUIPVEND + "|" + lp.EQUIP_EQUIPMODEL + "|" + lp.EQUIPUDA_TAGGING + "|" + lp.EQUIPUDA_OUTINDOORTAG + "|" + lp.TIME_SENT + "|" + errorMsg;

                            outputcrtLE += "!";
                        }
                    }

                    var queryLPNC = (from a in ctxData.BI_GRNOSPDLOADPATH_REQ
                                     join fx in ctxData.BI_GRNOSPDLDPATHCONS_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                     where a.PROC_ID == id
                                     select new
                                     {
                                         a.EQUIPA_NAME,
                                         a.EQUIPA_SITE,
                                         a.EQUIPA_TYPE,
                                         a.EQUIPA_SLOT,
                                         a.EQUIPA_CARD,
                                         a.EQUIPA_PORT,
                                         a.EQUIPZ_NAME,
                                         a.EQUIPZ_SITE,
                                         a.EQUIPZ_TYPE,
                                         a.EQUIPZ_SLOT,
                                         a.EQUIPZ_CARD,
                                         a.EQUIPZ_PORT,
                                         a.PATH_TYPE,
                                         a.PATH_BANDWIDTH,
                                         a.TIME_SENT,
                                         fx.CALLSTATUS_MSG
                                     });
                    System.Diagnostics.Debug.WriteLine(queryLPNC);
                    if (queryLPNC.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryLPNC)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.CALLSTATUS_MSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.CALLSTATUS_MSG;
                            }
                            string carry = "";
                            outputcrtLPNC += lp.EQUIPA_NAME + "|" + lp.EQUIPA_SITE + "|" + lp.EQUIPA_TYPE + "|" + lp.EQUIPA_SLOT + "|";
                            outputcrtLPNC += lp.EQUIPA_CARD + "|" + lp.EQUIPA_PORT + "|" + lp.EQUIPZ_NAME + "|" + lp.EQUIPZ_SITE + "|";
                            outputcrtLPNC += lp.EQUIPZ_TYPE + "|" + lp.EQUIPZ_SLOT + "|" + lp.EQUIPZ_CARD + "|" + lp.EQUIPZ_PORT + "|";
                            outputcrtLPNC += lp.PATH_TYPE + "|" + lp.PATH_BANDWIDTH + "|" + lp.TIME_SENT + "|" + errorMsg;

                            outputcrtLPNC += "!";
                        }
                    }

                    var queryLPC = (from a in ctxData.BI_GRNOSPDLDPATHCONS_REQ
                                    join fx in ctxData.BI_GRNOSPDLDPATHCONS_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                    where a.PROC_ID == id
                                    select new
                                    {
                                        a.ANAME,
                                        a.ASITE,
                                        a.ATYPE,
                                        a.ACARD2,
                                        a.APORT2,
                                        a.ACARD3,
                                        a.APORT3,
                                        a.ZNAME,
                                        a.ZSITE,
                                        a.ZTYPE,
                                        a.ZCARD,
                                        a.ZPORT,
                                        a.DPNAME,
                                        a.DPSITE,
                                        a.ACARD4,
                                        a.APORT4,
                                        a.TIME_SENT,
                                        fx.CALLSTATUS_MSG
                                    });
                    System.Diagnostics.Debug.WriteLine(queryLPNC);
                    if (queryLPC.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryLPC)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.CALLSTATUS_MSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.CALLSTATUS_MSG;
                            }
                            string carry = "";
                            outputcrtLPC += lp.ANAME + "|" + lp.ASITE + "|" + lp.ATYPE + "|" + lp.ACARD2 + "|";
                            outputcrtLPC += lp.APORT2 + "|" + lp.ACARD3 + "|" + lp.APORT3 + "|" + lp.ZNAME + "|";
                            outputcrtLPC += lp.ZSITE + "|" + lp.ZTYPE + "|" + lp.ZCARD + "|" + lp.ZPORT + "|";
                            outputcrtLPC += "-|-|";
                            outputcrtLPC += lp.DPNAME + "|" + lp.DPSITE + "|" + lp.TIME_SENT + "|" + errorMsg;

                            outputcrtLPC += "!";
                        }
                    }

                    var queryLOSITE = (from a in ctxData.BI_GRNOSPLOADSITE_REQ
                                       join fx in ctxData.BI_GRNOSPLOADSITE_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                       where a.PROC_ID == id
                                       select new
                                       {
                                           a.SITENAME,
                                           a.SITEDESC,
                                           a.ADDRSTREETTYPE,
                                           a.ADDRSTREET,
                                           a.ADDRCOUNTY,
                                           a.ADDRCITY,
                                           a.ADDRSTATE,
                                           a.ADDRPOSTCODE,
                                           a.ADDRCOUNTRY,
                                           a.UDAEQUIPLOC,
                                           a.UDACABLINGTYPE,
                                           a.UDAFIBERTOPREMISEEXIST,
                                           a.UDAACCESSRESTRICT,
                                           a.UDACONTACT,
                                           a.TIME_SENT,
                                           fx.CSMSG
                                       });
                    System.Diagnostics.Debug.WriteLine(queryLOSITE);
                    if (queryLOSITE.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryLOSITE)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.CSMSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.CSMSG;
                            }
                            string carry = "";
                            outputcrtLOSITE += lp.SITENAME + "|" + lp.SITEDESC + "|" + lp.ADDRSTREETTYPE + "," + lp.ADDRSTREET + ",";
                            outputcrtLOSITE += lp.ADDRCOUNTY + "," + lp.ADDRCITY + "," + lp.ADDRSTATE + "," + lp.ADDRPOSTCODE + ",";
                            outputcrtLOSITE += lp.ADDRCOUNTRY + "|" + lp.UDAEQUIPLOC + "|" + lp.UDACABLINGTYPE + "|" + lp.UDAFIBERTOPREMISEEXIST + "|";
                            outputcrtLOSITE += lp.UDAACCESSRESTRICT + "|" + lp.UDACONTACT + "|" + lp.TIME_SENT + "|" + errorMsg;

                            outputcrtLOSITE += "!";
                        }
                    }

                    var queryGSB = (from a in ctxData.BI_GRNOSPLOADSRVBND_REQ
                                    join fx in ctxData.BI_GRNOSPLOADSRVBND_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                    where a.PROC_ID == id
                                    //orderby a.PROJECT_NO, a.WBS_NUM
                                    select new
                                    {
                                        a.SRVBNDID,
                                        a.ADDRSITENO,
                                        a.ADDRSITEFLOOR,
                                        a.ADDRSITEBUILD,
                                        a.ADDRSTREETTYPE,
                                        a.ADDRSTREETNAME,
                                        a.ADDRSECTION,
                                        a.ADDRPOSTCODE,
                                        a.ADDRCITY,
                                        a.ADDRSTATE,
                                        a.ADDRCOUNTRY,
                                        a.ADDRPREMISETYPE,
                                        a.ADDRSITESDP,
                                        a.ADDRSITEDPEXC,
                                        fx.CALLMSG
                                    });
                    System.Diagnostics.Debug.WriteLine(queryGSB);
                    if (queryGSB.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryGSB)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.CALLMSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.CALLMSG;
                            }
                            string carry = "";
                            outputcrtGSB += lp.SRVBNDID + "|" + lp.ADDRSITENO + " " + lp.ADDRSITEFLOOR + " " + lp.ADDRSITEBUILD + "," + lp.ADDRSTREETTYPE + "," + lp.ADDRSTREETNAME + "," + lp.ADDRSECTION + ",";
                            outputcrtGSB += lp.ADDRPOSTCODE + "," + lp.ADDRCITY + "," + lp.ADDRSTATE + "," + lp.ADDRCOUNTRY + "|" + lp.ADDRPREMISETYPE + "|" + lp.ADDRSITESDP + "|" + lp.ADDRSITEDPEXC + "|" + errorMsg;

                            outputcrtGSB += "!";
                        }
                    }

                    var queryDM = (from a in ctxData.BI_GRNOSPDATAMATCH_REQ
                                   join fx in ctxData.BI_GRNOSPDATAMATCH_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                   where a.PROC_ID == id
                                   select new
                                   {
                                       a.EVENTNAME,
                                       a.EQUIP_EQUIPID,
                                       a.EQUIP_EXCHDESC,
                                       a.EQUIP_SITE,
                                       a.TIME_SENT,
                                       fx.CALLSTATUS_MSG
                                   });
                    System.Diagnostics.Debug.WriteLine(queryDM);
                    if (queryDM.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryDM)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.CALLSTATUS_MSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.CALLSTATUS_MSG;
                            }
                            string carry = "";
                            outputcrtDM += lp.EVENTNAME + "|" + lp.EQUIP_EQUIPID + "|" + lp.EQUIP_EXCHDESC + "|" + lp.EQUIP_SITE + "|";
                            outputcrtDM += lp.TIME_SENT + "|" + errorMsg;

                            outputcrtDM += "!";
                        }
                    }
                }
            }

            return Json(new
            {
                crtCardR = outputcrtCardR,
                crtFCR = outputcrtFC,
                crtFUR = outputcrtFU,
                crtNE = outputcrtNE,
                crtLE = outputcrtLE,
                crtLPNC = outputcrtLPNC,
                crtLPC = outputcrtLPC,
                crtLOSITE = outputcrtLOSITE,
                crtGSB = outputcrtGSB,
                outputcrtDM = outputcrtDM
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult GetListDataOSP(int id, string jobID)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            string outputcrtCardR = "";
            string outputcrtFC = "";
            string outputcrtFU = "";
            string outputcrtNE = "";
            string outputcrtCS = "";
            string outputcrtSB = "";
            string outputcrtLE = "";
            string outputcrtLPNC = "";
            string outputcrtLPC = "";
            string outputcrtLOSITE = "";
            string outputcrtGSB = "";
            string outputcrtDM = "";
            // region Mubin - CR58-20180330
            string outputAddCard = "";
            string outputDelEquip = "";
            string outputDelPath = "";
            string outputDelSB = "";
            // endRegion
            string outputswift = "";
            string outputmh = "";
            string outputductpath = "";
            string outputduct = "";
            string outputsubduct = "";
            using (EntitiesNetworkElement ctxData = new EntitiesNetworkElement())
            {
                var queryHandover = (from a in ctxData.BI_PROCESS
                                     where a.NEPS_JOB_ID.Trim() == jobID.Trim() && a.PROC_ID == id
                                     select a).Count();

                var queryHandoverGRN = (from a in ctxData.BI_PROC_GRN_ISP
                                        where a.NEPS_JOB_ID.Trim() == jobID.Trim() && a.PROC_ID == id
                                        select a).Count();

                var queryHandoverGRNOSP = (from a in ctxData.BI_PROC_GRN_OSP
                                           where a.NEPS_JOB_ID.Trim() == jobID.Trim() && a.PROC_ID == id
                                           select a).Count();

                if (queryHandover > 0)
                {

                    var queryNE = (from a in ctxData.BI_CREATENE_REQUEST_OSP
                                   join fx in ctxData.BI_CREATENE_REPLY_OSP on a.REQUEST_ID equals fx.REQUEST_ID
                                   where a.PROC_ID == id
                                   select new
                                   {
                                       a.EQUPLOCNTTNAME,
                                       a.EQUPEQUTABBREAVIATION,
                                       a.EQUPINDEX,
                                       a.EQUPSTATUS,
                                       a.EQUPMANRABBREAVIATION,
                                       a.EQUPEQUMMODEL,
                                       a.TIME_RETURNED,
                                       fx.ERRORMSG
                                   });
                    System.Diagnostics.Debug.WriteLine(queryNE);
                    if (queryNE.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryNE)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.ERRORMSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.ERRORMSG;
                            }
                            string carry = "";
                            outputcrtNE += lp.EQUPLOCNTTNAME + "|" + lp.EQUPEQUTABBREAVIATION + "|" + lp.EQUPINDEX + "|" + lp.EQUPEQUMMODEL + "|" + lp.EQUPMANRABBREAVIATION + "|" + lp.EQUPSTATUS + "|" + lp.TIME_RETURNED + "|" + errorMsg;

                            outputcrtNE += "!";
                        }
                    }

                    var queryNE_T = (from a in ctxData.BI_CHKMSAN_REQUEST_OSP
                                     join fx in ctxData.BI_CHKMSAN_REPLY_OSP on a.REQUEST_ID equals fx.REQUEST_ID
                                     where a.PROC_ID == id
                                     select new
                                     {
                                         a.EQUPLOCNTNAME,
                                         a.EQUPEQUTABBR,
                                         a.EQUPINDEX,
                                         a.TIME_RETURNED,
                                         fx.ERRORMSG
                                     });
                    System.Diagnostics.Debug.WriteLine(queryNE_T);
                    if (queryNE_T.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryNE_T)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.ERRORMSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.ERRORMSG;
                            }
                            string carry = "";
                            outputcrtNE += lp.EQUPLOCNTNAME + "|" + lp.EQUPEQUTABBR + "|" + lp.EQUPINDEX + "|-|-|-|" + lp.TIME_RETURNED + "|" + errorMsg;

                            outputcrtNE += "!";
                        }
                    }

                    var query = (from a in ctxData.BI_CREATECARD_REQUEST_OSP
                                 join fx in ctxData.BI_CREATECARD_REPLY_OSP on a.REQUEST_ID equals fx.REQUEST_ID
                                 where a.PROC_ID == id
                                 //orderby a.PROJECT_NO, a.WBS_NUM
                                 select new
                                 {
                                     a.EQUPID,
                                     a.CARDSLOT,
                                     a.CARDNAME,
                                     a.CARDMODEL,
                                     a.CARDCOUNTPORT,
                                     a.PORTSTARTNUM,
                                     a.CARDSTATUS,
                                     a.TIME_RETURNED,
                                     fx.ERRORMSG
                                 });
                    System.Diagnostics.Debug.WriteLine(query);
                    if (query.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in query)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.ERRORMSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.ERRORMSG;
                            }
                            string carry = "";
                            outputcrtCardR += lp.EQUPID + "|" + lp.CARDSLOT + "|" + lp.CARDNAME + "|" + lp.CARDMODEL + "|" + lp.CARDCOUNTPORT + "|" + lp.PORTSTARTNUM + "|" + lp.CARDSTATUS + "|" + lp.TIME_RETURNED + "|" + errorMsg;

                            outputcrtCardR += "!";
                        }
                    }

                    var queryFC = (from a in ctxData.BI_CREATEFC_REQUEST_OSP
                                   join fx in ctxData.BI_CREATEFC_REPLY_OSP on a.REQUEST_ID equals fx.REQUEST_ID
                                   where a.PROC_ID == id
                                   //orderby a.PROJECT_NO, a.WBS_NUM
                                   select new
                                   {
                                       a.LOCNTTNAME,
                                       a.FRANNAME,
                                       a.INDEXX,
                                       a.LOCATIONDETAIL,
                                       a.STATUS,
                                       a.TIME_RETURNED,
                                       fx.ERRORMSG
                                   });
                    System.Diagnostics.Debug.WriteLine(queryFC);
                    if (queryFC.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryFC)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.ERRORMSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.ERRORMSG;
                            }
                            string carry = "";
                            outputcrtFC += lp.FRANNAME + "|" + lp.LOCNTTNAME + "|" + lp.INDEXX + "|" + lp.LOCATIONDETAIL + "|" + lp.STATUS + "|" + lp.TIME_RETURNED + "|" + errorMsg;

                            outputcrtFC += "!";
                        }
                    }

                    var queryFC_T = (from a in ctxData.BI_UPDATEDPDIST_REQUEST_OSP
                                     join fx in ctxData.BI_UPDATEDPDIST_REPLY_OSP on a.REQUEST_ID equals fx.REQUEST_ID
                                     where a.PROC_ID == id
                                     //orderby a.PROJECT_NO, a.WBS_NUM
                                     select new
                                     {
                                         a.EQUPLOCNTNAME,
                                         a.EQUPEQUTABBR,
                                         a.EQUPINDEX,
                                         a.DISTANCE,
                                         a.TIME_RETURNED,
                                         fx.ERRORMSG
                                     });
                    System.Diagnostics.Debug.WriteLine(queryFC_T);
                    if (queryFC_T.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryFC_T)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.ERRORMSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.ERRORMSG;
                            }
                            string carry = "";
                            outputcrtFC += lp.EQUPEQUTABBR + "|" + lp.EQUPLOCNTNAME + "|" + lp.EQUPINDEX + "|" + lp.DISTANCE + "|-|" + lp.TIME_RETURNED + "|" + errorMsg;

                            outputcrtFC += "!";
                        }
                    }

                    var queryFU = (from a in ctxData.BI_CREATEFU_REQUEST_OSP
                                   join fx in ctxData.BI_CREATEFU_REPLY_OSP on a.REQUEST_ID equals fx.REQUEST_ID
                                   where a.PROC_ID == id
                                   //orderby a.PROJECT_NO, a.WBS_NUM
                                   select new
                                   {
                                       a.FRAUPOSITION,
                                       a.FRAUNAME,
                                       a.FUPTMANRABBREAVIATION,
                                       a.PRODUCTTYPE,
                                       a.TERMINATIONTYPE,
                                       a.COUNTPAIR,
                                       a.TIME_RETURNED,
                                       fx.ERRORMSG
                                   });
                    System.Diagnostics.Debug.WriteLine(queryFU);
                    if (queryFU.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryFU)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.ERRORMSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.ERRORMSG;
                            }
                            string carry = "";
                            outputcrtFU += lp.FRAUNAME + "|" + lp.FRAUPOSITION + "|" + lp.FUPTMANRABBREAVIATION + "|" + lp.PRODUCTTYPE + "|" + lp.TERMINATIONTYPE + "|" + lp.COUNTPAIR + "|" + lp.TIME_RETURNED + "|" + errorMsg;

                            outputcrtFU += "!";
                        }
                    }

                    var queryCS = (from a in ctxData.BI_CREATECS_REQUEST_OSP
                                   join fx in ctxData.BI_CREATECS_REPLY_OSP on a.REQUEST_ID equals fx.REQUEST_ID
                                   where a.PROC_ID == id
                                   //orderby a.PROJECT_NO, a.WBS_NUM
                                   select new
                                   {
                                       //a.EQUPID,
                                       a.CABS_NAME,
                                       a.CABS_CAST_TYPE,
                                       a.CSPT_PROD_NAME,
                                       a.MIN_NUMBER,
                                       a.MAX_NUMBER,
                                       a.LOCATION_A,
                                       a.LOCATION_B,
                                       fx.ERRORMSG
                                   });
                    System.Diagnostics.Debug.WriteLine(queryCS);
                    if (queryCS.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryCS)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.ERRORMSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.ERRORMSG;
                            }
                            string carry = "";
                            outputcrtCS += lp.CABS_NAME + "|" + lp.CABS_CAST_TYPE + "|" + lp.CSPT_PROD_NAME + "|" + lp.MIN_NUMBER + "|" + lp.MAX_NUMBER + "|" + lp.LOCATION_A + "|" + lp.LOCATION_B + "|" + errorMsg;
                            //outputcrtCS += lp.EQUPID + "|" + lp.CABS_NAME + "|" + lp.CABS_CAST_TYPE + "|" + lp.CSPT_PROD_NAME + "|" + lp.MIN_NUMBER + "|" + lp.MAX_NUMBER + "|" + lp.LOCATION_A + "|" + lp.LOCATION_B + "|" + errorMsg;
                            outputcrtCS += "!";
                        }
                    }

                    var querySB = (from a in ctxData.BI_CREATESB_REQUEST_OSP
                                   join fx in ctxData.BI_CREATESB_REPLY_OSP on a.REQUEST_ID equals fx.REQUEST_ID
                                   where a.PROC_ID == id
                                   //orderby a.PROJECT_NO, a.WBS_NUM
                                   select new
                                   {
                                       a.FRAC_ID,
                                       a.FIELD1,
                                       a.FIELD2,
                                       a.STREET_NUMBER,
                                       a.STREET_TYPE,
                                       a.STREET_NAME,
                                       a.SUBURB,
                                       a.STAT_ABBR,
                                       fx.ERRORMSG
                                   });
                    System.Diagnostics.Debug.WriteLine(querySB);
                    if (querySB.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in querySB)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.ERRORMSG == null)
                            {
                                errorMsg = "OK";
                            }
                            else
                            {
                                errorMsg = lp.ERRORMSG;
                            }
                            string carry = "";
                            outputcrtSB += lp.FRAC_ID + "|" + lp.FIELD1 + "|" + lp.FIELD2 + "|" + lp.STREET_NUMBER + "|" + lp.STREET_TYPE + "|" + lp.STREET_NAME + "|" + lp.SUBURB + "|" + lp.STAT_ABBR + "|" + errorMsg;

                            outputcrtSB += "!";
                        }
                    }
                }
                if (queryHandoverGRN > 0) // granite ISP
                {
                    var queryLE = (from a in ctxData.BI_GRNDLOADEQUIP_REQ
                                   join fx in ctxData.BI_GRNDLOADEQUIP_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                   where a.PROC_ID == id
                                   select new
                                   {
                                       a.EQUIP_EQUIPID,
                                       a.EQUIP_EQUIPCAT,
                                       a.EQUIP_EQUIPVEND,
                                       a.EQUIP_EQUIPMODEL,
                                       a.EQUIPUDA_TAGGING,
                                       a.EQUIPUDA_OUTINDOORTAG,
                                       a.TIME_SENT,
                                       fx.CALLSTATUS_MSG
                                   });
                    System.Diagnostics.Debug.WriteLine(queryLE);
                    if (queryLE.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryLE)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.CALLSTATUS_MSG == null)
                            {
                                errorMsg = "";
                            }
                            else
                            {
                                errorMsg = lp.CALLSTATUS_MSG;
                            }
                            string carry = "";
                            outputcrtLE += lp.EQUIP_EQUIPID + "|" + lp.EQUIP_EQUIPCAT + "|" + lp.EQUIP_EQUIPVEND + "|" + lp.EQUIP_EQUIPMODEL + "|" + lp.EQUIPUDA_TAGGING + "|" + lp.EQUIPUDA_OUTINDOORTAG + "|" + lp.TIME_SENT + "|" + errorMsg;

                            outputcrtLE += "!";
                        }
                    }

                    var queryLPNC = (from a in ctxData.BI_GRNDLDPATHNONCONS_REQ
                                     join fx in ctxData.BI_GRNDLDPATHNONCONS_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                     where a.PROC_ID == id
                                     select new
                                     {
                                         a.ANAME,
                                         a.ASITE,
                                         a.ATYPE,
                                         a.ASLOT,
                                         a.ACARD,
                                         a.APORT,
                                         a.ZNAME,
                                         a.ZSITE,
                                         a.ZTYPE,
                                         a.ZSLOT,
                                         a.ZCARD,
                                         a.ZPORT,
                                         a.PRIMARYSECONDARY,
                                         a.PATHBW,
                                         a.TIME_SENT,
                                         fx.CALLSTATUS_MSG
                                     });
                    System.Diagnostics.Debug.WriteLine(queryLPNC);
                    if (queryLPNC.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryLPNC)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.CALLSTATUS_MSG == null)
                            {
                                errorMsg = "";
                            }
                            else
                            {
                                errorMsg = lp.CALLSTATUS_MSG;
                            }
                            string carry = "";
                            outputcrtLPNC += lp.ANAME + "|" + lp.ASITE + "|" + lp.ATYPE + "|" + lp.ASLOT + "|";
                            outputcrtLPNC += lp.ACARD + "|" + lp.APORT + "|" + lp.ZNAME + "|" + lp.ZSITE + "|";
                            outputcrtLPNC += lp.ZTYPE + "|" + lp.ZSLOT + "|" + lp.ZCARD + "|" + lp.ZPORT + "|";
                            outputcrtLPNC += lp.PRIMARYSECONDARY + "|" + lp.PATHBW + "|" + lp.TIME_SENT + "|" + errorMsg;

                            outputcrtLPNC += "!";
                        }
                    }

                    var queryLPC = (from a in ctxData.BI_GRNDLDPATHCONS_REQ
                                    join fx in ctxData.BI_GRNDLDPATHCONS_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                    where a.PROC_ID == id
                                    select new
                                    {
                                        a.ANAME,
                                        a.ASITE,
                                        a.ATYPE,
                                        a.ACARD2,
                                        a.APORT2,
                                        a.ACARD3,
                                        a.APORT3,
                                        a.ZNAME,
                                        a.ZSITE,
                                        a.ZTYPE,
                                        a.ZCARD,
                                        a.ZPORT,
                                        a.DPNAME,
                                        a.DPSITE,
                                        a.TIME_SENT,
                                        fx.CALLSTATUS_MSG
                                    });
                    System.Diagnostics.Debug.WriteLine(queryLPC);
                    if (queryLPC.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryLPC)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.CALLSTATUS_MSG == null)
                            {
                                errorMsg = "";
                            }
                            else
                            {
                                errorMsg = lp.CALLSTATUS_MSG;
                            }
                            string carry = "";
                            outputcrtLPC += lp.ANAME + "|" + lp.ASITE + "|" + lp.ATYPE + "|" + lp.ACARD2 + "|";
                            outputcrtLPC += lp.APORT2 + "|" + lp.ACARD3 + "|" + lp.APORT3 + "|" + lp.ZNAME + "|";
                            outputcrtLPC += lp.ZSITE + "|" + lp.ZTYPE + "|" + lp.ZCARD + "|" + lp.ZPORT + "|";
                            outputcrtLPC += lp.DPNAME + "|" + lp.DPSITE + "|" + lp.TIME_SENT + "|" + errorMsg;

                            outputcrtLPC += "!";
                        }
                    }

                    var queryDM = (from a in ctxData.BI_GRNDATAMATCH_REQ
                                   join fx in ctxData.BI_GRNDATAMATCH_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                   where a.PROC_ID == id
                                   select new
                                   {
                                       a.EVENTNAME,
                                       a.EQUIP_EQUIPID,
                                       a.EQUIP_EXCHDESC,
                                       a.EQUIP_SITE,
                                       a.TIME_SENT,
                                       fx.CALLSTATUS_MSG
                                   });
                    System.Diagnostics.Debug.WriteLine(queryDM);
                    if (queryDM.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryDM)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.CALLSTATUS_MSG == null)
                            {
                                errorMsg = "";
                            }
                            else
                            {
                                errorMsg = lp.CALLSTATUS_MSG;
                            }
                            string carry = "";
                            outputcrtDM += lp.EVENTNAME + "|" + lp.EQUIP_EQUIPID + "|" + lp.EQUIP_EXCHDESC + "|" + lp.EQUIP_SITE + "|";
                            outputcrtDM += lp.TIME_SENT + "|" + errorMsg;

                            outputcrtDM += "!";
                        }
                    }
                }
                if (queryHandoverGRNOSP > 0) // granite OSP
                {
                    // region Mubin - CR14-20180330
                    var queryLE = (from a in ctxData.BI_GRNOSPDLOADEQUIP_REQ
                                   join fx in ctxData.BI_GRNOSPDLOADEQUIP_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                   where a.PROC_ID == id
                                   select new
                                   {
                                       a.EQUIP_EQUIPID,
                                       a.EQUIP_EQUIPCAT,
                                       a.EQUIP_EQUIPVEND,
                                       a.EQUIP_EQUIPMODEL,
                                       a.EQUIP_TEMPNAME,
                                       a.EQUIPUDA_TAGGING,
                                       a.EQUIPUDA_OUTINDOORTAG,
                                       a.PATCHCORD,
                                       a.TIME_SENT,
                                       fx.CALLSTATUS_MSG
                                   });
                    System.Diagnostics.Debug.WriteLine(queryLE);
                    if (queryLE.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryLE)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.CALLSTATUS_MSG == null)
                            {
                                errorMsg = "";
                            }
                            else
                            {
                                errorMsg = lp.CALLSTATUS_MSG;
                            }
                            string carry = "";
                            outputcrtLE += lp.EQUIP_EQUIPID + "|" + lp.EQUIP_EQUIPCAT + "|" + lp.EQUIP_EQUIPVEND + "|" + lp.EQUIP_EQUIPMODEL + "|" + lp.EQUIP_TEMPNAME + "|" + lp.EQUIPUDA_TAGGING + "|" + lp.EQUIPUDA_OUTINDOORTAG + "|" + lp.PATCHCORD + "|" + lp.TIME_SENT + "|" + errorMsg;

                            outputcrtLE += "!";
                        }
                    }
                    System.Diagnostics.Debug.WriteLine("LOAD EQUIP: " + outputcrtLE);
                    // endRegion

                    // region Mubin - CR58-20180330
                    var queryADDCARD = (from a in ctxData.BI_GRNOSPADDCARD_REQ
                                        join fx in ctxData.BI_GRNOSPADDCARD_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                        where a.PROC_ID == id
                                        select new
                                        {
                                            a.EQUIPMENTID,
                                            a.EQUIPVEND,
                                            a.EQUIPMODEL,
                                            a.TEMPLATENAME,
                                            a.SLOT_NO,
                                            a.CARD_TYPE,
                                            a.TOTAL_PORT,
                                            a.TIME_SENT,
                                            fx.MESSAGE
                                        });
                    System.Diagnostics.Debug.WriteLine(queryADDCARD);
                    if (queryADDCARD.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryADDCARD)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.MESSAGE == null)
                            {
                                errorMsg = "";
                            }
                            else
                            {
                                errorMsg = lp.MESSAGE;
                            }
                            string carry = "";
                            outputAddCard += lp.EQUIPMENTID + "|" + lp.EQUIPVEND + "|" + lp.EQUIPMODEL + "|" + lp.TEMPLATENAME + "|" + lp.SLOT_NO + "|" + lp.CARD_TYPE + "|" + lp.TOTAL_PORT + "|" + lp.TIME_SENT + "|" + errorMsg;

                            outputAddCard += "!";
                        }
                    }
                    System.Diagnostics.Debug.WriteLine("ADD CARD: " + outputAddCard);

                    var queryDelEquip = (from a in ctxData.BI_GRNOSPDELETEEQUIP_REQ
                                         join fx in ctxData.BI_GRNOSPDELETEEQUIP_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                         where a.PROC_ID == id
                                         select new
                                         {
                                             a.USERID,
                                             a.EQUIPMENTID,
                                             a.EQUIPMENTCATEGORY,
                                             a.SLOTNUMBER,
                                             a.TIME_SENT,
                                             fx.MESSAGE
                                         });
                    System.Diagnostics.Debug.WriteLine(queryDelEquip);
                    if (queryDelEquip.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryDelEquip)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.MESSAGE == null)
                            {
                                errorMsg = "";
                            }
                            else
                            {
                                errorMsg = lp.MESSAGE;
                            }
                            string carry = "";
                            outputDelEquip += lp.USERID + "|" + lp.EQUIPMENTID + "|" + lp.EQUIPMENTCATEGORY + "|" + lp.SLOTNUMBER + "|" + lp.TIME_SENT + "|" + errorMsg;

                            outputDelEquip += "!";
                        }
                    }
                    System.Diagnostics.Debug.WriteLine("DEL EQUIP: " + outputDelEquip);

                    var queryDelPath = (from a in ctxData.BI_GRNOSPDELETEPATH_REQ
                                        join fx in ctxData.BI_GRNOSPDELETEPATH_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                        where a.PROC_ID == id
                                        select new
                                        {
                                            a.USERID,
                                            a.NAME,
                                            a.TYPE,
                                            a.TIME_SENT,
                                            fx.MESSAGE
                                        });
                    System.Diagnostics.Debug.WriteLine(queryDelPath);
                    if (queryDelPath.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryDelPath)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.MESSAGE == null)
                            {
                                errorMsg = "";
                            }
                            else
                            {
                                errorMsg = lp.MESSAGE;
                            }
                            string carry = "";
                            outputDelPath += lp.USERID + "|" + lp.NAME + "|" + lp.TYPE + "|" + lp.TIME_SENT + "|" + errorMsg;

                            outputDelPath += "!";
                        }
                    }
                    System.Diagnostics.Debug.WriteLine("DEL PATH: " + outputDelPath);

                    var queryDelSB = (from a in ctxData.BI_GRNOSPDELETESB_REQ
                                      join fx in ctxData.BI_GRNOSPDELETESB_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                      where a.PROC_ID == id
                                      select new
                                      {
                                          a.USERID,
                                          a.SERVICEBOUNDARYID,
                                          a.SITEDP,
                                          a.TIME_SENT,
                                          fx.MESSAGE
                                      });
                    System.Diagnostics.Debug.WriteLine(queryDelSB);
                    if (queryDelSB.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryDelSB)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.MESSAGE == null)
                            {
                                errorMsg = "";
                            }
                            else
                            {
                                errorMsg = lp.MESSAGE;
                            }
                            string carry = "";
                            outputDelSB += lp.USERID + "|" + lp.SERVICEBOUNDARYID + "|" + lp.SITEDP + "|" + lp.TIME_SENT + "|" + errorMsg;

                            outputDelSB += "!";
                        }
                    }
                    System.Diagnostics.Debug.WriteLine("DEL SB: " + outputDelSB);
                    // endRegion

                    var queryLPNC = (from a in ctxData.BI_GRNDLDPATHNONCONS_REQ
                                     join fx in ctxData.BI_GRNDLDPATHNONCONS_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                     where a.PROC_ID == id
                                     select new
                                     {
                                         a.ANAME,
                                         a.ASITE,
                                         a.ATYPE,
                                         a.ASLOT,
                                         a.ACARD,
                                         a.APORT,
                                         a.ZNAME,
                                         a.ZSITE,
                                         a.ZTYPE,
                                         a.ZSLOT,
                                         a.ZCARD,
                                         a.ZPORT,
                                         a.PRIMARYSECONDARY,
                                         a.PATHBW,
                                         a.TIME_SENT,
                                         fx.CALLSTATUS_MSG
                                     });
                    System.Diagnostics.Debug.WriteLine(queryLPNC);
                    if (queryLPNC.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryLPNC)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.CALLSTATUS_MSG == null)
                            {
                                errorMsg = "";
                            }
                            else
                            {
                                errorMsg = lp.CALLSTATUS_MSG;
                            }
                            string carry = "";
                            outputcrtLPNC += lp.ANAME + "|" + lp.ASITE + "|" + lp.ATYPE + "|" + lp.ASLOT + "|";
                            outputcrtLPNC += lp.ACARD + "|" + lp.APORT + "|" + lp.ZNAME + "|" + lp.ZSITE + "|";
                            outputcrtLPNC += lp.ZTYPE + "|" + lp.ZSLOT + "|" + lp.ZCARD + "|" + lp.ZPORT + "|";
                            outputcrtLPNC += lp.PRIMARYSECONDARY + "|" + lp.PATHBW + "|" + lp.TIME_SENT + "|" + errorMsg;

                            outputcrtLPNC += "!";
                        }
                    }

                    // region Mubin - CR43-20180330
                    var queryLPC = (from a in ctxData.BI_GRNOSPDLDPATHCONS_REQ
                                    join fx in ctxData.BI_GRNOSPDLDPATHCONS_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                    where a.PROC_ID == id
                                    select new
                                    {
                                        a.ANAME,
                                        a.ASITE,
                                        a.ATYPE,
                                        a.ACARD2,
                                        a.APORT2,
                                        a.ACARD3,
                                        a.APORT3,
                                        a.ZNAME,
                                        a.ZSITE,
                                        a.ZTYPE,
                                        a.ZCARD,
                                        a.ZPORT,
                                        a.DPNAME,
                                        a.DPSITE,
                                        a.APORT4,
                                        a.ACARD4,
                                        a.DP_TAMBAHAN,
                                        a.TIME_SENT,
                                        fx.CALLSTATUS_MSG
                                    });
                    System.Diagnostics.Debug.WriteLine(queryLPNC);
                    if (queryLPC.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryLPC)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.CALLSTATUS_MSG == null)
                            {
                                errorMsg = "";
                            }
                            else
                            {
                                errorMsg = lp.CALLSTATUS_MSG;
                            }
                            string add_dp = "";
                            if (lp.DP_TAMBAHAN == "UPDATE")
                            {
                                add_dp = "YES";
                            }
                            else if (lp.DP_TAMBAHAN == "NEW")
                            {
                                add_dp = "NO";
                            }
                            string carry = "";
                            outputcrtLPC += lp.ANAME + "|" + lp.ASITE + "|" + lp.ATYPE + "|" + lp.ACARD2 + "|";
                            outputcrtLPC += lp.APORT2 + "|" + lp.ACARD3 + "|" + lp.APORT3 + "|" + lp.ZNAME + "|";
                            outputcrtLPC += lp.ZSITE + "|" + lp.ZTYPE + "|" + lp.ZCARD + "|" + lp.ZPORT + "|";
                            outputcrtLPC += lp.ACARD4 + "|" + lp.APORT4 + "|";
                            outputcrtLPC += lp.DPNAME + "|" + lp.DPSITE + "|" + add_dp + "|" + lp.TIME_SENT + "|" + errorMsg;

                            outputcrtLPC += "!";
                        }
                    }
                    // endRegion

                    var queryLOSITE = (from a in ctxData.BI_GRNOSPLOADSITE_REQ
                                       join fx in ctxData.BI_GRNOSPLOADSITE_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                       where a.PROC_ID == id
                                       select new
                                       {
                                           a.SITENAME,
                                           a.SITEDESC,
                                           a.ADDRSTREETTYPE,
                                           a.ADDRSTREET,
                                           a.ADDRCOUNTY,
                                           a.ADDRCITY,
                                           a.ADDRSTATE,
                                           a.ADDRPOSTCODE,
                                           a.ADDRCOUNTRY,
                                           a.UDAEQUIPLOC,
                                           a.UDACABLINGTYPE,
                                           a.UDAFIBERTOPREMISEEXIST,
                                           a.UDAACCESSRESTRICT,
                                           a.UDACONTACT,
                                           a.TIME_SENT,
                                           fx.CSMSG
                                       });
                    System.Diagnostics.Debug.WriteLine(queryLOSITE);
                    if (queryLOSITE.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryLOSITE)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.CSMSG == null)
                            {
                                errorMsg = "";
                            }
                            else
                            {
                                errorMsg = lp.CSMSG;
                            }
                            string carry = "";
                            outputcrtLOSITE += lp.SITENAME + "|" + lp.SITEDESC + "|" + lp.ADDRSTREETTYPE + "," + lp.ADDRSTREET + ",";
                            outputcrtLOSITE += lp.ADDRCOUNTY + "," + lp.ADDRCITY + "," + lp.ADDRSTATE + "," + lp.ADDRPOSTCODE + ",";
                            outputcrtLOSITE += lp.ADDRCOUNTRY + "|" + lp.UDAEQUIPLOC + "|" + lp.UDACABLINGTYPE + "|" + lp.UDAFIBERTOPREMISEEXIST + "|";
                            outputcrtLOSITE += lp.UDAACCESSRESTRICT + "|" + lp.UDACONTACT + "|" + lp.TIME_SENT + "|" + errorMsg;

                            outputcrtLOSITE += "!";
                        }
                    }

                    var queryGSB = (from a in ctxData.BI_GRNOSPLOADSRVBND_REQ
                                    join fx in ctxData.BI_GRNOSPLOADSRVBND_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                    where a.PROC_ID == id
                                    //orderby a.PROJECT_NO, a.WBS_NUM
                                    select new
                                    {
                                        a.SRVBNDID,
                                        a.ADDRSITENO,
                                        a.ADDRSITEFLOOR,
                                        a.ADDRSITEBUILD,
                                        a.ADDRSTREETTYPE,
                                        a.ADDRSTREETNAME,
                                        a.ADDRSECTION,
                                        a.ADDRPOSTCODE,
                                        a.ADDRCITY,
                                        a.ADDRSTATE,
                                        a.ADDRCOUNTRY,
                                        a.ADDRPREMISETYPE,
                                        a.ADDRSITESDP,
                                        a.DROPCABLEDISTANCE,
                                        fx.CALLMSG
                                    });
                    System.Diagnostics.Debug.WriteLine(queryGSB);
                    if (queryGSB.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryGSB)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.CALLMSG == null)
                            {
                                errorMsg = "";
                            }
                            else
                            {
                                errorMsg = lp.CALLMSG;
                            }
                            string carry = "";
                            outputcrtGSB += lp.SRVBNDID + "|" + lp.ADDRSITENO + " " + lp.ADDRSITEFLOOR + " " + lp.ADDRSITEBUILD + "," + lp.ADDRSTREETTYPE + "," + lp.ADDRSTREETNAME + "," + lp.ADDRSECTION + ",";
                            outputcrtGSB += lp.ADDRPOSTCODE + "," + lp.ADDRCITY + "," + lp.ADDRSTATE + "," + lp.ADDRCOUNTRY + "|" + lp.ADDRPREMISETYPE + "|" + lp.ADDRSITESDP + "|" + lp.DROPCABLEDISTANCE + "|" + errorMsg;

                            outputcrtGSB += "!";
                        }
                    }

                    var queryDM = (from a in ctxData.BI_GRNOSPDATAMATCH_REQ
                                   join fx in ctxData.BI_GRNOSPDATAMATCH_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                   where a.PROC_ID == id
                                   select new
                                   {
                                       a.EVENTNAME,
                                       a.EQUIP_EQUIPID,
                                       a.EQUIP_EXCHDESC,
                                       a.EQUIP_SITE,
                                       a.TIME_SENT,
                                       fx.CALLSTATUS_MSG
                                   });
                    System.Diagnostics.Debug.WriteLine(queryDM);
                    if (queryDM.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryDM)
                        {
                            counter++;
                            string errorMsg = "";
                            if (lp.CALLSTATUS_MSG == null)
                            {
                                errorMsg = "";
                            }
                            else
                            {
                                errorMsg = lp.CALLSTATUS_MSG;
                            }
                            string carry = "";
                            outputcrtDM += lp.EVENTNAME + "|" + lp.EQUIP_EQUIPID + "|" + lp.EQUIP_EXCHDESC + "|" + lp.EQUIP_SITE + "|";
                            outputcrtDM += lp.TIME_SENT + "|" + errorMsg;

                            outputcrtDM += "!";
                        }
                    }
                }
            }

            using (Entities9 nepsbiDB = new Entities9())
            {
                var queryHandover = (from a in nepsbiDB.BI_PROC_SWIFT
                                     where a.NEPS_JOB_ID.Trim() == jobID.Trim() && a.PROC_ID == id
                                     select a).Count();

                if (queryHandover > 0)
                {

                    var queryNE = (from a in nepsbiDB.BI_SWIFTELEM_REQ
                                   join fx in nepsbiDB.BI_SWIFTELEM_RES on a.REQUEST_ID equals fx.REQUEST_ID
                                   where a.PROC_ID == id
                                   select new
                                   {
                                       a.EQUIP_ELEMENTCODE,
                                       a.EQUIP_ELEMENTDESC,
                                       a.EQUIP_INVENTORYSYSTEM,
                                       a.EQUIP_PROJECTID,
                                       a.EQUIP_NEPSREFNUMBER,
                                       a.EQUIP_PTT,
                                       a.EQUIP_STATUS,
                                       a.TIME_SENT,
                                       fx.ERROR_MESSAGE
                                   });
                    System.Diagnostics.Debug.WriteLine(queryNE);
                    if (queryNE.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryNE)
                        {
                            counter++;
                            string errorMsg = "";

                            string carry = "";
                            outputswift += lp.EQUIP_ELEMENTCODE + "|" + lp.EQUIP_ELEMENTDESC + "|" + lp.EQUIP_INVENTORYSYSTEM + "|" + lp.EQUIP_PROJECTID + "|" + lp.EQUIP_PTT + "|" + lp.EQUIP_STATUS + "|" + lp.TIME_SENT + "|" + lp.ERROR_MESSAGE;

                            outputswift += "!";
                        }
                    }
                }
                var querymhHandover = (from a in nepsbiDB.BI_PROCESS_MHDUCT
                                       where a.NEPS_JOB_ID.Trim() == jobID.Trim() && a.PROC_ID == id
                                       select a).Count();

                if (querymhHandover > 0)
                {

                    var queryNE = (from a in nepsbiDB.BI_MH_DPNEW_REQUEST
                                   join fx in nepsbiDB.BI_MH_DPNEW_REPLY on a.REQ_ID equals fx.REQ_ID
                                   where a.PROC_ID == id
                                   select new
                                   {
                                       a.FEATUREIDENTIFIER,
                                       a.ASSIGNMENTID,
                                       a.DUCTSOURCEID,
                                       a.DUCTTERMINATIONID,
                                       a.TOTALLENGTH,
                                       fx.CREATED_DATE,
                                       fx.ERR_MSG,
                                       fx.ERR_TYPE
                                   });
                    System.Diagnostics.Debug.WriteLine(queryNE);
                    if (queryNE.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryNE)
                        {
                            counter++;
                            string errorMsg = "";

                            string carry = "";
                            outputductpath += lp.FEATUREIDENTIFIER + "|" + lp.ASSIGNMENTID + "|" + lp.DUCTSOURCEID + "|" + lp.DUCTTERMINATIONID + "|" + lp.TOTALLENGTH + "|" + lp.CREATED_DATE + "|" + lp.ERR_MSG;

                            outputductpath += "!";
                        }
                    }
                }

                if (querymhHandover > 0)
                {

                    var queryNE = (from a in nepsbiDB.BI_MH_MNEW_REQUEST
                                   join fx in nepsbiDB.BI_MH_MNEW_REPLY on a.REQ_ID equals fx.REQ_ID
                                   where a.PROC_ID == id
                                   select new
                                   {
                                       a.FEATUREIDENTIFIER,
                                       a.EXCHANGEABBREVIATION,
                                       a.ASSIGNMENTID,
                                       a.LATITUDE,
                                       a.LONGTITUDE,
                                       a.OWNER,
                                       fx.CREATED_DATE,
                                       fx.ERR_MSG,
                                       fx.ERR_TYPE
                                   });
                    System.Diagnostics.Debug.WriteLine(queryNE);
                    if (queryNE.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryNE)
                        {
                            counter++;
                            string errorMsg = "";

                            string carry = "";
                            outputmh += lp.FEATUREIDENTIFIER + "|" + lp.EXCHANGEABBREVIATION + "|" + lp.ASSIGNMENTID + "|" + lp.LATITUDE + "|" + lp.LONGTITUDE + "|" + lp.OWNER + "|" + lp.CREATED_DATE + "|" + lp.ERR_MSG;

                            outputmh += "!";
                        }
                    }
                }
                if (querymhHandover > 0)
                {

                    var queryNE = (from a in nepsbiDB.BI_MH_DNEW_REQUEST
                                   join fx in nepsbiDB.BI_MH_DNEW_REPLY on a.REQ_ID equals fx.REQ_ID
                                   where a.PROC_ID == id
                                   select new
                                   {
                                       a.FEATUREIDENTIFIER,
                                       a.ASSIGNMENTID,
                                       a.DUCTOWNERSHIP,
                                       fx.CREATED_DATE,
                                       fx.ERR_MSG,
                                       fx.ERR_TYPE
                                   });
                    System.Diagnostics.Debug.WriteLine(queryNE);
                    if (queryNE.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryNE)
                        {
                            counter++;
                            string errorMsg = "";

                            string carry = "";
                            outputduct += lp.FEATUREIDENTIFIER + "|" + lp.ASSIGNMENTID + "|" + lp.DUCTOWNERSHIP + "|" + lp.CREATED_DATE + "|" + lp.ERR_MSG;

                            outputduct += "!";
                        }
                    }
                }
                if (querymhHandover > 0)
                {

                    var queryNE = (from a in nepsbiDB.BI_MH_SDNEW_REQUEST
                                   join fx in nepsbiDB.BI_MH_SDNEW_REPLY on a.REQ_ID equals fx.REQ_ID
                                   where a.PROC_ID == id
                                   select new
                                   {
                                       a.SUBDUCTID,
                                       a.ASSIGNMENTID,
                                       a.SUBDUCTOWNERSHIP,
                                       a.NOOFWAYS,
                                       a.SUBDUCTOCCUPIED,
                                       fx.CREATED_DATE,
                                       fx.ERR_MSG,
                                       fx.ERR_TYPE
                                   });
                    System.Diagnostics.Debug.WriteLine(queryNE);
                    if (queryNE.Count() > 0)
                    {
                        int counter = 0;
                        foreach (var lp in queryNE)
                        {
                            counter++;
                            string errorMsg = "";

                            string carry = "";
                            outputsubduct += lp.SUBDUCTID + "|" + lp.ASSIGNMENTID + "|" + lp.NOOFWAYS + "|" + lp.SUBDUCTOCCUPIED + "|" + lp.SUBDUCTOWNERSHIP + "|" + lp.CREATED_DATE + "|" + lp.ERR_MSG;

                            outputsubduct += "!";
                        }
                    }
                }

            }

            return Json(new
            {
                crtCardR = outputcrtCardR,
                crtFCR = outputcrtFC,
                crtFUR = outputcrtFU,
                crtNE = outputcrtNE,
                crtCS = outputcrtCS,
                crtSB = outputcrtSB,
                crtLE = outputcrtLE,
                crtLPNC = outputcrtLPNC,
                crtLPC = outputcrtLPC,
                crtLOSITE = outputcrtLOSITE,
                crtGSB = outputcrtGSB,
                crtDM = outputcrtDM,
                #region Mubin - CR58-20180330
                addCard = outputAddCard,
                delEquip = outputDelEquip,
                delPath = outputDelPath,
                delSB = outputDelSB,
                #endregion
                swift = outputswift,
                manhole = outputmh,
                ductpath = outputductpath,
                duct = outputduct,
                subduct = outputsubduct
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult ChooseProject(string targetJob, string targetProject, string targetWBS)
        {
            System.Diagnostics.Debug.WriteLine(targetJob + "-!!!!!-" + targetProject);
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            string res = "";
            string res1 = "";
            string res2 = "";
            using (Entities ctxData = new Entities())
            {
                var CariOSP = (from p in ctxData.G3E_JOB
                               where p.SCHEME_NAME.ToUpper() == targetJob.ToUpper()
                               select p).Count();

                var CariISP = (from p in ctxData.WV_ISP_JOB
                               where p.SCHEME_NAME.ToUpper() == targetJob.ToUpper()
                               select p).Count();
                if (CariOSP > 0)
                {
                    res1 = myWebService.ChooseProject(targetJob.ToUpper(), targetProject, targetWBS);
                }
                if (CariISP > 0)
                {
                    var CariISP2 = (from p in ctxData.WV_ISP_JOB
                                    where p.SCHEME_NAME.ToUpper() == targetJob.ToUpper()
                                    select p).Single();
                    res2 = myWebService.ISPChooseProject(CariISP2.G3E_IDENTIFIER, targetProject, targetWBS);
                }
            }
            //string res = myWebService.AutoApproval(targetJob);
            if (res1 == "fail")
            {
                res = res2;
            }
            else
            {
                res = res1;
            }

            return Json(new
            {
                result = res
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult ISPChooseProject(string targetJob, string targetProject, string targetWBS)
        {
            System.Diagnostics.Debug.WriteLine(targetJob + "-!!!!!-" + targetProject);
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            string res = "";
            string res1 = "";
            string res2 = "";
            using (Entities ctxData = new Entities())
            {
                var CariISP = (from p in ctxData.WV_ISP_JOB
                               where p.SCHEME_NAME.ToUpper() == targetJob.ToUpper()
                               select p).Count();

                var CariDataISP = (from p in ctxData.WV_ISP_JOB
                                   where p.SCHEME_NAME.ToUpper() == targetJob.ToUpper()
                                   select p).Single();

                var CariOSP = (from p in ctxData.G3E_JOB
                               where p.SCHEME_NAME == CariDataISP.SCHEME_NAME
                               select p).Count();

                if (CariOSP > 0)
                {
                    var CariOSP2 = (from p in ctxData.G3E_JOB
                                    where p.G3E_IDENTIFIER.ToUpper() == CariDataISP.SCHEME_NAME.ToUpper()
                                    select p).Single();
                    res1 = myWebService.ChooseProject(CariOSP2.G3E_IDENTIFIER.ToUpper(), targetProject, targetWBS);
                }
                if (CariISP > 0)
                {
                    res2 = myWebService.ISPChooseProject(targetJob, targetProject, targetWBS);
                }
            }
            //string res = myWebService.ISPChooseProject(targetJob, targetProject);
            System.Diagnostics.Debug.WriteLine(res2);
            if (res2 == "fail")
            {
                res = res1;
            }
            else
            {
                res = res2;
            }
            return Json(new
            {
                result = res
            }, JsonRequestBehavior.AllowGet);
        }

        // region Mubin - CR72-2018 - 2 Jan 2018
        [HttpPost]
        public ActionResult UnmatchProject(string targetJob)
        {
            System.Diagnostics.Debug.WriteLine(targetJob + "-!!!!!");
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            string res = "";
            string res1 = "";
            string res2 = "";
            using (Entities ctxData = new Entities())
            {
                var CariOSP = (from p in ctxData.G3E_JOB
                               where p.SCHEME_NAME.ToUpper() == targetJob.ToUpper()
                               select p).Count();

                var CariISP = (from p in ctxData.WV_ISP_JOB
                               where p.SCHEME_NAME.ToUpper() == targetJob.ToUpper()
                               select p).Count();
                if (CariOSP > 0)
                {
                    res1 = myWebService.UnmatchProject(targetJob.ToUpper());
                }
                if (CariISP > 0)
                {
                    res2 = myWebService.ISPUnmatchProject(targetJob.ToUpper());
                }
            }
            if (res1 == "fail")
            {
                res = res2;
            }
            else
            {
                res = res1;
            }
            return Json(new
            {
                result = res
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult ISPUnmatchProject(string targetJob)
        {
            System.Diagnostics.Debug.WriteLine(targetJob + "-!!!!!");
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            string res = "";
            string res1 = "";
            string res2 = "";
            using (Entities ctxData = new Entities())
            {
                var CariISP = (from p in ctxData.WV_ISP_JOB
                               where p.SCHEME_NAME.ToUpper() == targetJob.ToUpper()
                               select p).Count();

                var CariOSP = (from p in ctxData.G3E_JOB
                               where p.SCHEME_NAME.ToUpper() == targetJob.ToUpper()
                               select p).Count();

                if (CariOSP > 0)
                {
                    res1 = myWebService.UnmatchProject(targetJob.ToUpper());
                }
                if (CariISP > 0)
                {
                    res2 = myWebService.ISPUnmatchProject(targetJob.ToUpper());
                }
            }
            if (res2 == "fail")
            {
                res = res1;
            }
            else
            {
                res = res2;
            }
            return Json(new
            {
                result = res
            }, JsonRequestBehavior.AllowGet);
        }
        // endRegion

        // region Mubin - CR74-20180330
        //region Fatihin - 19042018
        [HttpPost]
        public ActionResult AutoApproval(string targetJob, string type)
        {
            System.Diagnostics.Debug.WriteLine(targetJob + "-!!!!!-");
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            string res = "";
            string res1 = "";
            string res2 = "";
            int webser = 0;

            OSPHandoverServices.Error errorObj = null;
            Error_ISP errorISP = null;
            using (Entities ctxData = new Entities())
            {
                var CariOSP = (from p in ctxData.G3E_JOB
                               where p.G3E_IDENTIFIER == targetJob
                               select p).Count();

                var CariISP = (from p in ctxData.WV_ISP_JOB
                               where p.SCHEME_NAME == targetJob
                               select p).Count();

                //bool checkFeat;
                //checkFeat = myWebService.CheckFeaturejob(targetJob)
                // fatihin 19042018
                if (CariOSP > 0 && type.Contains("OSP"))
                {
                    var dataOSP = (from p in ctxData.G3E_JOB
                                   where p.G3E_IDENTIFIER == targetJob
                                   select p).Single();
                    if (dataOSP.JOB_TYPE == "Civil")
                    {
                        var queryNET = (from d in ctxData.GC_NETELEM
                                        where d.JOB_ID == targetJob && (d.FEATURE_STATE != "ASB" || d.FEATURE_STATE != "UC")
                                        select d).Count();
                        int checkFeat = queryNET;
                        if (checkFeat == 0)
                        {
                            res1 = myWebService.AutoApproval(targetJob);
                        }
                        else
                        { res1 = "fail"; }

                    }
                    else
                    {
                        neps_procmanClient testWSOSP = new neps_procmanClient();
                        testWSOSP.CreateProcess(out errorObj, targetJob);
                        //var queryNET = (from d in ctxData.GC_NETELEM
                        //                where d.JOB_ID == targetJob && (d.FEATURE_STATE != "ASB" || d.FEATURE_STATE != "UC")
                        //                select d).Count();
                        //int checkFeat = queryNET;
                        //if (checkFeat == 0)
                        //{
                        using (EntitiesNetworkElement chckData = new EntitiesNetworkElement())
                        {
                            var checkDataOSP = (from p in chckData.BI_PROCESS
                                                where p.NEPS_JOB_ID == targetJob && p.STATUS == "SUCCESS"
                                                select p).Count();

                            if (checkDataOSP > 0)
                            {
                                var queryASB = (from d in ctxData.GC_NETELEM
                                                where d.JOB_ID == targetJob && d.FEATURE_STATE != "ASB"
                                                select d).Count();


                                int checkASB = queryASB;
                                if (checkASB == 0)
                                {
                                    res1 = myWebService.AutoApproval(targetJob);
                                }
                                else
                                {
                                    res1 = "fail";
                                }
                            }
                            else
                            {
                                res1 = "fail";
                            }

                        }
                        int idNEPSBI_PROC;
                        using (EntitiesNetworkElement data_nepsbi = new EntitiesNetworkElement())
                        {
                            var id_process = (from p in data_nepsbi.BI_PROCESS
                                              select p).OrderByDescending(p => p.PROC_ID).First();
                            idNEPSBI_PROC = Convert.ToInt32(id_process.PROC_ID);
                        }

                        using (Entities6 chckData = new Entities6())
                        {
                            var dataGE = (from p in chckData.VW_CREATENE_GE_OSP
                                          where p.SCHEME_NAME == targetJob
                                          select p);

                            foreach (var data in dataGE)
                            {
                                int idNEPSBI_REQ;
                                using (EntitiesNetworkElement data_nepsbi = new EntitiesNetworkElement())
                                {
                                    var id_process = (from p in data_nepsbi.BI_CREATENE_REQUEST_OSP
                                                      select p).OrderByDescending(p => p.REQUEST_ID).First();
                                    idNEPSBI_REQ = Convert.ToInt32(id_process.REQUEST_ID) + 1;
                                }
                                Tools tool = new Tools();
                                string sqlStr = "insert into NEPSBI.BI_CREATENE_REQUEST_OSP(REQUEST_ID ,PROC_ID ,EQUPLOCNTTNAME ,EQUPEQUTABBREAVIATION ,EQUPINDEX ,LOCNDETAIL ,AREACODE ,AREADESCRIPTION,AREAARETCODE ,AREAAREACODE ,EQUPSTATUS ,EQUPMANRABBREAVIATION ,EQUPEQUMMODEL ,G3E_FID ,TIME_SENT,TIME_RETURNED) values";
                                OracleParameter[] oraPrm = new OracleParameter[16];
                                oraPrm[0] = new OracleParameter("v_REQUEST_ID", OracleDbType.Int32);
                                oraPrm[0].Value = idNEPSBI_REQ;
                                oraPrm[1] = new OracleParameter("v_PROC_ID", OracleDbType.Int32);
                                oraPrm[1].Value = idNEPSBI_PROC;
                                oraPrm[2] = new OracleParameter("v_EQUPLOCNTTNAME", OracleDbType.Varchar2);
                                oraPrm[2].Value = data.EQUIPLOCNAME;
                                oraPrm[3] = new OracleParameter("v_EQUPEQUTABBREAVIATION", OracleDbType.Varchar2);
                                oraPrm[3].Value = data.EQUIPMENTABB;
                                oraPrm[4] = new OracleParameter("v_EQUPINDEX", OracleDbType.Varchar2);
                                oraPrm[4].Value = data.GE_INDEX;
                                oraPrm[5] = new OracleParameter("v_LOCNDETAIL", OracleDbType.Varchar2);
                                oraPrm[5].Value = data.LOCDETAIL;
                                oraPrm[6] = new OracleParameter("v_AREACODE", OracleDbType.Varchar2);
                                oraPrm[6].Value = data.AREACODE;
                                oraPrm[7] = new OracleParameter("v_AREADESCRIPTION", OracleDbType.Varchar2);
                                oraPrm[7].Value = data.AREA_DESC;
                                oraPrm[8] = new OracleParameter("v_AREAARETCODE", OracleDbType.Varchar2);
                                oraPrm[8].Value = data.AREA;
                                oraPrm[9] = new OracleParameter("v_AREAAREACODE", OracleDbType.Varchar2);
                                oraPrm[9].Value = data.AREAREACODE;
                                oraPrm[10] = new OracleParameter("v_EQUPSTATUS", OracleDbType.Varchar2);
                                oraPrm[10].Value = data.STATUS;
                                oraPrm[11] = new OracleParameter("v_EQUPMANRABBREAVIATION", OracleDbType.Varchar2);
                                oraPrm[11].Value = data.MANUFACTURER;
                                oraPrm[12] = new OracleParameter("v_EQUPEQUMMODEL", OracleDbType.Varchar2);
                                oraPrm[12].Value = data.MODEL;
                                oraPrm[13] = new OracleParameter("v_G3E_FID", OracleDbType.Varchar2);
                                oraPrm[13].Value = data.G3E_FID;
                                oraPrm[14] = new OracleParameter("v_TIME_SENT", OracleDbType.TimeStamp);
                                oraPrm[14].Value = System.DateTime.Now;
                                oraPrm[15] = new OracleParameter("v_TIME_RETURNED", OracleDbType.TimeStamp);
                                oraPrm[15].Value = System.DateTime.Now;
                                string result2 = tool.ExecuteStored(connString, sqlStr, CommandType.StoredProcedure, oraPrm, false);
                                System.Diagnostics.Debug.WriteLine("result :" + result2);


                                createNEReply DataResNE = new createNEReply();
                                TM_NEPSManagementAPIClient WS = new TM_NEPSManagementAPIClient();
                                DataResNE = WS.createNetworkElement(data.EQUIPLOCNAME, data.EQUIPMENTABB, data.GE_INDEX, data.LOCDETAIL, data.AREACODE, data.AREA_DESC, data.AREA, data.AREAREACODE, data.STATUS, data.MANUFACTURER, data.MODEL, data.G3E_FID.ToString());

                                int idNEPSBI_RES;
                                using (EntitiesNetworkElement data_nepsbi = new EntitiesNetworkElement())
                                {
                                    var id_process = (from p in data_nepsbi.BI_CREATENE_REPLY_OSP
                                                      select p).OrderByDescending(p => p.REPLY_ID).First();
                                    idNEPSBI_RES = Convert.ToInt32(id_process.REPLY_ID) + 1;
                                }
                                Tools tool2 = new Tools();
                                string sqlStr2 = "insert into NEPSBI.BI_CREATENE_REPLY_OSP(REPLY_ID,REQUEST_ID ,PROC_ID ,ERRORCODE ,ERRORMSG ,EQUP_ID ,G3E_FID,TIME_RETURNED) values";
                                OracleParameter[] oraPrm2 = new OracleParameter[8];
                                oraPrm2[0] = new OracleParameter("v_REPLY_ID", OracleDbType.Int32);
                                oraPrm2[0].Value = idNEPSBI_RES;
                                oraPrm2[1] = new OracleParameter("v_REQUEST_ID", OracleDbType.Int32);
                                oraPrm2[1].Value = idNEPSBI_REQ;
                                oraPrm2[2] = new OracleParameter("v_PROC_ID", OracleDbType.Int32);
                                oraPrm2[2].Value = idNEPSBI_PROC;
                                oraPrm2[3] = new OracleParameter("v_ERRORCODE", OracleDbType.Varchar2);
                                oraPrm2[3].Value = DataResNE.errorCode;
                                oraPrm2[4] = new OracleParameter("v_ERRORMSG", OracleDbType.Varchar2);
                                oraPrm2[4].Value = DataResNE.errorMsg;
                                oraPrm2[5] = new OracleParameter("v_EQUP_ID", OracleDbType.Varchar2);
                                oraPrm2[5].Value = DataResNE.equp_Id;
                                oraPrm2[6] = new OracleParameter("v_G3E_FID", OracleDbType.Varchar2);
                                oraPrm2[6].Value = data.G3E_FID;
                                oraPrm2[7] = new OracleParameter("v_TIME_RETURNED", OracleDbType.TimeStamp);
                                oraPrm2[7].Value = System.DateTime.Now;
                                string result3 = tool2.ExecuteStored(connString, sqlStr2, CommandType.StoredProcedure, oraPrm2, false);
                                System.Diagnostics.Debug.WriteLine("result :" + result3);
                            }

                            var dataGECARD = (from p in chckData.VW_CREATECARD_GE_OSP
                                              where p.JOBID == targetJob
                                              select p);

                            foreach (var dataCard in dataGECARD)
                            {
                                int idNEPSBI_REQ;
                                string val_equpid = "";
                                int count_equip = 0;
                                using (EntitiesNetworkElement data_nepsbi = new EntitiesNetworkElement())
                                {
                                    var id_process = (from p in data_nepsbi.BI_CREATECARD_REQUEST_OSP
                                                      select p).OrderByDescending(p => p.REQUEST_ID).First();
                                    idNEPSBI_REQ = Convert.ToInt32(id_process.REQUEST_ID) + 1;

                                    string msanFid = dataCard.MSAN_FID.ToString();
                                    count_equip = (from p in data_nepsbi.BI_CREATENE_REPLY_OSP
                                                   where p.G3E_FID == msanFid && (p.EQUP_ID != null && p.EQUP_ID != 0)
                                                   select p.EQUP_ID).Count();
                                    if (count_equip > 0)
                                    {
                                        var equpid = (from p in data_nepsbi.BI_CREATENE_REPLY_OSP
                                                      where p.G3E_FID == msanFid && (p.EQUP_ID != null && p.EQUP_ID != 0)
                                                      select p.EQUP_ID).First();
                                        val_equpid = equpid.ToString();
                                    }
                                }
                                string lowPort = dataCard.PORTSTARTNUM;
                                if (dataCard.PORTSTARTNUM == "" || dataCard.PORTSTARTNUM == null)
                                {
                                    lowPort = "1";
                                }
                                string hiport = dataCard.CARDCOUNTPORT;
                                if (dataCard.CARDCOUNTPORT == "" || dataCard.CARDCOUNTPORT == null)
                                {
                                    hiport = "1";
                                }
                                if (count_equip > 0)
                                {
                                    Tools tool = new Tools();
                                    string sqlStr = "insert into NEPSBI.BI_CREATECARD_REQUEST_OSP(REQUEST_ID ,PROC_ID ,EQUPID ,CARDSLOT ,CARDNAME ,CARDMODEL ,CARDSTATUS ,CARDCOUNTPORT ,PORTSTARTNUM ,G3E_FID ,TIME_SENT,TIME_RETURNED) values";
                                    OracleParameter[] oraPrm = new OracleParameter[12];
                                    oraPrm[0] = new OracleParameter("v_REQUEST_ID", OracleDbType.Int32);
                                    oraPrm[0].Value = idNEPSBI_REQ;
                                    oraPrm[1] = new OracleParameter("v_PROC_ID", OracleDbType.Int32);
                                    oraPrm[1].Value = idNEPSBI_PROC;
                                    oraPrm[2] = new OracleParameter("v_EQUPID", OracleDbType.Varchar2);
                                    oraPrm[2].Value = val_equpid;
                                    oraPrm[3] = new OracleParameter("v_CARDSLOT", OracleDbType.Varchar2);
                                    oraPrm[3].Value = dataCard.CARDSLOT;
                                    oraPrm[4] = new OracleParameter("v_CARDNAME", OracleDbType.Varchar2);
                                    oraPrm[4].Value = dataCard.CARDNAME;
                                    oraPrm[5] = new OracleParameter("v_CARDMODEL", OracleDbType.Varchar2);
                                    oraPrm[5].Value = dataCard.CARDMODEL;
                                    oraPrm[6] = new OracleParameter("v_CARDSTATUS", OracleDbType.Varchar2);
                                    oraPrm[6].Value = "PLANNED";
                                    oraPrm[7] = new OracleParameter("v_CARDCOUNTPORT", OracleDbType.Varchar2);
                                    oraPrm[7].Value = hiport;
                                    oraPrm[8] = new OracleParameter("v_PORTSTARTNUM", OracleDbType.Varchar2);
                                    oraPrm[8].Value = lowPort;
                                    oraPrm[9] = new OracleParameter("v_G3E_FID", OracleDbType.Varchar2);
                                    oraPrm[9].Value = dataCard.G3E_FID.ToString();
                                    oraPrm[10] = new OracleParameter("v_TIME_SENT", OracleDbType.TimeStamp);
                                    oraPrm[10].Value = System.DateTime.Now;
                                    oraPrm[11] = new OracleParameter("v_TIME_RETURNED", OracleDbType.TimeStamp);
                                    oraPrm[11].Value = System.DateTime.Now;
                                    string result2 = tool.ExecuteStored(connString, sqlStr, CommandType.StoredProcedure, oraPrm, false);
                                    System.Diagnostics.Debug.WriteLine("result :" + result2);

                                    createCardReply DataResCard = new createCardReply();
                                    TM_NEPSManagementAPIClient WS = new TM_NEPSManagementAPIClient();
                                    DataResCard = WS.createCard(val_equpid, dataCard.CARDSLOT, dataCard.CARDNAME, dataCard.CARDMODEL, "PLANNED", "", Convert.ToInt32(dataCard.CARDCOUNTPORT), dataCard.PORTSTARTNUM, dataCard.G3E_FID.ToString());

                                    int idNEPSBI_RES;
                                    using (EntitiesNetworkElement data_nepsbi = new EntitiesNetworkElement())
                                    {
                                        var id_process = (from p in data_nepsbi.BI_CREATECARD_REPLY_OSP
                                                          select p).OrderByDescending(p => p.REPLY_ID).First();
                                        idNEPSBI_RES = Convert.ToInt32(id_process.REPLY_ID) + 1;
                                    }

                                    Tools tool2 = new Tools();
                                    string sqlStr2 = "insert into NEPSBI.BI_CREATECARD_REPLY_OSP(REPLY_ID,REQUEST_ID ,PROC_ID ,ERRORCODE ,ERRORMSG ,GE3FID,TIME_RETURNED) values";
                                    OracleParameter[] oraPrm2 = new OracleParameter[7];
                                    oraPrm2[0] = new OracleParameter("v_REPLY_ID", OracleDbType.Int32);
                                    oraPrm2[0].Value = idNEPSBI_RES;
                                    oraPrm2[1] = new OracleParameter("v_REQUEST_ID", OracleDbType.Int32);
                                    oraPrm2[1].Value = idNEPSBI_REQ;
                                    oraPrm2[2] = new OracleParameter("v_PROC_ID", OracleDbType.Int32);
                                    oraPrm2[2].Value = idNEPSBI_PROC;
                                    oraPrm2[3] = new OracleParameter("v_ERRORCODE", OracleDbType.Varchar2);
                                    oraPrm2[3].Value = DataResCard.errorCode;
                                    oraPrm2[4] = new OracleParameter("v_ERRORMSG", OracleDbType.Varchar2);
                                    oraPrm2[4].Value = DataResCard.errorMsg;
                                    oraPrm2[5] = new OracleParameter("v_G3E_FID", OracleDbType.Varchar2);
                                    oraPrm2[5].Value = dataCard.G3E_FID;
                                    oraPrm2[6] = new OracleParameter("v_TIME_RETURNED", OracleDbType.TimeStamp);
                                    oraPrm2[6].Value = System.DateTime.Now;
                                    string result3 = tool2.ExecuteStored(connString, sqlStr2, CommandType.StoredProcedure, oraPrm2, false);
                                    System.Diagnostics.Debug.WriteLine("result :" + result3);
                                }
                            }
                        }


                        //}
                        //else
                        //{
                        //    res1 = "fail";
                        //}
                    }
                }
                // region Mubin - CR74-20180330
                // region fatihin - 19042018
                if (CariISP > 0 && type.Contains("ISP"))
                {
                    // endRegion
                    neps_procman_ispClient testWSISP = new neps_procman_ispClient();
                    testWSISP.CreateProcess(out errorISP, targetJob);

                    var CariISP2 = (from p in ctxData.WV_ISP_JOB
                                    where p.SCHEME_NAME == targetJob
                                    select p).Single();

                    using (EntitiesNetworkElement chckData = new EntitiesNetworkElement())
                    {
                        var checkDataISP = (from p in chckData.BI_PROCESS_ISP
                                            where p.NEPS_JOB_ID == targetJob && p.STATUS == "SUCCESS"
                                            select p).Count();
                        if (checkDataISP > 0)
                        {
                            if (CariISP2.JOB_TYPE == "ISP" || CariISP2.JOB_TYPE == "Others")
                            {
                                res2 = myWebService.ISPAutoApproval(CariISP2.G3E_IDENTIFIER);
                            }
                            else
                            {
                                //if (checkFeat == 0)
                                //{
                                res2 = myWebService.ISPAutoApproval(CariISP2.G3E_IDENTIFIER);
                                //}
                                //else
                                //{ res2 = "fail"; }
                            }
                        }
                        else
                            res2 = "fail";
                    }
                }
            }
            //string res = myWebService.AutoApproval(targetJob);
            if (res1 == "fail")
            {
                res = res2;
            }
            else
            {
                res = res1;
            }

            //MailMessage msg = new MailMessage();
            //msg.IsBodyHtml = true;
            //msg.From = new MailAddress("dev-noreply@tm.com", "NEPS Webview.");
            //msg.To.Add("haidarsukur@gmail.com");
            //msg.Subject = "TESTING";
            //msg.Body = "TEST";
            //SmtpClient emailClient = new SmtpClient("smtp.tm.com.my");
            //System.Net.NetworkCredential SMTPUserInfo = new System.Net.NetworkCredential("neps@tm.com.my", "nepsadmin");
            //emailClient.UseDefaultCredentials = false;
            //emailClient.Port = 25;
            //emailClient.Credentials = SMTPUserInfo;
            //emailClient.Send(msg);

            return Json(new
            {
                result = res
            }, JsonRequestBehavior.AllowGet);
        }

        internal class ExampleAPIProxy
        {

            private static WebService2 ExampleAPI = new WebService2("http://10.14.29.4:7280/prj_HsbbEai_Sync_War/ws/NEPSLoadPathConsumerRequest.wsdl");    // DEFAULT location of the WebService, containing the WebMethods

            public static void ChangeUrl(string webserviceEndpoint)
            {
                ExampleAPI = new WebService2(webserviceEndpoint);
            }

            public static string ExampleWebMethod(string name, string userID, string jobid)
            {
                using (Entities6 E6_data = new Entities6())
                {
                    var path_WS_URL = (from dataWS in E6_data.REF_WEBSERVICES
                                       where dataWS.SERVICE == "GRANITE" && dataWS.SERVICE_TYPE == "PATHCONSUMER"
                                       select new { dataWS.SERVICE_URL }).Single();

                    ChangeUrl(path_WS_URL.SERVICE_URL.ToString());
                }

                ExampleAPI.AddParameter("eventName", name);                    // Case Sensitive! To avoid typos, just copy the WebMethod's signature and paste it
                ExampleAPI.AddParameter("UserId", userID);
                ExampleAPI.AddParameter("JobId", jobid);  // all parameters are passed as strings
                try
                {
                    ExampleAPI.Invoke("NEPSLoadPathConsumer");                // name of the WebMethod to call (Case Sentitive again!)
                }
                finally { }

                return ExampleAPI.ResultString;                           // you can either return a string or an XML, your choice
            }
        }

        // region Mubin - CR74-20180330
        [HttpPost]
        //Fatihin 16052018
        public ActionResult GraniteApproval2(string targetJob, string JOBTYPE)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            GraniteISP.Error errorObj = null;
            string res = "";
            int result = 0;

            if (!targetJob.Contains("CIVIL"))
            {
                using (Entities ctxData = new Entities())
                {
                    var CariOSP = (from p in ctxData.G3E_JOB
                                   where p.G3E_IDENTIFIER == targetJob
                                   select p).Count();

                    var CariISP = (from p in ctxData.WV_ISP_JOB
                                   where p.G3E_IDENTIFIER == targetJob
                                   select p).Count();

                    //bool callProc = myWebService.CallProcedureGRN(targetJob);

                    //Fatihin 16052018
                    if (CariOSP > 0 && JOBTYPE.Contains("OSP"))
                    {
                        var dataOSP = (from p in ctxData.G3E_JOB
                                       where p.G3E_IDENTIFIER == targetJob
                                       select p).Single();

                        EntitiesNetworkElement nepsbiDB = new EntitiesNetworkElement();
                        var maxProcValue = nepsbiDB.BI_PROC_GRN_OSP.Max(x => x.PROC_ID);
                        int v_proc_id = Convert.ToInt32(maxProcValue) + 1;

                        BI_PROC_GRN_OSP v_proc = new BI_PROC_GRN_OSP();
                        v_proc.PROC_ID = v_proc_id;
                        v_proc.NEPS_JOB_ID = targetJob;
                        v_proc.STATUS = "NEW";
                        v_proc.TIME_CREATED = System.DateTime.Now;
                        v_proc.USER_ID = dataOSP.G3E_OWNER;
                        v_proc.CLASS_NAME = "EdgeFrontier.OSP.Granite";
                        nepsbiDB.BI_PROC_GRN_OSP.AddObject(v_proc);
                        nepsbiDB.SaveChanges();
                    }
                    //Fatihin 16052018
                    if (CariISP > 0 && JOBTYPE.Contains("ISP"))
                    {
                        Entities9 nepsbiDB = new Entities9();
                        var dataISP = (from p in ctxData.WV_ISP_JOB
                                       where p.G3E_IDENTIFIER == targetJob
                                       select p).Single();

                        var maxProcValue = nepsbiDB.BI_PROC_GRN_ISP21.Max(x => x.PROC_ID);
                        int v_proc_id = Convert.ToInt32(maxProcValue) + 1;

                        BI_PROC_GRN_ISP2 v_proc = new BI_PROC_GRN_ISP2();
                        v_proc.PROC_ID = v_proc_id;
                        v_proc.NEPS_JOB_ID = targetJob;
                        v_proc.STATUS = "NEW";
                        v_proc.TIME_CREATED = System.DateTime.Now;
                        v_proc.USER_ID = dataISP.G3E_OWNER;
                        v_proc.CLASS_NAME = "EdgeFrontier.ISP.Granite";
                        nepsbiDB.BI_PROC_GRN_ISP21.AddObject(v_proc);
                        nepsbiDB.SaveChanges();
                        //var dataISP = (from p in ctxData.WV_ISP_JOB
                        //               where p.G3E_IDENTIFIER == targetJob
                        //               select p).Single();

                        //string username = dataISP.G3E_OWNER;

                        //neps_procman_grn_ispClient Grn = new neps_procman_grn_ispClient();
                        //result = Grn.CreateProcess(out errorObj, targetJob, username);
                    }

                    //myWebService.GRNAutoApproval(targetJob);

                    if (result == 0)
                    {

                        if (CariOSP > 0)
                        {
                            myWebService.AutoApproval(targetJob);
                        }

                        if (CariISP > 0)
                        {
                            myWebService.ISPAutoApproval(targetJob);
                        }
                    }
                }
            }

            else
            {
                Entities7 neps = new Entities7();
                int count = (from p in neps.VW_MANHOLEDUCT_JOB
                             where (p.JOB_ID.Contains(targetJob))
                             select p).Count();

                if (count > 0)
                {
                    Entities9 nepsbiDB = new Entities9();
                    var maxProcValue = nepsbiDB.BI_PROCESS_MHDUCT.Max(x => x.PROC_ID);
                    int v_proc_id = Convert.ToInt32(maxProcValue) + 1;

                    BI_PROCESS_MHDUCT v_proc = new BI_PROCESS_MHDUCT();
                    v_proc.PROC_ID = v_proc_id;
                    v_proc.NEPS_JOB_ID = targetJob;
                    v_proc.STATUS = "NEW";
                    v_proc.TIME_CREATED = System.DateTime.Now;
                    nepsbiDB.BI_PROCESS_MHDUCT.AddObject(v_proc);
                    nepsbiDB.SaveChanges();
                }
            }
            /*string result2 = "";
            Tools tool = new Tools();
            OracleParameter[] oraPrm2;
            oraPrm2 = new OracleParameter[2];
            oraPrm2[0] = new OracleParameter("schemeName", OracleDbType.Varchar2);
            oraPrm2[0].Value = targetJob;
            oraPrm2[1] = new OracleParameter("SEND_TYPE", OracleDbType.Varchar2);
            oraPrm2[1].Value = "ASB";
            result2 = tool.ExecuteStored(connStringFlash, "NEPS_FLASH_CREATE_STRPRD", CommandType.StoredProcedure, oraPrm2, false);
            System.Diagnostics.Debug.WriteLine("NEPS_FLASH_CREATE_STRPRD result is " + result2);*/

            return Json(new
            {
                result = res
            }, JsonRequestBehavior.AllowGet);
        }
        // endRegion

        [HttpPost]
        public ActionResult GraniteApproval(string targetJob)
        {
            System.Diagnostics.Debug.WriteLine(targetJob + "-!!!!!-");
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            string res = "";
            string username = "";
            //int webser = 0;
            int result = 0;

            GraniteOSP.Error errorObject = null;
            GraniteISP.Error errorObj = null;

            if (!targetJob.Contains("CIVIL"))
            {
                using (Entities ctxData = new Entities())
                {


                    var CariOSP = (from p in ctxData.G3E_JOB
                                   where p.G3E_IDENTIFIER == targetJob
                                   select p).Count();

                    var CariISP = (from p in ctxData.WV_ISP_JOB
                                   where p.G3E_IDENTIFIER == targetJob
                                   select p).Count();

                    bool callProc = myWebService.CallProcedureGRN(targetJob);

                    System.Diagnostics.Debug.WriteLine(" isp aaaaa");
                    System.Diagnostics.Debug.WriteLine(callProc);

                    if (CariOSP > 0)
                    {
                        var dataOSP = (from p in ctxData.G3E_JOB
                                       where p.G3E_IDENTIFIER == targetJob
                                       select p).Single();
                        username = dataOSP.G3E_OWNER;

                        neps_procman_grn_ospClient Grn = new neps_procman_grn_ospClient();
                        result = Grn.CreateProcess(out errorObject, targetJob, username);


                        //add Granite service

                        //ExampleAPIProxy.ExampleWebMethod("NEPSLoadPathConsumer", username, targetJob);

                        /*****************************TRY USING WEB REFERENCE*******************************************/
                        //var GrnData = (from a in ctxData.WV_FDP_LOAD_SITE
                        //               where a.SCHEME_NAME == targetJob
                        //               select a);
                        //foreach (var grn in GrnData)
                        //{
                        //    NEPSLoadSiteRequestSetSite graniteSite = new NEPSLoadSiteRequestSetSite();
                        //    graniteSite.SiteRegion = grn.SITEREGION;
                        //    graniteSite.SiteState = grn.SITESTATE;
                        //    graniteSite.SitePTT = grn.SITEPTT;
                        //    graniteSite.SiteExchange = grn.SITEEXC;
                        //    graniteSite.SiteName = grn.SITENAME;
                        //    graniteSite.SiteDesc = grn.SITEDESC;
                        //    graniteSite.SiteType = grn.SITETYPE;

                        //    NEPSLoadSiteRequestSiteAddress graniteSiteAdd = new NEPSLoadSiteRequestSiteAddress();
                        //    graniteSiteAdd.SiteFloor = grn.SITEFLOOR;
                        //    graniteSiteAdd.SiteNo = grn.SITENO;
                        //    graniteSiteAdd.SiteBuilding = grn.SITEBUILDING;
                        //    graniteSiteAdd.AddressStreet = grn.ADDRESSSTREET;
                        //    graniteSiteAdd.AddressCity = grn.ADDRESSCITY;
                        //    graniteSiteAdd.AddressState = grn.ADDRESSSTATE;
                        //    graniteSiteAdd.Postcode = grn.POSTCODE;
                        //    graniteSiteAdd.StreetType = grn.STREETTYPE;
                        //    graniteSiteAdd.County = grn.COUNTY;
                        //    graniteSiteAdd.Country = grn.COUNTRY;
                        //    graniteSiteAdd.SiteComments = grn.COMMENTS;

                        //    NEPSLoadSiteRequestSiteDpUDA graniteSiteUDA = new NEPSLoadSiteRequestSiteDpUDA();
                        //    graniteSiteUDA.EquipmentLocation = grn.EQUIPMENTLOCATION;
                        //    graniteSiteUDA.CablingType = grn.CABLINGTYPE;
                        //    graniteSiteUDA.CopperOwnByTM = grn.COPPEROWNBYTM;
                        //    graniteSiteUDA.FiberToPremiseExist = grn.FIBERTOPREMISEEXIST;
                        //    graniteSiteUDA.TotalServiceBondary = grn.TOTALSERVICEBOUNDARY.ToString();
                        //    graniteSiteUDA.AccessRestrictions = grn.ACCESSRESTRICTION;
                        //    graniteSiteUDA.Contact = grn.CONTACT;
                        //    graniteSiteUDA.Comments = grn.COMMENTS;

                        //    NEPSLoadSiteResponseListAddress ListAddress;
                        //    NEPSLoadSiteResponseListSiteUDA ListSiteUDA;
                        //    NEPSLoadSiteResponseCallstatus Callstatus;
                        //    NEPSLoadSiteResponseSiteList results;
                        //    try
                        //    {
                        //        NEPSLoadSiteService WS = new NEPSLoadSiteService();
                        //        results = WS.NEPSLoadSite("NEPSLoadSite", "a", graniteSite, graniteSiteAdd, graniteSiteUDA, out ListAddress, out ListSiteUDA, out Callstatus);

                        //        // System.Diagnostics.Debug.WriteLine(Callstatus.Status);
                        //    }
                        //    catch (SystemException e)
                        //    {
                        //        System.Diagnostics.Debug.WriteLine(e);
                        //    }

                        //}

                    }
                    if (CariISP > 0)
                    {
                        //Entities9 nepsbiDB = new Entities9();
                        //var dataISP = (from p in ctxData.WV_ISP_JOB
                        //               where p.G3E_IDENTIFIER == targetJob
                        //               select p).Single();

                        //var maxProcValue = nepsbiDB.BI_PROCESS_ISP21.Max(x => x.PROC_ID);
                        //int v_proc_id = Convert.ToInt32(maxProcValue) + 1;

                        //BI_PROCESS_ISP2 v_proc = new BI_PROCESS_ISP2();
                        //v_proc.PROC_ID = v_proc_id;
                        //v_proc.NEPS_JOB_ID = targetJob;
                        //v_proc.STATUS = "NEW";
                        //v_proc.TIME_CREATED = System.DateTime.Now;
                        //v_proc.CLASS_NAME = "EdgeFrontier.ISP.Granite";
                        //nepsbiDB.BI_PROCESS_ISP21.AddObject(v_proc);
                        //nepsbiDB.SaveChanges();
                        var dataISP = (from p in ctxData.WV_ISP_JOB
                                       where p.G3E_IDENTIFIER == targetJob
                                       select p).Single();

                        username = dataISP.G3E_OWNER;
                        neps_procman_grn_ispClient Grn = new neps_procman_grn_ispClient();
                        result = Grn.CreateProcess(out errorObj, targetJob, username);

                    }

                    myWebService.GRNAutoApproval(targetJob);

                    if (result == 0)
                    {
                        if (CariOSP > 0)
                        {
                            myWebService.AutoApproval(targetJob);
                        }
                        if (CariISP > 0)
                        {
                            myWebService.ISPAutoApproval(targetJob);
                        }
                    }


                }
            }

            Entities7 neps = new Entities7();
            int count = (from p in neps.VW_MANHOLEDUCT_JOB
                         where (p.JOB_ID.Contains(targetJob))
                         select p).Count();

            if (count > 0)
            {
                Entities9 nepsbiDB = new Entities9();
                var maxProcValue = nepsbiDB.BI_PROCESS_MHDUCT.Max(x => x.PROC_ID);
                int v_proc_id = Convert.ToInt32(maxProcValue) + 1;

                BI_PROCESS_MHDUCT v_proc = new BI_PROCESS_MHDUCT();
                v_proc.PROC_ID = v_proc_id;
                v_proc.NEPS_JOB_ID = targetJob;
                v_proc.STATUS = "NEW";
                v_proc.TIME_CREATED = System.DateTime.Now;
                nepsbiDB.BI_PROCESS_MHDUCT.AddObject(v_proc);
                nepsbiDB.SaveChanges();
            }

            //add for manhole-duct integration


            //MailMessage msg = new MailMessage();
            //msg.IsBodyHtml = true;
            //msg.From = new MailAddress("dev-noreply@tm.com", "NEPS Webview.");
            //msg.To.Add("haidarsukur@gmail.com");
            //msg.Subject = "TESTING";
            //msg.Body = "TEST";
            //SmtpClient emailClient = new SmtpClient("smtp.tm.com.my");
            //System.Net.NetworkCredential SMTPUserInfo = new System.Net.NetworkCredential("neps@tm.com.my", "nepsadmin");
            //emailClient.UseDefaultCredentials = false;
            //emailClient.Port = 25;
            //emailClient.Credentials = SMTPUserInfo;
            //emailClient.Send(msg);

            return Json(new
            {
                result = res
            }, JsonRequestBehavior.AllowGet);
        }



        //void WS_NEPSLoadSiteCompleted(object sender, NEPSLoadSiteCompletedEventArgs e)
        //{
        //    throw new NotImplementedException();
        //}

        //FATIHIN 26012019
        [HttpPost]
        public ActionResult NISApproval(string targetJob, string type)
        {
            System.Diagnostics.Debug.WriteLine(targetJob + "-!!!!!-");
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            string res1 = "";
            string res2 = "";

            using (Entities ctxData = new Entities())
            {
                var OSP = (from p in ctxData.G3E_JOB
                           where p.G3E_IDENTIFIER == targetJob
                           select p).Count();

                var ISP = (from p in ctxData.WV_ISP_JOB
                           where p.SCHEME_NAME == targetJob
                           select p).Count();


                if (OSP > 0 && type.Contains("OSP"))
                {
                    var dataOSP = (from p in ctxData.G3E_JOB
                                   where p.G3E_IDENTIFIER == targetJob
                                   select p).Single();
                    if (dataOSP.JOB_TYPE == "Civil")
                    {
                        var queryNET = (from d in ctxData.GC_NETELEM
                                        where d.JOB_ID == targetJob && (d.FEATURE_STATE != "ASB" || d.FEATURE_STATE != "UC")
                                        select d).Count();
                        int checkFeat = queryNET;
                        if (checkFeat == 0)
                        {
                            res1 = myWebService.AutoApproval(targetJob);
                        }
                        else
                        {
                            res1 = "fail";
                        }

                    }
                    else
                    {
                        //INSERT BI_PROCESS
                        EntitiesNetworkElement nepsbiDB = new EntitiesNetworkElement();
                        var maxProcValue = nepsbiDB.BI_PROCESS.Max(x => x.PROC_ID);
                        int v_proc_id = Convert.ToInt32(maxProcValue) + 1;

                        BI_PROCESS v_proc = new BI_PROCESS();
                        v_proc.PROC_ID = v_proc_id;
                        v_proc.NEPS_JOB_ID = targetJob;
                        v_proc.STATUS = "NEW";
                        v_proc.TIME_CREATED = System.DateTime.Now;
                        v_proc.CLASS_NAME = "EdgeFrontier.OSP.NIS";
                        nepsbiDB.BI_PROCESS.AddObject(v_proc);
                        nepsbiDB.SaveChanges();
                    }
                }

                if (ISP > 0 && type.Contains("ISP"))
                {
                    //INSERT BI_PROCESS_ISP
                    EntitiesNetworkElement nepsbiDB = new EntitiesNetworkElement();
                    var maxProcValue = nepsbiDB.BI_PROCESS_ISP.Max(x => x.PROC_ID);
                    int v_proc_id = Convert.ToInt32(maxProcValue) + 1;

                    BI_PROCESS_ISP v_proc = new BI_PROCESS_ISP();
                    v_proc.PROC_ID = v_proc_id;
                    v_proc.NEPS_JOB_ID = targetJob;
                    v_proc.STATUS = "NEW";
                    v_proc.TIME_CREATED = System.DateTime.Now;
                    v_proc.CLASS_NAME = "EdgeFrontier.ISP.NIS";
                    nepsbiDB.BI_PROCESS_ISP.AddObject(v_proc);
                    nepsbiDB.SaveChanges();

                    var CariISP2 = (from p in ctxData.WV_ISP_JOB
                                    where p.SCHEME_NAME == targetJob
                                    select p).Single();

                    using (EntitiesNetworkElement chckData = new EntitiesNetworkElement())
                    {
                        var checkDataISP = (from p in chckData.BI_PROCESS_ISP
                                            where p.NEPS_JOB_ID == targetJob && p.STATUS == "SUCCESS"
                                            select p).Count();
                        if (checkDataISP > 0)
                        {
                            res1 = myWebService.ISPAutoApproval(CariISP2.G3E_IDENTIFIER);

                        }
                        else
                            res1 = "fail";
                    }
                }
            }
            //Fatihin 17042019
           /* string result = "";
            Tools tool = new Tools();
            OracleParameter[] oraPrm2;
            oraPrm2 = new OracleParameter[2];
            oraPrm2[0] = new OracleParameter("schemeName", OracleDbType.Varchar2);
            oraPrm2[0].Value = targetJob;
            oraPrm2[1] = new OracleParameter("SEND_TYPE", OracleDbType.Varchar2);
            oraPrm2[1].Value = "ASB";
            result = tool.ExecuteStored(connStringFlash, "NEPS_FLASH_CREATE_STRPRD", CommandType.StoredProcedure, oraPrm2, false);
            System.Diagnostics.Debug.WriteLine("NEPS_FLASH_CREATE_STRPRD result is " + result);*/

            return Json(new
            {
                result1 = res1,
                result2 = res2

            }, JsonRequestBehavior.AllowGet);

        }

        [HttpPost]
        public ActionResult SendSwift(string targetJob)
        {
            bool success;

            Tools tool = new Tools();
            int v_proc_id = 0;
            using (Entities9 ctxCheckProcId = new Entities9())
            {
                var maxValue = ctxCheckProcId.BI_PROC_SWIFT.Max(x => x.PROC_ID);
                v_proc_id = Convert.ToInt32(maxValue) + 1;
            }
            string sqlStr = "insert into NEPSBI.BI_PROC_SWIFT(PROC_ID, STATUS, TIME_CREATED, NEPS_JOB_ID) values";
            OracleParameter[] oraPrm = new OracleParameter[4];

            oraPrm[0] = new OracleParameter("v_PROC_ID", OracleDbType.Varchar2);
            oraPrm[0].Value = v_proc_id;
            oraPrm[1] = new OracleParameter("v_STATUS", OracleDbType.Varchar2);
            oraPrm[1].Value = "NEW";
            oraPrm[2] = new OracleParameter("v_TIME_CREATED", OracleDbType.Date);
            oraPrm[2].Value = System.DateTime.Now;
            oraPrm[3] = new OracleParameter("v_NEPS_JOB_ID", OracleDbType.Varchar2);
            oraPrm[3].Value = targetJob;

            string result = tool.ExecuteStored(connString, sqlStr, CommandType.StoredProcedure, oraPrm, false);
            success = true;

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult AssetGems(string targetJob)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            success = myWebService.assetGems(targetJob);
            System.Diagnostics.Debug.WriteLine("ssssss : " + success);
            success = true;
            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult CheckAssetGems(string id)
        {
            string outputCheckAssetGEMS = "";

            using (Entities7 ctxCheckAssetGEMS = new Entities7())
            {
                //var queryCheckAssetGEMS = from a in ctxCheckAssetGEMS.VW_ASSET_TO_GEMS
                //                          where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper()
                //                          select new
                //                          {
                //                              a.G3E_FNO,
                //                              a.G3E_FID,
                //                              a.PROJECT_NO,
                //                              a.WBS_NUM,
                //                              a.ADDRESS1,
                //                              a.ADDRESS2,
                //                              a.ADDRESS3
                //                          };

                //if (queryCheckAssetGEMS.Count() > 0)
                //{
                //    foreach (var a in queryCheckAssetGEMS)
                //    {
                //        outputCheckAssetGEMS += "[" + a.G3E_FNO + "|" + a.G3E_FID + "|" + a.PROJECT_NO + "|" + a.WBS_NUM + "|" + a.ADDRESS1 + "|" + a.ADDRESS2 + "|" + a.ADDRESS3 + "|";
                //    }
                //}
            }

            return Json(new
            {
                CheckAsset = outputCheckAssetGEMS

            }, JsonRequestBehavior.AllowGet);

        }

        [HttpPost]
        public ActionResult ResentApproval(string targetJob, string targetDesc)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            string res = "";

            res = myWebService.ResendApproval(targetJob, targetDesc);

            //MailMessage msg = new MailMessage();
            //msg.IsBodyHtml = true;
            //msg.From = new MailAddress("dev-noreply@tm.com", "NEPS Webview.");
            //msg.To.Add("haidarsukur@gmail.com");
            //msg.Subject = "TESTING";
            //msg.Body = "TEST";
            //SmtpClient emailClient = new SmtpClient("smtp.tm.com.my");
            //System.Net.NetworkCredential SMTPUserInfo = new System.Net.NetworkCredential("neps@tm.com.my", "nepsadmin");
            //emailClient.UseDefaultCredentials = false;
            //emailClient.Port = 25;
            //emailClient.Credentials = SMTPUserInfo;
            //emailClient.Send(msg);

            return Json(new
            {
                result = res
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        //fatihin 19042018
        public ActionResult ISPAutoApproval(string targetJob, string type)
        {
            System.Diagnostics.Debug.WriteLine(targetJob + "-!!!!!-");
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            string res = "";
            string res1 = "";
            string res2 = "";
            int chckISP = 0;

            Error_ISP errorObj = null;
            OSPHandoverServices.Error errorOSP = null;
            using (Entities ctxData = new Entities())
            {
                var CariISP = (from p in ctxData.WV_ISP_JOB
                               where p.G3E_IDENTIFIER == targetJob
                               select p).Count();

                var CariDataISP = (from p in ctxData.WV_ISP_JOB
                                   where p.G3E_IDENTIFIER == targetJob
                                   select p).Single();

                var CariOSP = (from p in ctxData.G3E_JOB
                               where p.G3E_IDENTIFIER == CariDataISP.SCHEME_NAME
                               select p).Count();

                int checkFeat;

                var queryNET = (from d in ctxData.GC_NETELEM
                                where d.JOB_ID == targetJob && d.FEATURE_STATE != "ASB"
                                select d).Count();

                checkFeat = queryNET;
                //fatihin 19042018
                if (CariOSP > 0 && type.Contains("OSP"))
                {
                    var CariOSP2 = (from p in ctxData.G3E_JOB
                                    where p.G3E_IDENTIFIER == CariDataISP.SCHEME_NAME
                                    select p).Single();
                    if (checkFeat == 0)
                    {
                        neps_procmanClient testWSOSP = new neps_procmanClient();
                        testWSOSP.CreateProcess(out errorOSP, targetJob);

                        using (EntitiesNetworkElement chckData = new EntitiesNetworkElement())
                        {
                            var checkDataOSP = (from p in chckData.BI_PROCESS
                                                where p.NEPS_JOB_ID == targetJob && p.STATUS == "SUCCESS"
                                                select p).Count();
                            if (checkDataOSP > 0)
                            {
                                res1 = myWebService.AutoApproval(CariOSP2.G3E_IDENTIFIER);
                            }
                            else
                                res1 = "fail";
                        }
                    }
                    else
                    { res1 = "fail"; }

                }
                //fatihin 19042018
                if (CariISP > 0 && type.Contains("ISP"))
                {
                    // endRegion
                    neps_procman_ispClient testWS = new neps_procman_ispClient();
                    testWS.CreateProcess(out errorObj, targetJob);

                    System.Diagnostics.Debug.WriteLine(testWS + "-HANDOVER ISP-");
                    using (EntitiesNetworkElement chckData = new EntitiesNetworkElement())
                    {
                        var checkDataISP = (from p in chckData.BI_PROCESS_ISP
                                            where p.NEPS_JOB_ID == targetJob && p.STATUS == "SUCCESS"
                                            select p).Count();
                        if (checkDataISP > 0)
                        {
                            if (CariDataISP.JOB_TYPE == "ISP" || CariDataISP.JOB_TYPE == "Others")
                            {
                                res2 = myWebService.ISPAutoApproval(targetJob);
                            }
                            else
                            {
                                if (checkFeat == 0)
                                {
                                    res2 = myWebService.ISPAutoApproval(targetJob);
                                }
                                else
                                { res1 = "fail"; }
                            }
                        }
                        else
                            res2 = "fail";

                    }
                }
            }
            //string res = myWebService.ISPAutoApproval(targetJob);
            if (res2 == "fail")
            {
                res = res1;
            }
            else
            {
                res = res2;
            }
            return Json(new
            {
                result = res
            }, JsonRequestBehavior.AllowGet);
        }

        // state: ASB, Status: COMPLETED

        [HttpPost]
        public ActionResult SendNotification(string subject, string message,
            string from, string recipient, string targetUser, string jobId, string schemeName)
        {
            string result = "";

            System.Diagnostics.Debug.WriteLine("sending...");
            System.Diagnostics.Debug.WriteLine(jobId);
            System.Diagnostics.Debug.WriteLine(schemeName);
            System.Diagnostics.Debug.WriteLine(subject);
            System.Diagnostics.Debug.WriteLine(message);
            System.Diagnostics.Debug.WriteLine(recipient);
            System.Diagnostics.Debug.WriteLine("targetUser: " + targetUser);
            System.Diagnostics.Debug.WriteLine("############");

            //string[] targetUsr = targetUser.Split('|');
            //for (int i = 0; i < targetUsr.Length - 1; i++)
            //{
            //    WebView.WebService._base myWebService;
            //    myWebService = new WebService._base();

            //    myWebService.AddApproval(schemeName,targetUsr[i], jobId);
            //}
            try
            {

                TcpClient tcpclient = new TcpClient(); // create an instance of TcpClient
                tcpclient.Connect("SMSGVS34.tm.my", 110); // HOST NAME POP SERVER and gmail uses port number 995 for POP 
                System.Net.Security.SslStream sslstream = new SslStream(tcpclient.GetStream()); // This is Secure Stream // opened the connection between client and POP Server
                sslstream.AuthenticateAsClient("SMSGVS34.tm.my"); // authenticate as client 
                //bool flag = sslstream.IsAuthenticated; // check flag 
                System.IO.StreamWriter sw = new StreamWriter(sslstream); // Asssigned the writer to stream
                System.IO.StreamReader reader = new StreamReader(sslstream); // Assigned reader to stream
                sw.WriteLine("USER neps"); // refer POP rfc command, there very few around 6-9 command
                sw.Flush(); // sent to server
                sw.WriteLine("PASS nepsadmin");
                sw.Flush();
                sw.WriteLine("RETR 1"); // this will retrive your first email 
                sw.Flush();
                sw.WriteLine("Quit "); // close the connection
                sw.Flush();

                string str = string.Empty;
                string strTemp = string.Empty;

                while ((strTemp = reader.ReadLine()) != null)
                {
                    if (strTemp == ".") // find the . character in line
                    {
                        break;
                    }
                    if (strTemp.IndexOf("-ERR") != -1)
                    {
                        break;
                    }
                    str += strTemp;
                }

                Response.Write(str);

                Response.Write("<BR>" + "Congratulation.. ....!!! You read your first gmail email ");
            }

            catch (Exception ex)
            {
                Response.Write(ex.Message);
            }

            string[] arr = recipient.Split('|');
            string emailList = "";
            string nameList = "";

            for (int i = 0; i < arr.Length - 1; i++)
            {
                string[] extractArr = arr[i].Split('#');

                emailList = emailList + extractArr[1] + ",";
                nameList = nameList + extractArr[0] + ",";
            }

            emailList = emailList.Substring(0, emailList.Length - 1);
            nameList = emailList.Substring(0, nameList.Length - 1);

            System.Diagnostics.Debug.WriteLine("EMAIL LIST: " + emailList);
            try
            {
                //MailMessage msg = new MailMessage();
                //msg.IsBodyHtml = true;
                //msg.From = new MailAddress("neps@tm.com.my", "NEPS");
                //msg.To.Add(emailList);
                //msg.Subject = subject;
                //msg.Body = message;
                //SmtpClient emailClient = new SmtpClient("smtp.tm.com.my");
                //System.Net.NetworkCredential SMTPUserInfo = new System.Net.NetworkCredential("tmmaster\neps", "nepsadmin");
                //emailClient.UseDefaultCredentials = false;
                //emailClient.Port = 25;
                //emailClient.Credentials = SMTPUserInfo;
                //emailClient.Send(msg);

                result = "ok";
            }
            catch (Exception ex)
            {
                result = "fail" + "|" + ex.ToString();
                System.Diagnostics.Debug.WriteLine(ex.ToString());
            }

            return Json(new
            {
                result = result
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult Handover(string id)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            System.Diagnostics.Debug.WriteLine("handover aaa");
            ArrayList lst = new ArrayList();
            List<string> innerLst = new List<string>();
            innerLst.Add("btn-add");
            innerLst.Add("#AvailableUsers");
            innerLst.Add("#RequestedSelected");
            lst.Add(innerLst);

            innerLst = new List<string>();
            innerLst.Add("btn-add-fibre");
            innerLst.Add("#FibreAvailable");
            innerLst.Add("#FibreRequested");
            lst.Add(innerLst);

            ViewBag.addbutton = lst;
            string schemeType = "";
            string JobID = "";
            if (id != null)
            {
                using (Entities ctxData = new Entities())
                {
                    var Cari = (from p in ctxData.G3E_JOB
                                where p.G3E_IDENTIFIER == id
                                select p).Count();
                    if (Cari > 0)
                    {
                        var query = (from p in ctxData.G3E_JOB
                                     where p.G3E_IDENTIFIER == id
                                     select p).Single();

                        ViewBag.jobId = id;
                        JobID = id;
                        ViewBag.scheme_name = query.SCHEME_NAME;
                        schemeType = query.JOB_TYPE;
                        ViewBag.JobType = "OSP";
                    }
                    else
                    {
                        var query2 = (from p in ctxData.WV_ISP_JOB
                                      where p.G3E_IDENTIFIER == id
                                      select p).Single();

                        ViewBag.jobId = id;
                        JobID = id;
                        ViewBag.scheme_name = query2.SCHEME_NAME;
                        schemeType = query2.JOB_TYPE;
                        ViewBag.JobType = "ISP";
                    }
                }
            }

            string grpId; // check user group ---------------------------not done yet (04-May-2012)
            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_USER
                             where p.USERNAME.ToUpper() == User.Identity.Name.ToUpper() || p.USERNAME.ToLower() == User.Identity.Name.ToLower()
                             select p).Single();
                grpId = query.GROUPID;
            }
            int checkApp = 0;
            string grpApprovalId = "";
            using (Entities ctxData = new Entities())
            {
                if (grpId == "8" && schemeType != "Fiber (Equip)" && schemeType != "Copper (Equip)" && schemeType != "HSBB (Equip)")
                {
                    checkApp = (from d in ctxData.WV_USER
                                join fx in ctxData.WV_USER_APPROVE on d.USERNAME equals fx.APPROVAL_USER
                                where d.GROUPID == "15" && fx.JOB_ID == id
                                orderby d.USER_ID
                                select d).Count();
                    grpApprovalId = "15";
                }
                else if (schemeType == "Civil" || schemeType == "Fiber E/Side" || schemeType == "Fiber (Equip)" || schemeType == "E/Side" || schemeType == "D/Side" || schemeType == "Copper (Equip)" || schemeType == "HSBB E/Side" || schemeType == "HSBB D/Side" || schemeType == "HSBB (Equip)")
                {
                    if (grpId == "4" || grpId == "5" || grpId == "6" || grpId == "7")
                    {
                        checkApp = (from d in ctxData.WV_USER
                                    join fx in ctxData.WV_USER_APPROVE on d.USERNAME equals fx.APPROVAL_USER
                                    where d.GROUPID == "13" && fx.JOB_ID == id
                                    orderby d.USER_ID
                                    select d).Count();
                        grpApprovalId = "13";
                    }
                }
                else if (schemeType == "Fiber Trunk" || schemeType == "Fiber Junction")
                {
                    if (grpId == "4" || grpId == "5" || grpId == "6" || grpId == "7")
                    {
                        checkApp = (from d in ctxData.WV_USER
                                    join fx in ctxData.WV_USER_APPROVE on d.USERNAME equals fx.APPROVAL_USER
                                    where d.GROUPID == "14" && fx.JOB_ID == id
                                    orderby d.USER_ID
                                    select d).Count();
                        grpApprovalId = "14";
                    }
                }
                else
                {
                    checkApp = (from d in ctxData.WV_USER
                                join fx in ctxData.WV_USER_APPROVE on d.USERNAME equals fx.APPROVAL_USER
                                where d.GROUPID == "10" && fx.JOB_ID == id
                                orderby d.USER_ID
                                select d).Count();
                    grpApprovalId = "10";
                }


                if (checkApp != 0)
                {
                    HandoverModel.ViewModel model = new HandoverModel.ViewModel { AvailableUsers = myWebService.GetHandoverUser(0, 1001, id, grpApprovalId), RequestedUsers = new List<HandoverModel.user>() };
                    return View(model);
                }
                else
                {
                    HandoverModel.ViewModel model = new HandoverModel.ViewModel { AvailableUsers = myWebService.GetHandoverUser(0, 100, id, grpApprovalId), RequestedUsers = new List<HandoverModel.user>() };
                    return View(model);
                }
            }
        }

        [HttpPost]
        public ActionResult Handover(string targetId, HandoverModel.ViewModel m)
        {
            System.Diagnostics.Debug.WriteLine(targetId);
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            int senarai_nama = 0;
            string appusername = "";
            string grpApprovalId = "";
            List<int> list = new List<int>();
            for (int i = 0; i < m.RequestedSelected.Length; i++)
            {
                //senarai_nama = m.RequestedUsers.ToString();
                list.Add(m.RequestedSelected[i]);
                senarai_nama = m.RequestedSelected[i];
                System.Diagnostics.Debug.WriteLine(senarai_nama);

                using (Entities ctxData = new Entities())
                {
                    var query = (from d in ctxData.WV_USER
                                 where d.USER_ID == senarai_nama
                                 select d).Single();

                    appusername = query.USERNAME;
                    grpApprovalId = query.GROUPID;
                }
                Tools tool = new Tools();
                bool success = true;
                string sqlCmd = "";
                string sqlCmd2 = "";
                string sqlCmd3 = "";
                DateTime thisDay = DateTime.Now;
                using (Entities ctxData = new Entities())
                {
                    var CariOSP = (from p in ctxData.G3E_JOB
                                   where p.G3E_IDENTIFIER == targetId
                                   select p).Count();

                    var CariISP = (from p in ctxData.WV_ISP_JOB
                                   where p.G3E_IDENTIFIER == targetId
                                   select p).Count();

                    sqlCmd = "INSERT INTO WV_USER_APPROVE (JOB_ID, REMARKS, STATUS, GLOBAL_STATUS, APPROVAL_USER, CREATED_AT) VALUES ( '" + targetId + "', 'PENDING APPROVAL', 'PENDING', 'INCOMPLETE','" + appusername + "', '" + thisDay.ToString("dd/MMMM/yyyy") + "')";
                    if (CariOSP > 0)
                    {
                        sqlCmd2 = "UPDATE G3E_JOB SET JOB_STATE= 'PENDING APP' WHERE G3E_IDENTIFIER ='" + targetId + "'";
                        using (Entities ctxData2 = new Entities())
                        {
                            tool.ExecuteSql(ctxData2, sqlCmd2);

                            var query = (from p in ctxData.G3E_JOB
                                         where p.G3E_IDENTIFIER == targetId
                                         select p).Single();

                            ViewBag.jobId = targetId;
                            ViewBag.scheme_name = query.SCHEME_NAME;
                            ViewBag.JobType = "OSP";
                        }
                    }
                    if (CariISP > 0)
                    {
                        sqlCmd3 = "UPDATE WV_ISP_JOB SET JOB_STATE= 'PENDING APP' WHERE G3E_IDENTIFIER ='" + targetId + "'";
                        using (Entities ctxData2 = new Entities())
                        {
                            tool.ExecuteSql(ctxData2, sqlCmd3);

                            var query = (from p in ctxData.WV_ISP_JOB
                                         where p.G3E_IDENTIFIER == targetId
                                         select p).Single();

                            ViewBag.jobId = targetId;
                            ViewBag.scheme_name = query.SCHEME_NAME;
                            ViewBag.JobType = "ISP";
                        }
                    }
                }
                using (Entities ctxData = new Entities())
                {
                    success = tool.ExecuteSql(ctxData, sqlCmd);
                }
                System.Diagnostics.Debug.WriteLine("ok X?" + success);
            }

            HandoverModel.ViewModel model = new HandoverModel.ViewModel { AvailableUsers = myWebService.GetHandoverUser(0, 1001, targetId, grpApprovalId), RequestedUsers = new List<HandoverModel.user>(), ListSend = list };
            return View(model);
        }

        public ActionResult FileList(string id) // get list file by Job No
        {
            //string path = AppDomain.CurrentDomain.BaseDirectory + "App_Data/uploads/job"+id+"/";
            string path = Server.MapPath("~/App_Data/uploads/job" + id + "/");
            DirectoryInfo salesFTPDirectory = null;
            FileInfo[] filesupload = null;
            var result = "";
            var checkJob = "";
            bool folderCreated = false;

            using (Entities ctxData = new Entities())
            {
                var queryOSP = (from p in ctxData.G3E_JOB
                                where p.G3E_IDENTIFIER == id
                                select p).Count();
                if (queryOSP > 0)
                {
                    var query = (from p in ctxData.G3E_JOB
                                 where p.G3E_IDENTIFIER == id
                                 select p).Single();
                    checkJob = query.SCHEME_NAME;
                }
                else
                {
                    var query2 = (from p in ctxData.WV_ISP_JOB
                                  where p.G3E_IDENTIFIER == id
                                  select p).Single();

                    checkJob = query2.SCHEME_NAME;
                }
                ViewBag.SchemeName = checkJob;

            }

            if (Directory.Exists(path)) //check for existing directory
            {
                string salesFTPPath = Server.MapPath("~/App_Data/uploads/job" + id + "/");
                salesFTPDirectory = new DirectoryInfo(salesFTPPath);
                filesupload = salesFTPDirectory.GetFiles();
                Session["pathId"] = id;

                filesupload = filesupload.OrderBy(f => f.Name).ToArray();
                ViewBag.AttachLink = id;

                return View(filesupload);
            }
            else
            {
                //create job directory
                string newPath = System.IO.Path.Combine(Server.MapPath("~/App_Data/uploads/job" + id));
                System.IO.Directory.CreateDirectory(newPath);
                System.Diagnostics.Debug.WriteLine(newPath);

                folderCreated = true;
                Session["pathId"] = id;
                ViewBag.AttachLink = id;
                return View();
            }
        }

        public ActionResult FileListRedmark(string id) // get list file by Job No
        {
            //string path = AppDomain.CurrentDomain.BaseDirectory + "App_Data/uploads/job"+id+"/";
            string path = Server.MapPath("~/App_Data/uploads/job" + id + "/");
            DirectoryInfo salesFTPDirectory = null;
            FileInfo[] filesupload = null;
            var result = "";
            string grpId = "";
            bool folderCreated = false;

            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_NONNETWORK_JOB
                             where p.G3E_IDENTIFIER == id
                             select p).Single();

                ViewBag.STATUSSEMASA = query.JOB_STATE;
                ViewBag.DESCRIPTION = query.G3E_DESCRIPTION;
                ViewBag.REMARKS = query.SCH_DESC1;

                var query2 = (from p in ctxData.WV_USER
                              where p.USERNAME == User.Identity.Name.ToUpper()
                              select p).Single();
                grpId = query2.GROUPID;
            }
            List<SelectListItem> list1 = new List<SelectListItem>();
            if (grpId == "2" || grpId == "16" || grpId == "17")
            {
                list1.Add(new SelectListItem() { Text = ViewBag.STATUSSEMASA, Value = ViewBag.STATUSSEMASA, Selected = true });
                list1.Add(new SelectListItem() { Text = "REJECT", Value = "REJECT" });
                list1.Add(new SelectListItem() { Text = "ACCEPTED", Value = "ACCEPTED" });
                list1.Add(new SelectListItem() { Text = "COMPLETED", Value = "COMPLETED" });
            }
            else if (grpId == "9")
            {
                list1.Add(new SelectListItem() { Text = ViewBag.STATUSSEMASA, Value = ViewBag.STATUSSEMASA, Selected = true });
                list1.Add(new SelectListItem() { Text = "RESENT REDMARK", Value = "PENDING" });
            }

            ViewBag.UPDATESTATUS = list1;

            if (Directory.Exists(path)) //check for existing directory
            {
                string salesFTPPath = Server.MapPath("~/App_Data/uploads/job" + id + "/");
                salesFTPDirectory = new DirectoryInfo(salesFTPPath);
                filesupload = salesFTPDirectory.GetFiles();
                Session["pathId"] = id;

                filesupload = filesupload.OrderBy(f => f.Name).ToArray();
                ViewBag.AttachLink = id;

                return View(filesupload);
            }
            else
            {
                //create job directory
                string newPath = System.IO.Path.Combine(Server.MapPath("~/App_Data/uploads"), "job" + id);
                System.IO.Directory.CreateDirectory(newPath);
                System.Diagnostics.Debug.WriteLine(newPath);

                folderCreated = true;
                Session["pathId"] = id;
                ViewBag.AttachLink = id;
                return View();
            }
        }

        // region Mubin - CR14-20180330
        public ActionResult ViewHandover(string id, int? page, string type) // get list file by Job No
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            WebService._base.OSPJob jobs = new WebService._base.OSPJob();

            jobs = myWebService.GetHandoverJob(id, 0, 100, type);

            string input = "\\\\adsvr";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;
            ViewBag.SchemeName = id;
            ViewBag.searchKey = id;
            ViewBag.type = type;

            int pageSize = 5;
            int pageNumber = (page ?? 1);
            return View(jobs.JobList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult ViewHandoverISP(string id, int? page, string type) // get list file by Job No
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            WebService._base.OSPJob jobs = new WebService._base.OSPJob();

            jobs = myWebService.GetHandoverISP(id, 0, 100, type);

            string input = "\\\\adsvr";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;
            ViewBag.SchemeName = id;
            ViewBag.searchKey = id;
            ViewBag.type = type;

            int pageSize = 5;
            int pageNumber = (page ?? 1);
            return View(jobs.JobList.ToPagedList(pageNumber, pageSize));
        }
        // endRegion

        public ActionResult Download(string id, string filename) //download file
        {
            string file = Server.MapPath("~/App_Data/uploads/job" + id + "/");
            var filePath = Path.Combine(file, filename);

            return File(filePath, "application/octet-stream", filename);
        }

        [HttpPost]
        public ActionResult NewAttachment(HttpPostedFileBase files) // add file by folder
        {

            string extension = Path.GetExtension(files.FileName); // get extension file
            string schemeName = "";
            //string result = "";
            //WebView.WebService._base myWebService;
            //myWebService = new WebService._base();

            //bool selected = false;
            if (files != null && files.ContentLength > 0) // filter file
            {
                var fileName = Path.GetFileName(files.FileName);
                var path = Path.Combine(Server.MapPath("~/App_Data/uploads/job" + Session["pathId"] + "/"), fileName);
                string scheme = Session["pathId"].ToString();
                if (extension == ".xml")
                {
                    try
                    {
                        DateTime thisDay = DateTime.Now;

                        using (Entities ctxData = new Entities())
                        {
                            var query = (from p in ctxData.G3E_JOB
                                         where p.G3E_IDENTIFIER == scheme
                                         select p).Single();

                            schemeName = query.SCHEME_NAME;
                        }
                        if (schemeName == null)
                        {
                            using (Entities ctxData = new Entities())
                            {
                                var query = (from p in ctxData.WV_ISP_JOB
                                             where p.G3E_IDENTIFIER == scheme
                                             select p).Single();

                                schemeName = query.SCHEME_NAME;
                            }
                        }

                        var path2 = Path.Combine(Server.MapPath("~/App_Data/uploads/job" + Session["pathId"] + "/" + schemeName + thisDay.ToString("MMdd") + "-" + thisDay.ToString("HH") + "" + thisDay.ToString("mm") + extension));
                        //var path3 = EXEC_NAME + "_" + CABINET + "_" + thisDay.ToString("yyyyMMdd") + "_" + thisDay.ToString("HH") + "" + thisDay.ToString("mm") + extension;

                        //result = myWebService.insertXml(path3, scheme, User.Identity.Name);

                        files.SaveAs(path2);

                        //StreamReader fileStream = new StreamReader(path);
                        //string fileContent = fileStream.ReadToEnd();
                        //fileStream.Close();

                        //StreamWriter ansiWriter = new StreamWriter(path, false);
                        //ansiWriter.Write(fileContent, Encoding.Default);
                        //ansiWriter.Close();

                        //System.Diagnostics.Debug.WriteLine("Blah");

                        //result = myWebService.insertXml(path, scheme);
                    }
                    catch (Exception ex)
                    {
                        // Log error.
                        System.Diagnostics.Debug.WriteLine(ex.ToString());
                    }

                }
                else
                {
                    files.SaveAs(path);
                }

            }

            return RedirectToAction("FileList/" + Session["pathId"]);
        }

        [HttpPost]
        public ActionResult NewAttachmentRedmark(HttpPostedFileBase files) // add file by folder
        {
            string extension = Path.GetExtension(files.FileName); // get extension file
            string schemeName = "";

            bool selected = false;
            if (files != null && files.ContentLength > 0) // filter file
            {
                var fileName = Path.GetFileName(files.FileName);
                var path = Path.Combine(Server.MapPath("~/App_Data/uploads/job" + Session["pathId"] + "/"), fileName);
                string scheme = Session["pathId"].ToString();
                if (extension == ".xml")
                {
                    try
                    {
                        DateTime thisDay = DateTime.Now;

                        using (Entities ctxData = new Entities())
                        {
                            var query = (from p in ctxData.WV_NONNETWORK_JOB
                                         where p.G3E_IDENTIFIER == scheme
                                         select p).Single();

                            schemeName = query.SCHEME_NAME;
                        }

                        System.Diagnostics.Debug.WriteLine(schemeName);
                        schemeName = schemeName.Replace("/", "");
                        schemeName = schemeName.Replace("*", "");
                        var path2 = Path.Combine(Server.MapPath("~/App_Data/uploads/job" + Session["pathId"] + "/" + schemeName + "-" + thisDay.ToString("HH") + "" + thisDay.ToString("mm") + extension));
                        files.SaveAs(path2);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(ex.ToString());
                    }
                }
                else
                {
                    files.SaveAs(path);
                }

            }

            return RedirectToAction("FileListRedmark/" + Session["pathId"]);
        }

        [HttpPost]
        public ActionResult UpdatedStatusRedmark(String id, String status, String description, String remarks) // add file by folder
        {

            System.Diagnostics.Debug.WriteLine(status);
            string success = "";
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.UpdateStatusRedmark(id, status, description, remarks);

            string ownerUsername = "";
            string ownerEmail = "";
            string schemeName = "";
            string jdescription = "";
            string jremarks = "";
            string andEmail = "";
            string andTel = "";
            string namaFailUpload = "";
            string EmailListStr = "";
            string jobNo = "";
            string andName;

            using (Entities ctxData = new Entities())
            {
                var query = (from d in ctxData.WV_USER
                             join fx in ctxData.WV_NONNETWORK_JOB on d.USERNAME equals fx.G3E_OWNER
                             where fx.G3E_IDENTIFIER == id
                             select d).Single();
                ownerUsername = query.USERNAME;
                ownerEmail = query.EMAIL;

                var queryJob = (from d in ctxData.WV_NONNETWORK_JOB
                                where d.G3E_IDENTIFIER == id
                                select d).Single();
                schemeName = queryJob.SCHEME_NAME;
                jdescription = queryJob.G3E_DESCRIPTION;
                jremarks = queryJob.SCH_DESC1;
                jobNo = queryJob.G3E_IDENTIFIER;

                var queryAND = (from d in ctxData.WV_USER
                                where d.USERNAME.ToUpper() == User.Identity.Name.ToUpper() || d.USERNAME.ToLower() == User.Identity.Name.ToLower()
                                select d).Single();

                andEmail = queryAND.EMAIL;
                andTel = queryAND.NO_TEL;
                andName = queryAND.FULL_NAME;

                var querylist = (from up in ctxData.WV_USER
                                 where up.GROUPID == "16" && up.PTT_STATE.Trim() == queryAND.PTT_STATE.Trim()
                                 select up);

                int i = 0;

                EmailListStr = ownerEmail + ",";

                foreach (var emailAND in querylist)
                {
                    if (i++ > 0)
                        EmailListStr += emailAND.EMAIL + ",";
                }

                EmailListStr = EmailListStr.TrimEnd(',');

                System.Diagnostics.Debug.WriteLine(EmailListStr);
            }

            string salesFTPPath = Server.MapPath("~/App_Data/uploads/job" + id + "/");
            DirectoryInfo namaFail = new DirectoryInfo(salesFTPPath);
            FileInfo[] fi = namaFail.GetFiles("*.xml");
            foreach (FileInfo namaFailRedMark in fi)
            {
                namaFailUpload = namaFailRedMark.Name;
                System.Diagnostics.Debug.WriteLine(namaFailRedMark.Name);
            }

            DateTime thisDay = DateTime.Now;
            if (status == "REJECT")
            {
                try
                {
                    MailMessage msg = new MailMessage();
                    msg.IsBodyHtml = true;
                    msg.From = new MailAddress("neps@tm.com.my", "NEPS");
                    msg.To.Add(EmailListStr);
                    msg.Subject = "Rejected Redmark File for Scheme " + schemeName;
                    msg.Body = "<h1>FILES DETAILS</h1>SCHEME NAME	: " + schemeName + "<br/><br/>JOB NO : " + jobNo + "<br/><br/>DESCRIPTION : " + jdescription + "<br/><br/>REMARKS : " + jremarks + " <br/><br/>REDMARK FILE NAME: " + namaFailUpload;
                    msg.Body += "<br/><br/> <h1>AND DETAILS</h1> <br>";
                    msg.Body += "AND ID : " + User.Identity.Name + "<br/><br/>AND NAME	: " + andName + "<br/><br/>AND EMAIL	: " + andEmail + "<br/><br/>AND PHONE NUMBER: " + andTel + "<br/><br/>Please log in to <a href='http://10.41.101.168/'>NEPS WEBVIEW  </a>to download the file.";
                    msg.IsBodyHtml = true;
                    SmtpClient emailClient = new SmtpClient("smtp.tm.com.my", 25);
                    emailClient.UseDefaultCredentials = false;
                    emailClient.Credentials = new NetworkCredential("neps", "nepsadmin", "tmmaster");
                    emailClient.EnableSsl = false;
                    emailClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                    emailClient.Send(msg);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(ex.ToString());
                }
            }
            if (status == "PENDING")
            {
                try
                {
                    MailMessage msg = new MailMessage();
                    msg.IsBodyHtml = true;
                    msg.From = new MailAddress("neps@tm.com.my", "NEPS");
                    msg.To.Add(EmailListStr);
                    msg.Subject = "Resent Redmark File for Scheme " + schemeName + "Rejected";
                    msg.Body = "<h1>FILES DETAILS</h1>SCHEME NAME	: " + schemeName + "<br/><br/>JOB NO : " + jobNo + "<br/><br/>DESCRIPTION : " + jdescription + "<br/><br/>REMARKS : " + jremarks + " <br/><br/>REDMARK FILE NAME: " + namaFailUpload;
                    msg.Body += "<br/><br/> <h1>RNO DETAILS</h1> <br>";
                    msg.Body += "RNO ID : " + User.Identity.Name + "<br/><br/>RNO NAME	: " + andName + "<br/><br/>RNO EMAIL	: " + andEmail + "<br/><br/>RNO PHONE NUMBER: " + andTel + "<br/><br/>.Please log in to <a href='http://10.41.101.168/'>NEPS WEBVIEW  </a>to download the file.";
                    msg.IsBodyHtml = true;
                    SmtpClient emailClient = new SmtpClient("smtp.tm.com.my", 25);
                    emailClient.UseDefaultCredentials = false;
                    emailClient.Credentials = new NetworkCredential("neps", "nepsadmin", "tmmaster");
                    emailClient.EnableSsl = false;
                    emailClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                    emailClient.Send(msg);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(ex.ToString());
                }
            }
            if (status == "COMPLETED")
            {
                try
                {
                    MailMessage msg = new MailMessage();
                    msg.IsBodyHtml = true;
                    msg.From = new MailAddress("neps@tm.com.my", "NEPS");
                    msg.To.Add(EmailListStr);
                    msg.Subject = "Redmark Job Completed for Scheme " + schemeName;
                    msg.Body = "<h1>FILES DETAILS</h1>SCHEME NAME	: " + schemeName + "<br/><br/>JOB NO : " + jobNo + "<br/><br/>DESCRIPTION : " + jdescription + " <br/><br/>REDMARK FILE NAME: " + namaFailUpload;
                    msg.Body += "<br/><br/> <h1>AND DETAILS</h1> <br>";
                    msg.Body += "AND ID : " + User.Identity.Name + "<br/><br/>AND NAME	: " + andName + "<br/><br/>AND EMAIL	: " + andEmail + "<br/><br/>AND PHONE NUMBER: " + andTel + "<br/><br/>Job is succesfully updated.";
                    msg.IsBodyHtml = true;
                    SmtpClient emailClient = new SmtpClient("smtp.tm.com.my", 25);
                    emailClient.UseDefaultCredentials = false;
                    emailClient.Credentials = new NetworkCredential("neps", "nepsadmin", "tmmaster");
                    emailClient.EnableSsl = false;
                    emailClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                    emailClient.Send(msg);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(ex.ToString());
                }
            }
            return Json(new
            {
                Success = success
            });
        }

        public ActionResult DeleteFile(string id) //delete file
        {
            FileInfo fileID = new FileInfo(Server.MapPath("~/App_Data/uploads/job" + Session["pathId"] + "/" + id));
            fileID.Delete();
            return RedirectToAction("FileList/" + Session["pathId"]);
        }

        public ActionResult DeleteFileRNO(string id) //delete file
        {
            FileInfo fileID = new FileInfo(Server.MapPath("~/App_Data/uploads/job" + Session["pathId"] + "/" + id));
            fileID.Delete();
            return RedirectToAction("FileListRedmark/" + Session["pathId"]);
        }

        public ActionResult testLoad() //WSDL file
        {
            return View();
        }

        [HttpPost]
        public ActionResult testWSDL() // add file by folder
        {
            //using (Entities_NEPS ctxData = new Entities_NEPS())
            //{
            //    var CariOSP = (from p in ctxData.GC_NETELEM 
            //                   join fx in ctxData.GC_FDP on fx.G3E_FNO = p.G3E_FNO
            //                   where p.JOB_ID == targetJob
            //                   select p).Count();
            //}
            NEPSLoadEquipmentService osp = new NEPSLoadEquipmentService();

            //Error errorObj = null;
            string test = "NEPSLoadEquipment";
            NEPSLoadEquipmentRequestSetEquip SetEquip = new NEPSLoadEquipmentRequestSetEquip();
            SetEquip.EquipID = "EPEHWWMU15";
            SetEquip.EquipCat = "EPE";
            SetEquip.EquipVend = "HUAWEI";
            SetEquip.EquipModel = "NE40E-8";
            SetEquip.Region = "CENTRAL";
            SetEquip.State = "SELANGOR";
            SetEquip.ExchDesc = "WANGSA MAJU";
            SetEquip.Site = "WMU";
            //SetEquip.MngtIP = "10.26.1.22";
            SetEquip.TempName = "EPE HUAWEI NE40E-8 19 SLOTS CONTAINER";

            NEPSLoadEquipmentRequestSetEquipUDA SetEquipUDA = new NEPSLoadEquipmentRequestSetEquipUDA();
            SetEquipUDA.Tagging = "BAU";
            SetEquipUDA.OutdoorIndoorTagging = "INDOOR";
            SetEquipUDA.DpExtensionFlag = "NULL";
            SetEquipUDA.NetworkName = "EPEHWWMU15";

            NEPSLoadEquipmentResponseListEquipUDA ListEquipUDA = null;
            NEPSLoadEquipmentResponseCallstatus CallStatus = null;

            try
            {
                osp.NEPSLoadEquipment(test, SetEquip, SetEquipUDA, out ListEquipUDA, out CallStatus);
                System.Diagnostics.Debug.WriteLine("aaaa");
                //System.Diagnostics.Debug.WriteLine(test + "|" + SetEquip.EquipID + "|" + SetEquipUDA.Tagging + "|" + ListEquipUDA.Tagging + "|" + CallStatus.Status);
                System.Diagnostics.Debug.WriteLine(CallStatus.Status);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("BBBB");
                System.Diagnostics.Debug.WriteLine(ex.Message.ToString());
                System.Diagnostics.Debug.WriteLine("CCCC");
            }
            //NEPSLoadEquipmentResponseCallstatus success = new NEPSLoadEquipmentResponseCallstatus();
            //if (success.EquipCat == null)
            //{

            //}
            //else
            //{
            //    System.Diagnostics.Debug.WriteLine("DDDD");
            //}
            return Json(new
            {
                //Success = success
            });
        }

        public ActionResult ISP_3D() // get list file by Job No
        {
            List<SelectListItem> list1 = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.WV_EXC_MAST
                            orderby p.EXC_NAME ascending
                            select new { Text = p.EXC_NAME, Value = p.EXC_ABB };

                foreach (var a in query)
                {
                    list1.Add(new SelectListItem() { Text = a.Text, Value = a.Value.ToString() });
                }

                ViewBag.excabb = list1;
            }
            return View();
        }

        public ActionResult IOSPLink(string tjOb)
        {
            System.Diagnostics.Debug.WriteLine("get lalaaaa" + tjOb);

            //WebView.WebService._base myWebService;
            //myWebService = new WebService._base();

            //myWebService.linkIOSP(tjOb,tGTECH);

            string ID = "";

            using (Entities_NRM ctxNRM = new Entities_NRM())
            {
                var query = (from a in ctxNRM.PROJECTs
                             where a.NAME.Contains(tjOb)
                             select a).Single();

                return Json(new
                {
                    ID = query.ID.ToString(),
                    visionael = visionael
                }, JsonRequestBehavior.AllowGet);
            }
        }

        //public ActionResult LOAD_SITE(string ID) {

        //    string LOADSITE="";

        //    using (Entities ctxNeps = new Entities()) {
        //        var queryFDPLSITE = from a in ctxNeps.WV_FDP_LOAD_SITE
        //                            where ID.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
        //                             select new 
        //                             {
        //                                 a.SITENAME,
        //                                 a.SITEDESC,
        //                                 Address = a.STREETTYPE + " " + a.ADDRESSSTREET + " " + a.COUNTY + " " + a.ADDRESSCITY + " " + a.ADDRESSSTATE + " " + a.POSTCODE + " " + a.COUNTRY,
        //                                 a.EQUIPMENTLOCATION,
        //                                 a.CABLINGTYPE,
        //                                 a.FIBERTOPREMISEEXIST,
        //                                 a.ACCESSRESTRICTION,
        //                                 a.CONTACT
        //                             };

        //        if (queryFDPLSITE.Count() > 0) {

        //            foreach (var a in queryFDPLSITE)
        //            {
        //                LOADSITE = "[" + a.SITENAME + "|" + a.SITEDESC + "|" + a.Address + "|" + a.EQUIPMENTLOCATION + "|" + a.CABLINGTYPE + "|" + a.FIBERTOPREMISEEXIST + "|" + a.ACCESSRESTRICTION + "|" + a.CONTACT;
        //            }
        //        }
        //        var queryFDCLSITE = from a in ctxNeps.WV_FDC_LOAD_SITE
        //                            where ID.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
        //                            select new
        //                            {
        //                                a.SITENAME,
        //                                a.SITEDESC,
        //                                Address = a.STREETTYPE + " " + a.ADDRESSSTREET + " " + a.COUNTY + " " + a.ADDRESSCITY + " " + a.ADDRESSSTATE + " " + a.POSTCODE + " " + a.COUNTRY,
        //                                a.EQUIPMENTLOCATION,
        //                                a.CABLINGTYPE,
        //                                a.FIBERTOPREMISEEXIST,
        //                                a.ACCESSRESTRICTION,
        //                                a.CONTACT
        //                            };

        //        if (queryFDCLSITE.Count() > 0)
        //        {

        //            foreach (var a in queryFDCLSITE)
        //            {
        //                LOADSITE = "[" + a.SITENAME + "|" + a.SITEDESC + "|" + a.Address + "|" + a.EQUIPMENTLOCATION + "|" + a.CABLINGTYPE + "|" + a.FIBERTOPREMISEEXIST + "|" + a.ACCESSRESTRICTION + "|" + a.CONTACT;
        //            }
        //        }

        //        var queryDBLSITE = from a in ctxNeps.WV_DB_LOAD_SITE
        //                           where ID.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
        //                            select new
        //                            {
        //                                a.SITENAME,
        //                                a.SITEDESC,
        //                                Address = a.STREETTYPE + " " + a.ADDRESSSTREET + " " + a.COUNTY + " " + a.ADDRESSCITY + " " + a.ADDRESSSTATE + " " + a.POSTCODE + " " + a.COUNTRY,
        //                                a.EQUIPMENTLOCATION,
        //                                a.CABLINGTYPE,
        //                                a.FIBERTOPREMISEEXIST,
        //                                a.ACCESSRESTRICTION,
        //                                a.CONTACT
        //                            };

        //        if (queryDBLSITE.Count() > 0)
        //        {

        //            foreach (var a in queryDBLSITE)
        //            {
        //                LOADSITE = "[" +  a.SITENAME + "|" + a.SITEDESC + "|" + a.Address + "|" + a.EQUIPMENTLOCATION + "|" + a.CABLINGTYPE + "|" + a.FIBERTOPREMISEEXIST + "|" + a.ACCESSRESTRICTION + "|" + a.CONTACT;
        //            }
        //        }

        //        var queryFTBSITE = from a in ctxNeps.WV_FTB_LOAD_SITE
        //                           where ID.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
        //                           select new
        //                           {
        //                               a.SITENAME,
        //                               a.SITEDESC,
        //                               Address = a.STREETTYPE + " " + a.ADDRESSSTREET + " " + a.COUNTY + " " + a.ADDRESSCITY + " " + a.ADDRESSSTATE + " " + a.POSTCODE + " " + a.COUNTRY,
        //                               a.EQUIPMENTLOCATION,
        //                               a.CABLINGTYPE,
        //                               a.FIBERTOPREMISEEXIST,
        //                               a.ACCESSRESTRICTION,
        //                               a.CONTACT
        //                           };

        //        if (queryFTBSITE.Count() > 0)
        //        {

        //            foreach (var a in queryFTBSITE)
        //            {
        //                LOADSITE = "[" + a.SITENAME + "|" + a.SITEDESC + "|" + a.Address + "|" + a.EQUIPMENTLOCATION + "|" + a.CABLINGTYPE + "|" + a.FIBERTOPREMISEEXIST + "|" + a.ACCESSRESTRICTION + "|" + a.CONTACT;
        //            }
        //        }

        //        var queryTIEFDPSITE = from a in ctxNeps.WV_TIEFDP_LOAD_SITE
        //                              where ID.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
        //                           select new
        //                           {
        //                               a.SITENAME,
        //                               a.SITEDESC,
        //                               Address = a.STREETTYPE + " " + a.ADDRESSSTREET + " " + a.COUNTY + " " + a.ADDRESSCITY + " " + a.ADDRESSSTATE + " " + a.POSTCODE + " " + a.COUNTRY,
        //                               a.EQUIPMENTLOCATION,
        //                               a.CABLINGTYPE,
        //                               a.FIBERTOPREMISEEXIST,
        //                               a.ACCESSRESTRICTION,
        //                               a.CONTACT
        //                           };

        //        if (queryTIEFDPSITE.Count() > 0)
        //        {

        //            foreach (var a in queryTIEFDPSITE)
        //            {
        //                LOADSITE = "[" + a.SITENAME + "|" + a.SITEDESC + "|" + a.Address + "|" + a.EQUIPMENTLOCATION + "|" + a.CABLINGTYPE + "|" + a.FIBERTOPREMISEEXIST + "|" + a.ACCESSRESTRICTION + "|" + a.CONTACT;
        //            }
        //        }
        //        var queryVDSL2SITE = from a in ctxNeps.WV_VDSL2_LOAD_SITE
        //                             where ID.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
        //                              select new
        //                              {
        //                                  a.SITENAME,
        //                                  a.SITEDESC,
        //                                  Address = a.STREETTYPE + " " + a.ADDRESSSTREET + " " + a.COUNTY + " " + a.ADDRESSCITY + " " + a.ADDRESSSTATE + " " + a.POSTCODE + " " + a.COUNTRY,
        //                                  a.EQUIPMENTLOCATION,
        //                                  a.CABLINGTYPE,
        //                                  a.FIBERTOPREMISEEXIST,
        //                                  a.ACCESSRESTRICTION,
        //                                  a.CONTACT
        //                              };

        //        if (queryVDSL2SITE.Count() > 0)
        //        {

        //            foreach (var a in queryVDSL2SITE)
        //            {
        //                LOADSITE = "[" + a.SITENAME + "|" + a.SITEDESC + "|" + a.Address + "|" + a.EQUIPMENTLOCATION + "|" + a.CABLINGTYPE + "|" + a.FIBERTOPREMISEEXIST + "|" + a.ACCESSRESTRICTION + "|" + a.CONTACT;
        //            }
        //        }
        //        return View(LOADSITE);
        //    }
        //}

        public ActionResult CheckJob(string id)
        {
            string LOADSITE = "";
            string LOADEQUIP = "";
            string LOADBOUNDRY = "";
            string PATH = "";
            string CreateNE = "";
            string CreateCard = "";
            string CreateFC = "";
            //string PATHCOMSUMER = "";
            List<string> PATHCOMSUMER = new List<string>();
            int b = 0;
            string[] fno = new string[20];
            int count_fno = 0;
            // region Mubin - CR74-2018 - 23 Feb 2018
            Tools tool = new Tools();
            string result = "";
            string sqlStr = "";
            OracleParameter[] oraPrm;
            // endRegion
            using (OracleConnection connection = new OracleConnection(connString))
            {
                OracleCommand command = new OracleCommand("SELECT DISTINCT G3E_FNO FROM GC_NETELEM WHERE JOB_ID = '" + id + "' AND G3E_FNO IN (5600,5100,9900,5900,5800,9800,9200,9100,9500,9600,9000,13000)", connection);
                connection.Open();
                OracleDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    fno[count_fno] = reader["G3E_FNO"].ToString();
                    count_fno++;
                }
            }
            // region Mubin - CR74-2018 - 9 Mar 2018
            using (Entities9 ctxNeps = new Entities9())
            {
                System.Diagnostics.Debug.WriteLine("count_fno is " + count_fno);

                oraPrm = new OracleParameter[1];
                oraPrm[0] = new OracleParameter("schemeName", OracleDbType.Varchar2);
                oraPrm[0].Value = id;
                result = tool.ExecuteStored(connString, "CLR_GRN_NIS_STG_TBL", CommandType.StoredProcedure, oraPrm, false);
                System.Diagnostics.Debug.WriteLine(result + " in delete staging table");

                for (int i = 0; i < count_fno; i++)
                {
                    oraPrm = new OracleParameter[1];
                    oraPrm[0] = new OracleParameter("schemeName", OracleDbType.Varchar2);
                    oraPrm[0].Value = id;

                    System.Diagnostics.Debug.WriteLine("fno[i] is " + fno[i]);
                    #region 5600 FDP
                    if (fno[i].Contains("5600"))
                    {
                        sqlStr = "GRN_LOAD_FDP";
                        /*var queryFDPLSITE = from a in ctxNeps.WV_FDP_LOAD_SITE2
                                            where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                            select new
                                            {
                                                a.SITENAME,
                                                a.SITEDESC,
                                                Address = a.STREETTYPE + " " + a.ADDRESSSTREET + " " + a.COUNTY + " " + a.ADDRESSCITY + " " + a.ADDRESSSTATE + " " + a.POSTCODE + " " + a.COUNTRY,
                                                a.EQUIPMENTLOCATION,
                                                a.CABLINGTYPE,
                                                a.FIBERTOPREMISEEXIST,
                                                a.ACCESSRESTRICTION,
                                                a.CONTACT
                                            };

                        if (queryFDPLSITE.Count() > 0)
                        {

                            foreach (var a in queryFDPLSITE)
                            {
                                LOADSITE += "[" + a.SITENAME + "|" + a.SITEDESC + "|" + a.Address + "|" + a.EQUIPMENTLOCATION + "|" + a.CABLINGTYPE + "|" + a.FIBERTOPREMISEEXIST + "|" + a.ACCESSRESTRICTION + "|" + a.CONTACT;
                            }
                        }

                        var queryFDPLEQUIP = from a in ctxNeps.WV_FDP_LOAD_EQUIP2
                                             where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                             select new
                                             {
                                                 a.EQUIPID,
                                                 a.EQUIPCAT,
                                                 a.EQUIPVEND
                                             };

                        if (queryFDPLEQUIP.Count() > 0)
                        {

                            foreach (var a in queryFDPLEQUIP)
                            {
                                LOADEQUIP += "[" + a.EQUIPID + "|" + a.EQUIPCAT + "|" + a.EQUIPVEND;
                            }
                        }

                        var queryFDPSERVICEBOUNDARY = from a in ctxNeps.WV_FDP_LOAD_SERVICEBOUNDARY2
                                                      where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                                      select new
                                                      {
                                                          a.SERVICEBOUNDARYID,
                                                          Address = a.SITENO + " " + a.STREETYPE + " " + a.SECTION + " " + a.POSTCODE + " " + a.CITY + " " + a.STATE + " " + a.COUNTRY,
                                                          a.PREMISETYPE,
                                                          a.SITESDP,
                                                          a.DROPCABLEDISTANCE
                                                      };

                        if (queryFDPSERVICEBOUNDARY.Count() > 0)
                        {

                            foreach (var a in queryFDPSERVICEBOUNDARY)
                            {
                                b++;
                                LOADBOUNDRY = LOADBOUNDRY + "[" + a.SERVICEBOUNDARYID + "_00" + b + "|" + a.Address + "|" + a.PREMISETYPE + "|" + a.SITESDP + "|" + a.DROPCABLEDISTANCE + "|";
                            }
                        }*/
                    }

                    #endregion

                    #region 5100 FDC
                    if (fno[i].Contains("5100"))
                    {
                        sqlStr = "GRN_LOAD_FDC";
                        /*var queryFDCLSITE = from a in ctxNeps.WV_FDC_LOAD_SITE2
                                            where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                            select new
                                            {
                                                a.SITENAME,
                                                a.SITEDESC,
                                                Address = a.STREETTYPE + " " + a.ADDRESSSTREET + " " + a.COUNTY + " " + a.ADDRESSCITY + " " + a.ADDRESSSTATE + " " + a.POSTCODE + " " + a.COUNTRY,
                                                a.EQUIPMENTLOCATION,
                                                a.CABLINGTYPE,
                                                a.FIBERTOPREMISEEXIST,
                                                a.ACCESSRESTRICTION,
                                                a.CONTACT
                                            };

                        if (queryFDCLSITE.Count() > 0)
                        {

                            foreach (var a in queryFDCLSITE)
                            {
                                LOADSITE += "[" + a.SITENAME + "|" + a.SITEDESC + "|" + a.Address + "|" + a.EQUIPMENTLOCATION + "|" + a.CABLINGTYPE + "|" + a.FIBERTOPREMISEEXIST + "|" + a.ACCESSRESTRICTION + "|" + a.CONTACT;
                            }
                        }
                        var queryFDCLEQUIP = from a in ctxNeps.WV_FDC_LOAD_EQUIP2
                                             where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                             select new
                                             {
                                                 a.EQUIPID,
                                                 a.EQUIPCAT,
                                                 a.EQUIPVEND
                                             };

                        if (queryFDCLEQUIP.Count() > 0)
                        {

                            foreach (var a in queryFDCLEQUIP)
                            {
                                LOADEQUIP += "[" + a.EQUIPID + "|" + a.EQUIPCAT + "|" + a.EQUIPVEND;
                            }
                        }*/
                    }
                    #endregion

                    #region 9900 DB
                    if (fno[i].Contains("9900"))
                    {
                        sqlStr = "GRN_LOAD_DB";
                        /*var queryDBLSITE = from a in ctxNeps.WV_DB_LOAD_SITE2
                                           where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                           select new
                                           {
                                               a.SITENAME,
                                               a.SITEDESC,
                                               Address = a.STREETTYPE + " " + a.ADDRESSSTREET + " " + a.COUNTY + " " + a.ADDRESSCITY + " " + a.ADDRESSSTATE + " " + a.POSTCODE + " " + a.COUNTRY,
                                               a.EQUIPMENTLOCATION,
                                               a.CABLINGTYPE,
                                               a.FIBERTOPREMISEEXIST,
                                               a.ACCESSRESTRICTION,
                                               a.CONTACT
                                           };

                        if (queryDBLSITE.Count() > 0)
                        {

                            foreach (var a in queryDBLSITE)
                            {
                                LOADSITE += "[" + a.SITENAME + "|" + a.SITEDESC + "|" + a.Address + "|" + a.EQUIPMENTLOCATION + "|" + a.CABLINGTYPE + "|" + a.FIBERTOPREMISEEXIST + "|" + a.ACCESSRESTRICTION + "|" + a.CONTACT;
                            }
                        }

                        var queryDBLEQUIP = from a in ctxNeps.WV_DB_LOAD_EQUIP2
                                            where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                            select new
                                            {
                                                a.EQUIPID,
                                                a.EQUIPCAT,
                                                a.EQUIPVEND
                                            };

                        if (queryDBLEQUIP.Count() > 0)
                        {

                            foreach (var a in queryDBLEQUIP)
                            {
                                LOADEQUIP += "[" + a.EQUIPID + "|" + a.EQUIPCAT + "|" + a.EQUIPVEND;
                            }
                        }
                        var queryDBSERVICEBOUNDARY = from a in ctxNeps.WV_DB_LOAD_SERVICEBOUNDARY2
                                                     where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                                     select new
                                                     {
                                                         a.SERVICEBOUNDARYID,
                                                         Address = a.SITENO + " " + a.STREETYPE + " " + a.SECTION + " " + a.POSTCODE + " " + a.CITY + " " + a.STATE + " " + a.COUNTRY,
                                                         a.PREMISETYPE,
                                                         a.SITESDP,
                                                         a.DROPCABLEDISTANCE
                                                     };

                        if (queryDBSERVICEBOUNDARY.Count() > 0)
                        {

                            foreach (var a in queryDBSERVICEBOUNDARY)
                            {
                                b++;
                                LOADBOUNDRY = LOADBOUNDRY + "[" + a.SERVICEBOUNDARYID + "_00" + b + "|" + a.Address + "|" + a.PREMISETYPE + "|" + a.SITESDP + "|" + a.DROPCABLEDISTANCE + "|";
                            }
                        }*/
                    }
                    #endregion

                    #region 5900 FTB
                    if (fno[i].Contains("5900"))
                    {
                        sqlStr = "GRN_LOAD_FTB";
                        /*var queryFTBSITE = from a in ctxNeps.WV_FTB_LOAD_SITE2
                                           where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                           select new
                                           {
                                               a.SITENAME,
                                               a.SITEDESC,
                                               Address = a.STREETTYPE + " " + a.ADDRESSSTREET + " " + a.COUNTY + " " + a.ADDRESSCITY + " " + a.ADDRESSSTATE + " " + a.POSTCODE + " " + a.COUNTRY,
                                               a.EQUIPMENTLOCATION,
                                               a.CABLINGTYPE,
                                               a.FIBERTOPREMISEEXIST,
                                               a.ACCESSRESTRICTION,
                                               a.CONTACT
                                           };

                        if (queryFTBSITE.Count() > 0)
                        {

                            foreach (var a in queryFTBSITE)
                            {
                                LOADSITE += "[" + a.SITENAME + "|" + a.SITEDESC + "|" + a.Address + "|" + a.EQUIPMENTLOCATION + "|" + a.CABLINGTYPE + "|" + a.FIBERTOPREMISEEXIST + "|" + a.ACCESSRESTRICTION + "|" + a.CONTACT;
                            }
                        }

                        var queryFTBEQUIP = from a in ctxNeps.WV_FTB_LOAD_EQUIP2
                                            where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                            select new
                                            {
                                                a.EQUIPID,
                                                a.EQUIPCAT,
                                                a.EQUIPVEND
                                            };

                        if (queryFTBEQUIP.Count() > 0)
                        {

                            foreach (var a in queryFTBEQUIP)
                            {
                                LOADEQUIP += "[" + a.EQUIPID + "|" + a.EQUIPCAT + "|" + a.EQUIPVEND;
                            }
                        }

                        var queryFTBSERVICEBOUNDARY = from a in ctxNeps.WV_FTB_LOAD_SERVICEBOUNDARY2
                                                      where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                                      select new
                                                      {
                                                          a.SERVICEBOUNDARYID,
                                                          Address = a.SITENO + " " + a.STREETYPE + " " + a.SECTION + " " + a.POSTCODE + " " + a.CITY + " " + a.STATE + " " + a.COUNTRY,
                                                          a.PREMISETYPE,
                                                          a.SITESDP,
                                                          a.DROPCABLEDISTANCE
                                                      };

                        if (queryFTBSERVICEBOUNDARY.Count() > 0)
                        {

                            foreach (var a in queryFTBSERVICEBOUNDARY)
                            {
                                b++;
                                LOADBOUNDRY = LOADBOUNDRY + "[" + a.SERVICEBOUNDARYID + "_00" + b + "|" + a.Address + "|" + a.PREMISETYPE + "|" + a.SITESDP + "|" + a.DROPCABLEDISTANCE + "|";
                            }
                        }*/
                    }

                    #endregion

                    #region 5800 Tie FDP
                    if (fno[i].Contains("5800"))
                    {
                        sqlStr = "GRN_LOAD_TIEFDP";
                        /*var queryTIEFDPSITE = from a in ctxNeps.WV_TIEFDP_LOAD_SITE2
                                              where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                              select new
                                              {
                                                  a.SITENAME,
                                                  a.SITEDESC,
                                                  Address = a.STREETTYPE + " " + a.ADDRESSSTREET + " " + a.COUNTY + " " + a.ADDRESSCITY + " " + a.ADDRESSSTATE + " " + a.POSTCODE + " " + a.COUNTRY,
                                                  a.EQUIPMENTLOCATION,
                                                  a.CABLINGTYPE,
                                                  a.FIBERTOPREMISEEXIST,
                                                  a.ACCESSRESTRICTION,
                                                  a.CONTACT
                                              };

                        if (queryTIEFDPSITE.Count() > 0)
                        {

                            foreach (var a in queryTIEFDPSITE)
                            {
                                LOADSITE += "[" + a.SITENAME + "|" + a.SITEDESC + "|" + a.Address + "|" + a.EQUIPMENTLOCATION + "|" + a.CABLINGTYPE + "|" + a.FIBERTOPREMISEEXIST + "|" + a.ACCESSRESTRICTION + "|" + a.CONTACT;
                            }
                        }

                        var queryTIEFDPEQUIP = from a in ctxNeps.WV_TIEFDP_LOAD_EQUIP2
                                               where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                               select new
                                               {
                                                   a.EQUIPID,
                                                   a.EQUIPCAT,
                                                   a.EQUIPVEND
                                               };

                        if (queryTIEFDPEQUIP.Count() > 0)
                        {

                            foreach (var a in queryTIEFDPEQUIP)
                            {
                                LOADEQUIP += "[" + a.EQUIPID + "|" + a.EQUIPCAT + "|" + a.EQUIPVEND;
                            }
                        }

                        var queryTIEFDPESERVICEBOUNDARY = from a in ctxNeps.WV_TIEFDP_LOAD_SERVICEBOUNDARY2
                                                          where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                                          select new
                                                          {
                                                              a.SERVICEBOUNDARYID,
                                                              Address = a.SITENO + " " + a.STREETYPE + " " + a.SECTION + " " + a.POSTCODE + " " + a.CITY + " " + a.STATE + " " + a.COUNTRY,
                                                              a.PREMISETYPE,
                                                              a.SITESDP,
                                                              a.DROPCABLEDISTANCE
                                                          };

                        if (queryTIEFDPESERVICEBOUNDARY.Count() > 0)
                        {

                            foreach (var a in queryTIEFDPESERVICEBOUNDARY)
                            {
                                b++;
                                LOADBOUNDRY = LOADBOUNDRY + "[" + a.SERVICEBOUNDARYID + "_00" + b + "|" + a.Address + "|" + a.PREMISETYPE + "|" + a.SITESDP + "|" + a.DROPCABLEDISTANCE + "|";
                            }
                        }*/
                    }
                    #endregion

                    #region 9800 VDSL2
                    if (fno[i].Contains("9800"))
                    {
                        sqlStr = "GRN_LOAD_VDSL2";
                        /*var queryVDSL2SITE = from a in ctxNeps.WV_VDSL2_LOAD_SITE2
                                             where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                             select new
                                             {
                                                 a.SITENAME,
                                                 a.SITEDESC,
                                                 Address = a.STREETTYPE + " " + a.ADDRESSSTREET + " " + a.COUNTY + " " + a.ADDRESSCITY + " " + a.ADDRESSSTATE + " " + a.POSTCODE + " " + a.COUNTRY,
                                                 a.EQUIPMENTLOCATION,
                                                 a.CABLINGTYPE,
                                                 a.FIBERTOPREMISEEXIST,
                                                 a.ACCESSRESTRICTION,
                                                 a.CONTACT
                                             };

                        if (queryVDSL2SITE.Count() > 0)
                        {

                            foreach (var a in queryVDSL2SITE)
                            {
                                LOADSITE += "[" + a.SITENAME + "|" + a.SITEDESC + "|" + a.Address + "|" + a.EQUIPMENTLOCATION + "|" + a.CABLINGTYPE + "|" + a.FIBERTOPREMISEEXIST + "|" + a.ACCESSRESTRICTION + "|" + a.CONTACT;
                            }
                        }

                        var queryVDSL2EQUIP = from a in ctxNeps.WV_VDSL2_LOAD_EQUIP2
                                              where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                              select new
                                              {
                                                  a.EQUIPID,
                                                  a.EQUIPCAT,
                                                  a.EQUIPVEND
                                              };

                        if (queryVDSL2EQUIP.Count() > 0)
                        {

                            foreach (var a in queryVDSL2EQUIP)
                            {
                                LOADEQUIP += "[" + a.EQUIPID + "|" + a.EQUIPCAT + "|" + a.EQUIPVEND;
                            }
                        }

                        var queryVDSL2SERVICEBOUNDARY = from a in ctxNeps.WV_VDSL2_LOAD_SERVICEBOUNDARY2
                                                        where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                                        select new
                                                        {
                                                            a.SERVICEBOUNDARYID,
                                                            Address = a.SITENO + " " + a.STREETYPE + " " + a.SECTION + " " + a.POSTCODE + " " + a.CITY + " " + a.STATE + " " + a.COUNTRY,
                                                            a.PREMISETYPE,
                                                            a.SITESDP,
                                                            a.DROPCABLEDISTANCE
                                                        };

                        if (queryVDSL2SERVICEBOUNDARY.Count() > 0)
                        {

                            foreach (var a in queryVDSL2SERVICEBOUNDARY)
                            {
                                b++;
                                LOADBOUNDRY = LOADBOUNDRY + "[" + a.SERVICEBOUNDARYID + "_00" + b + "|" + a.Address + "|" + a.PREMISETYPE + "|" + a.SITESDP + "|" + a.DROPCABLEDISTANCE + "|";
                            }
                        }*/
                    }
                    #endregion

                    #region 9200 SDF
                    if (fno[i].Contains("9200"))
                    {
                        sqlStr = "GRN_LOAD_SDF";
                        /*var querySDFSITE = from a in ctxNeps.WV_SDF_LOAD_SITE2
                                           where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                           select new
                                           {
                                               a.SITENAME,
                                               a.SITEDESC,
                                               Address = a.STREETTYPE + " " + a.ADDRESSSTREET + " " + a.COUNTY + " " + a.ADDRESSCITY + " " + a.ADDRESSSTATE + " " + a.POSTCODE + " " + a.COUNTRY,
                                               a.EQUIPMENTLOCATION,
                                               a.CABLINGTYPE,
                                               a.FIBERTOPREMISEEXIST,
                                               a.ACCESSRESTRICTION,
                                               a.CONTACT
                                           };

                        if (querySDFSITE.Count() > 0)
                        {

                            foreach (var a in querySDFSITE)
                            {
                                LOADSITE += "[" + a.SITENAME + "|" + a.SITEDESC + "|" + a.Address + "|" + a.EQUIPMENTLOCATION + "|" + a.CABLINGTYPE + "|" + a.FIBERTOPREMISEEXIST + "|" + a.ACCESSRESTRICTION + "|" + a.CONTACT;
                            }
                        }

                        var querySDFEQUIP = from a in ctxNeps.WV_SDF_LOAD_EQUIP2
                                            where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                            select new
                                            {
                                                a.EQUIPID,
                                                a.EQUIPCAT,
                                                a.EQUIPVEND
                                            };

                        if (querySDFEQUIP.Count() > 0)
                        {

                            foreach (var a in querySDFEQUIP)
                            {
                                LOADEQUIP += "[" + a.EQUIPID + "|" + a.EQUIPCAT + "|" + a.EQUIPVEND;
                            }
                        }
                        var querySDFSERVICEBOUNDARY = from a in ctxNeps.WV_SDF_LOAD_SERVICEBND2
                                                      where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                                      select new
                                                      {
                                                          a.SERVICEBOUNDARYID,
                                                          Address = a.SITENO + " " + a.STREETYPE + " " + a.SECTION + " " + a.POSTCODE + " " + a.CITY + " " + a.STATE + " " + a.COUNTRY,
                                                          a.PREMISETYPE,
                                                          a.SITESDP,
                                                          a.DROPCABLEDISTANCE
                                                      };

                        if (querySDFSERVICEBOUNDARY.Count() > 0)
                        {

                            foreach (var a in querySDFSERVICEBOUNDARY)
                            {
                                b++;
                                LOADBOUNDRY = LOADBOUNDRY + "[" + a.SERVICEBOUNDARYID + "_00" + b + "|" + a.Address + "|" + a.PREMISETYPE + "|" + a.SITESDP + "|" + a.DROPCABLEDISTANCE + " | ";
                            }
                        }*/
                    }
                    #endregion

                    #region path from WV_LOAD_PATH_CONSUMER and WV_LOAD_PATH
                    using (Entities1 ctxNeps2 = new Entities1())
                    {
                        var queryPATH = from a in ctxNeps2.WV_LOAD_PATH2
                                        where id.Trim().ToUpper() == a.JOBID.Trim().ToUpper()
                                        select new
                                        {
                                            a.PATHNAME,
                                            a.ANAME,
                                            a.ATYPE,
                                            a.ASITE,
                                            a.ZNAME,
                                            a.ZTYPE,
                                            a.PATHTYPE
                                        };

                        if (queryPATH.Count() > 0)
                        {

                            foreach (var a in queryPATH)
                            {
                                PATH += PATH + "[" + a.PATHNAME + "|" + a.ANAME + "|" + a.ATYPE + "|" + a.ASITE + "|" + a.ZNAME + "|" + a.ZTYPE + "|" + a.PATHTYPE;
                            }
                        }
                    }
                    // endRegion
                    #endregion


                    using (Entities6 ctxData = new Entities6())
                    {
                        #region 9100 IPMSAN
                        if (fno[i].Contains("9100"))
                        {
                            // region Mubin - CR74-2018 - 23 Feb 2018
                            sqlStr = "NIS_CREATE_IPMSAN";
                            /*var queryCreateNe91 = from a in ctxData.VW_CREATENE_OSP_9100
                                                  where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper() && !id.Trim().ToUpper().Contains("COPPER TRANSFER")
                                                  select new
                                                  {
                                                      a.G3E_FID,
                                                      a.G3E_FNO,
                                                      a.LOCNTTNAME,
                                                      a.EQUPABB,
                                                      a.INDEX1,
                                                      a.EQUPMODEL,
                                                      a.MANRABB
                                                  };

                            if (queryCreateNe91.Count() > 0)
                            {

                                foreach (var a in queryCreateNe91)
                                {
                                    CreateNE += "[" + a.G3E_FID + "|" + a.LOCNTTNAME + "|" + a.EQUPABB + "|" + a.INDEX1 + "|" + a.EQUPMODEL + "|" + a.MANRABB;
                                }
                            }
                            var queryCreateNe91T = from a in ctxData.VW_CREATENE_OSP_TRANSFER91
                                                   where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper()
                                                   select new
                                                   {
                                                       a.G3E_FID,
                                                       a.G3E_FNO,
                                                       a.LOCNTTNAME,
                                                       a.EQUPABB,
                                                       a.INDEX1,
                                                       a.EQUPMODEL,
                                                       a.MANRABB
                                                   };

                            if (queryCreateNe91T.Count() > 0)
                            {

                                foreach (var a in queryCreateNe91T)
                                {
                                    CreateNE += "[" + a.G3E_FID + "|" + a.LOCNTTNAME + "|" + a.EQUPABB + "|" + a.INDEX1 + "|" + a.EQUPMODEL + "|" + a.MANRABB;
                                }
                            }
                            var queryCreateCard91 = from a in ctxData.VW_CREATECARD_OSP_9100
                                                    where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                                    select new
                                                    {
                                                        a.G3E_FID,
                                                        a.RT_CODE,
                                                        a.SLOT,
                                                        a.CARDNAME,
                                                        a.CARDMODEL,
                                                        a.TOTALCOUNT,
                                                        a.PORTSTARTNO
                                                    };

                            if (queryCreateCard91.Count() > 0)
                            {

                                foreach (var a in queryCreateCard91)
                                {
                                    CreateCard += "[" + a.G3E_FID + "|" + a.RT_CODE + "|" + a.SLOT + "|" + a.CARDNAME + "|" + a.CARDMODEL + "|" + a.TOTALCOUNT + "|" + a.PORTSTARTNO;
                                }
                            }*/
                        }
                        #endregion

                        #region 13000 DP
                        if (fno[i].Contains("13000"))
                        {
                            sqlStr = "NIS_CREATE_DP";
                            /*var queryCreateFC = from a in ctxData.VW_CREATEFC_OSP_130001
                                                where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper() && !id.Trim().ToUpper().Contains("COPPER TRANSFER")
                                                select new { a.G3E_FID, a.G3E_FNO, a.FRANNAME, a.INDEX1, a.LOCATION, a.LOCNTTNAME };
                            if (queryCreateFC.Count() > 0)
                            {

                                foreach (var a in queryCreateFC)
                                {
                                    CreateFC += "[" + a.G3E_FID + "|" + a.FRANNAME + "|" + a.INDEX1 + "|" + a.LOCATION + "|" + a.LOCNTTNAME;
                                }
                            }
                            var queryCreateFCT = from a in ctxData.VW_CREATEFC_13000_TRANS
                                                 where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper()
                                                 select new { a.G3E_FID, a.G3E_FNO, a.DISTANCE, a.INDEX1, a.LOCATION, a.LOCNTTNAME };
                            if (queryCreateFCT.Count() > 0)
                            {

                                foreach (var a in queryCreateFCT)
                                {
                                    CreateFC += "[" + a.G3E_FID + "|" + a.DISTANCE + "|" + a.INDEX1 + "|" + a.LOCATION + "|" + a.LOCNTTNAME;
                                }
                            }*/
                        }
                        #endregion

                        #region 9500 Minimux
                        if (fno[i].Contains("9500"))
                        {
                            sqlStr = "NIS_CREATE_MINIMUX";
                            /*var queryCreateNe95 = from a in ctxData.VW_CREATENE_OSP_9500
                                                  where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper() && !id.Trim().ToUpper().Contains("COPPER TRANSFER")
                                                  select new
                                                  {
                                                      a.G3E_FID,
                                                      a.G3E_FNO,
                                                      a.LOCNTTNAME,
                                                      a.EQUPABB,
                                                      a.INDEX1,
                                                      a.EQUPMODEL,
                                                      a.MANRABB
                                                  };

                            if (queryCreateNe95.Count() > 0)
                            {

                                foreach (var a in queryCreateNe95)
                                {
                                    CreateNE += "[" + a.G3E_FID + "|" + a.LOCNTTNAME + "|" + a.EQUPABB + "|" + a.INDEX1 + "|" + a.EQUPMODEL + "|" + a.MANRABB;
                                }
                            }

                            var queryCreateNe95T = from a in ctxData.VW_CREATENE_OSP_TRANSFER95
                                                   where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper()
                                                   select new
                                                   {
                                                       a.G3E_FID,
                                                       a.G3E_FNO,
                                                       a.LOCNTTNAME,
                                                       a.EQUPABB,
                                                       a.INDEX1,
                                                       a.EQUPMODEL,
                                                       a.MANRABB
                                                   };

                            if (queryCreateNe95T.Count() > 0)
                            {

                                foreach (var a in queryCreateNe95T)
                                {
                                    CreateNE += "[" + a.G3E_FID + "|" + a.LOCNTTNAME + "|" + a.EQUPABB + "|" + a.INDEX1 + "|" + a.EQUPMODEL + "|" + a.MANRABB;
                                }
                            }

                            var queryCreateCard95 = from a in ctxData.VW_CREATECARD_OSP_9500
                                                    where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                                    select new
                                                    {
                                                        a.G3E_FID,
                                                        a.MUX_CODE,
                                                        a.SLOT,
                                                        a.CARDNAME,
                                                        a.CARDMODEL,
                                                        a.TOTALCOUNT,
                                                        a.PORTSTARTNO
                                                    };

                            if (queryCreateCard95.Count() > 0)
                            {

                                foreach (var a in queryCreateCard95)
                                {
                                    CreateCard += "[" + a.G3E_FID + "|" + a.MUX_CODE + "|" + a.SLOT + "|" + a.CARDNAME + "|" + a.CARDMODEL + "|" + a.TOTALCOUNT + "|" + a.PORTSTARTNO;
                                }
                            }
                            var queryCreateFCFDP = from a in ctxData.VW_CREATEFC_OSP_9500_FDPAC
                                                   where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper()
                                                   select new { a.G3E_FID, a.G3E_FNO, a.FRANNAME, a.INDEX1, a.LOCATION, a.LOCNTTNAME };
                            if (queryCreateFCFDP.Count() > 0)
                            {

                                foreach (var a in queryCreateFCFDP)
                                {
                                    CreateFC += "[" + a.G3E_FID + "|" + a.FRANNAME + "|" + a.INDEX1 + "|" + a.LOCATION + "|" + a.LOCNTTNAME;
                                }
                            }
                            var queryCreateFCFSDF = from a in ctxData.VW_CREATEFC_OSP_9500_SDF
                                                    where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper()
                                                    select new { a.G3E_FID, a.G3E_FNO, a.FRANNAME, a.INDEX1, a.LOCATION, a.LOCNTTNAME };
                            if (queryCreateFCFSDF.Count() > 0)
                            {

                                foreach (var a in queryCreateFCFSDF)
                                {
                                    CreateFC += "[" + a.G3E_FID + "|" + a.FRANNAME + "|" + a.INDEX1 + "|" + a.LOCATION + "|" + a.LOCNTTNAME;
                                }
                            }
                            var queryCreateFCFDDF = from a in ctxData.VW_CREATEFC_OSP_9500_DDFAC
                                                    where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper()
                                                    select new { a.G3E_FID, a.G3E_FNO, a.FRANNAME, a.INDEX1, a.LOCATION, a.LOCNTTNAME };
                            if (queryCreateFCFDDF.Count() > 0)
                            {

                                foreach (var a in queryCreateFCFDDF)
                                {
                                    CreateFC += "[" + a.G3E_FID + "|" + a.FRANNAME + "|" + a.INDEX1 + "|" + a.LOCATION + "|" + a.LOCNTTNAME;
                                }
                            }*/
                        }
                        #endregion

                        #region 9600 Remote Terminal
                        if (fno[i].Contains("9600"))
                        {
                            sqlStr = "NIS_CREATE_REMOTE_TERMINAL";
                            /*var queryCreateNe96 = from a in ctxData.VW_CREATENE_OSP_9600
                                                  where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper() && !id.Trim().ToUpper().Contains("COPPER TRANSFER")
                                                  select new
                                                  {
                                                      a.G3E_FID,
                                                      a.G3E_FNO,
                                                      a.LOCNTTNAME,
                                                      a.EQUPABB,
                                                      a.INDEX1,
                                                      a.EQUPMODEL,
                                                      a.MANRABB
                                                  };

                            if (queryCreateNe96.Count() > 0)
                            {

                                foreach (var a in queryCreateNe96)
                                {
                                    CreateNE += "[" + a.G3E_FID + "|" + a.LOCNTTNAME + "|" + a.EQUPABB + "|" + a.INDEX1 + "|" + a.EQUPMODEL + "|" + a.MANRABB;
                                }
                            }

                            var queryCreateNe96T = from a in ctxData.VW_CREATENE_OSP_TRANSFER96
                                                   where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper()
                                                   select new
                                                   {
                                                       a.G3E_FID,
                                                       a.G3E_FNO,
                                                       a.LOCNTTNAME,
                                                       a.EQUPABB,
                                                       a.INDEX1,
                                                       a.EQUPMODEL,
                                                       a.MANRABB
                                                   };

                            if (queryCreateNe96T.Count() > 0)
                            {

                                foreach (var a in queryCreateNe96T)
                                {
                                    CreateNE += "[" + a.G3E_FID + "|" + a.LOCNTTNAME + "|" + a.EQUPABB + "|" + a.INDEX1 + "|" + a.EQUPMODEL + "|" + a.MANRABB;
                                }
                            }
                            var queryCreateCard96 = from a in ctxData.VW_CREATECARD_OSP_9600
                                                    where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                                    select new
                                                    {
                                                        a.G3E_FID,
                                                        a.RT_CODE,
                                                        a.SLOT,
                                                        a.CARDNAME,
                                                        a.CARDMODEL,
                                                        a.TOTALCOUNT,
                                                        a.PORTSTARTNO
                                                    };

                            if (queryCreateCard96.Count() > 0)
                            {

                                foreach (var a in queryCreateCard96)
                                {
                                    CreateCard += "[" + a.G3E_FID + "|" + a.RT_CODE + "|" + a.SLOT + "|" + a.CARDNAME + "|" + a.CARDMODEL + "|" + a.TOTALCOUNT + "|" + a.PORTSTARTNO;
                                }
                            }*/
                        }
                        #endregion

                        #region 9000 GE Aggregator
                        if (fno[i].Contains("9000"))
                        {
                            sqlStr = "NIS_CREATE_GE_AGGREGATOR";
                            /*var queryCreateNe90 = from a in ctxData.VW_CREATENE_OSP_9000
                                                  where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper() && !id.Trim().ToUpper().Contains("COPPER TRANSFER")
                                                  select new
                                                  {
                                                      a.G3E_FID,
                                                      a.G3E_FNO,
                                                      a.LOCNTTNAME,
                                                      a.EQUPABB,
                                                      a.INDEX1,
                                                      a.EQUPMODEL,
                                                      a.MANRABB
                                                  };

                            if (queryCreateNe90.Count() > 0)
                            {

                                foreach (var a in queryCreateNe90)
                                {
                                    CreateNE += "[" + a.G3E_FID + "|" + a.LOCNTTNAME + "|" + a.EQUPABB + "|" + a.INDEX1 + "|" + a.EQUPMODEL + "|" + a.MANRABB;
                                }
                            }

                            var queryCreateNe90T = from a in ctxData.VW_CREATENE_OSP_TRANSFER90
                                                   where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper()
                                                   select new
                                                   {
                                                       a.G3E_FID,
                                                       a.G3E_FNO,
                                                       a.LOCNTTNAME,
                                                       a.EQUPABB,
                                                       a.INDEX1,
                                                       a.EQUPMODEL,
                                                       a.MANRABB
                                                   };

                            if (queryCreateNe90T.Count() > 0)
                            {

                                foreach (var a in queryCreateNe90T)
                                {
                                    CreateNE += "[" + a.G3E_FID + "|" + a.LOCNTTNAME + "|" + a.EQUPABB + "|" + a.INDEX1 + "|" + a.EQUPMODEL + "|" + a.MANRABB;
                                }
                            }

                            var queryCreateCard90 = from a in ctxData.VW_CREATECARD_OSP_9000
                                                    where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                                    select new
                                                    {
                                                        a.G3E_FID,
                                                        a.RT_CODE,
                                                        a.SLOT,
                                                        a.CARDNAME,
                                                        a.CARDMODEL,
                                                        a.TOTALCOUNT,
                                                        a.PORTSTARTNO
                                                    };

                            if (queryCreateCard90.Count() > 0)
                            {

                                foreach (var a in queryCreateCard90)
                                {
                                    CreateCard += "[" + a.G3E_FID + "|" + a.RT_CODE + "|" + a.SLOT + "|" + a.CARDNAME + "|" + a.CARDMODEL + "|" + a.TOTALCOUNT + "|" + a.PORTSTARTNO;
                                }
                            }*/
                        }
                        #endregion
                    }
                    result = tool.ExecuteStored(connString, sqlStr, CommandType.StoredProcedure, oraPrm, false);
                    System.Diagnostics.Debug.WriteLine("str prd result is " + result);
                }

                #region
                var querySITE = from a in ctxNeps.GRN_LOADSITE
                                where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                orderby a.LOADSITE_ID
                                select a;

                if (querySITE.Count() > 0)
                {
                    foreach (var a in querySITE)
                    {
                        LOADSITE += "[" + a.LOADSITE_ID + "|" + a.CLASS_NAME + "|" + a.SCHEME_NAME + "|" + a.G3E_FID + "|" + a.USERID + "|" +
                                          a.SITEREGION + "|" + a.SITESTATE + "|" + a.SITEPTT + "|" + a.SITEEXC + "|" + a.SITENAME + "|" + a.SITEDESC + "|" +
                                          a.SITETYPE + "|" + a.SITEFLOOR + "|" + a.SITENO + "|" + a.SITEBUILDING + "|" + a.ADDRESSSTREET + "|" +
                                          a.ADDRESSCITY + "|" + a.ADDRESSSTATE + "|" + a.POSTCODE + "|" + a.STREETTYPE + "|" + a.COUNTY + "|" +
                                          a.COUNTRY + "|" + a.SITECOMMENT + "|" + a.EQUIPMENTLOCATION + "|" + a.CABLINGTYPE + "|" + a.COPPEROWNBYTM + "|" +
                                          a.FIBERTOPREMISEEXIST + "|" + a.TOTALSERVICEBOUNDARY + "|" + a.ACCESSRESTRICTION + "|" + a.CONTACT + "|" +
                                          a.COMMENTS + "|" + a.GRN_STATUS;
                    }
                }

                var queryEQUIP = from a in ctxNeps.GRN_LOADEQUIP
                                 where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                 orderby a.LOADEQUIP_ID
                                 select a;

                if (queryEQUIP.Count() > 0)
                {
                    foreach (var a in queryEQUIP)
                    {
                        LOADEQUIP += "[" + a.LOADEQUIP_ID + "|" + a.CLASS_NAME + "|" + a.G3E_FID + "|" + a.SCHEME_NAME + "|" + a.USERID + "|" +
                                           a.EQUIPID + "|" + a.EQUIPCAT + "|" + a.EQUIPVEND + "|" + a.EQUIMODEL + "|" + a.REGION + "|" + a.STATE + "|" +
                                           a.EXCDESC + "|" + a.SITE + "|" + a.MNGTIP + "|" + a.TEMPNAME + "|" + a.INSERVICE + "|" + a.TAGGING + "|" +
                                           a.OUTDOORINDOORTAGGING + "|" + a.DPEXTENSIONFLAG + "|" + a.NETWORKNAME + "|" + a.GRN_STATUS;
                    }
                }

                var querySERVICEBOUNDARY = from a in ctxNeps.GRN_LOADSERVBOUND
                                           where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                           orderby a.LOADSERVBOUND_ID
                                           select a;

                if (querySERVICEBOUNDARY.Count() > 0)
                {
                    foreach (var a in querySERVICEBOUNDARY)
                    {
                        b++;
                        LOADBOUNDRY = LOADBOUNDRY + "[" + a.LOADSERVBOUND_ID + "|" + a.G3E_FID + "|" + a.SCHEME_NAME + "|" + a.USERID + "|" + a.G3E_ID + "|" +
                                                          a.SERVICEBOUNDARYID + "_00" + b + "|" + a.SITENO + "|" + a.SITEFLOOR + "|" + a.SITEBUILDING + "|" +
                                                          a.STREETYPE + "|" + a.STREETNAME + "|" + a.SECTION + "|" + a.POSTCODE + "|" + a.CITY + "|" +
                                                          a.STATE + "|" + a.COUNTRY + "|" + a.PREMISETYPE + "|" + a.SERVCATEGORY + "|" + a.SITESDP + "|" +
                                                          a.SITEDPEXC + "|" + a.DDP + "|" + a.SMARTDEVELOPER + "|" + a.SMARTPROJECTID + "|" +
                                                          a.GRN_STATUS + "|" + a.COS + "|" + a.DROPCABLEDISTANCE + " | ";
                    }
                }

                var queryCreateNe = from a in ctxNeps.NIS_CREATENE
                                    where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper()
                                    orderby a.CREATENE_ID
                                    select a;

                if (queryCreateNe.Count() > 0)
                {
                    foreach (var a in queryCreateNe)
                    {
                        CreateNE += "[" + a.CREATENE_ID + "|" + a.JOB_ID + "|" + a.G3E_FID + "|" + a.G3E_FNO + "|" + a.LOCNTTNAME + "|" + a.EQUPABB + "|" +
                                          a.INDEX1 + "|" + a.EQUPMODEL + "|" + a.MANRABB;
                    }
                }

                var queryCreateCard = from a in ctxNeps.NIS_CREATECARD
                                      where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                      orderby a.CREATECARD_ID
                                      select a;

                if (queryCreateCard.Count() > 0)
                {
                    foreach (var a in queryCreateCard)
                    {
                        CreateCard += "[" + a.CREATECARD_ID + "|" + a.G3E_FID + "|" + a.MSAN_FID + "|" + a.RT_CODE + "|" + a.MUX_CODE + "|" + a.SCHEME_NAME + "|" +
                                            a.SLOT + "|" + a.CARDNAME + "|" + a.CARDMODEL + "|" + a.TOTALCOUNT + "|" + a.PORTSTARTNO;
                    }
                }
                var queryCreateFC = from a in ctxNeps.NIS_CREATEFC
                                    where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper()
                                    orderby a.CREATEFC_ID
                                    select a;

                if (queryCreateFC.Count() > 0)
                {
                    foreach (var a in queryCreateFC)
                    {
                        CreateFC += "[" + a.CREATEFC_ID + "|" + a.JOB_ID + "|" + a.G3E_FID + "|" + a.G3E_FNO + "|" + a.FRANNAME + "|" + a.INDEX1 + "|" +
                                          a.LOCNTTNAME + "|" + a.LOCATION + "|" + a.DISTANCE + "|" + a.CABINET;
                    }
                }
                #endregion // endRegion

                using (Entities7 ctxData = new Entities7())
                {

                    #region path ODP

                    /*var queryPATHCONSUMERODP = from a in ctxData.WV_LOAD_PATH_CONSUMER_ODP
                                                   where id.Trim().ToUpper() == a.JOBID.Trim().ToUpper()
                                                   orderby a.DPSITE, a.ACARD2, a.APORT2, a.ACARD3, a.APORT3
                                                   select new
                                                   {
                                                       a.ANAME,
                                                       a.ATYPE,
                                                       a.ACARD2,
                                                       a.APORT2,
                                                       a.ACARD3,
                                                       a.DPSITE,
                                                       a.ZNAME,
                                                       a.ZTYPE,
                                                       a.ZCARD,
                                                       a.ZPORT,
                                                       a.ACARD4,
                                                       a.APORT4,
                                                   };

                        if (queryPATHCONSUMERODP.Count() > 0)
                        {
                            //System.Diagnostics.Debug.WriteLine("a = " + queryPATHCONSUMER.Count());
                            foreach (var a in queryPATHCONSUMERODP)
                            {
                                //System.Diagnostics.Debug.WriteLine("a = " + a.ACARD2.Count());
                                //System.Diagnostics.Debug.WriteLine("b = " + PATHCOMSUMER);
                                PATHCOMSUMER.Add(a.ANAME + "|" + a.ATYPE + "|" + a.ACARD2 + "|" + a.APORT2 + "|" + a.ACARD3 + "|" + a.DPSITE + "|" + a.ZNAME + "|" + a.ZTYPE + "|" + a.ZCARD + "|" + a.ZPORT + "|" + a.ACARD4 + "|" + a.APORT4);
                            }
                        }*/

                    #endregion

                    // region Mubin - CR74-2018 - 23 Feb 2018
                    #region path consumer
                    var queryPATHCONSUMER = from a in ctxData.WV_LOAD_PATH_CONSUMER2
                                            where id.Trim().ToUpper() == a.JOBID.Trim().ToUpper()
                                            orderby a.DPSITE, a.ACARD2, a.APORT2, a.ACARD3, a.APORT3
                                            select new
                                            {
                                                a.ANAME,
                                                a.ATYPE,
                                                a.ACARD2,
                                                a.APORT2,
                                                a.ACARD3,
                                                a.DPSITE,
                                                a.ZNAME,
                                                a.ZTYPE,
                                                a.ZCARD,
                                                a.ZPORT,
                                                a.ACARD4,
                                                a.APORT4,
                                            };

                    if (queryPATHCONSUMER.Count() > 0)
                    {
                        //System.Diagnostics.Debug.WriteLine("a = " + queryPATHCONSUMER.Count());
                        foreach (var a in queryPATHCONSUMER)
                        {
                            //System.Diagnostics.Debug.WriteLine("a = " + a.ACARD2.Count());
                            //System.Diagnostics.Debug.WriteLine("b = " + PATHCOMSUMER);
                            PATHCOMSUMER.Add(a.ANAME + "|" + a.ATYPE + "|" + a.ACARD2 + "|" + a.APORT2 + "|" + a.ACARD3 + "|" + a.DPSITE + "|" + a.ZNAME + "|" + a.ZTYPE + "|" + a.ZCARD + "|" + a.ZPORT + "|" + a.ACARD4 + "|" + a.APORT4);
                        }
                    }
                    #endregion // endRegion
                }

            }

            return Json(new
            {
                Lsite = LOADSITE,
                Lequip = LOADEQUIP,
                Lboundry = LOADBOUNDRY,
                pConsumer = PATHCOMSUMER,
                path = PATH,
                LCreateNE = CreateNE,
                LCreateCard = CreateCard,
                LCreateFC = CreateFC
            }, JsonRequestBehavior.AllowGet);


        }

        // region Mubin - CR14-20180330
        public ActionResult CheckJobProcess(string integ, string job)
        {
            int count = 0;
            int hasRelatedFNO = 0;
            string result = "";
            Tools tool = new Tools();
            OracleParameter[] oraPrm;
            oraPrm = new OracleParameter[1];
            oraPrm[0] = new OracleParameter("schemeName", OracleDbType.Varchar2);
            oraPrm[0].Value = job;


            if (integ == "grn")
            {
                /*using (OracleConnection connection = new OracleConnection(connString))
                {
                    OracleCommand command = new OracleCommand("SELECT COUNT(*) FROM GC_NETELEM WHERE JOB_ID = '" + job + "' AND G3E_FNO IN (5600,5100,9900,5900,5800,9800,9200,8300)", connection);
                    connection.Open();
                    OracleDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                    {
                        var rawCount = reader["COUNT(*)"];
                        count = Convert.ToInt32(rawCount);
                    }
                }*/

                /*oraPrm = new OracleParameter[1];
                oraPrm[0] = new OracleParameter("schemeName", OracleDbType.Varchar2);
                oraPrm[0].Value = job;*/
                result = tool.ExecuteStored(connString2, "GRN_LOAD_STRPRD", CommandType.StoredProcedure, oraPrm, false);
                System.Diagnostics.Debug.WriteLine("GRN_LOAD_STRPRD result is " + result);
                hasRelatedFNO = 1;

                /*oraPrm = new OracleParameter[1];
                oraPrm[0] = new OracleParameter("schemeName", OracleDbType.Varchar2);
                oraPrm[0].Value = job;*/
                result = tool.ExecuteStored(connString2, "GRN_ADDCARD_STRPRD", CommandType.StoredProcedure, oraPrm, false);
                System.Diagnostics.Debug.WriteLine("GRN_ADDCARD_STRPRD result is " + result);

               
                result = tool.ExecuteStored(connString2, "GRN_DELETEEQUIP_STRPRD", CommandType.StoredProcedure, oraPrm, false);
                System.Diagnostics.Debug.WriteLine("GRN_DELETEEQUIP_STRPRD result is " + result);

                result = tool.ExecuteStored(connString2, "GRN_DELETEPATH_STRPRD", CommandType.StoredProcedure, oraPrm, false);
                System.Diagnostics.Debug.WriteLine("GRN_DELETEPATH_STRPRD result is " + result);

                result = tool.ExecuteStored(connString2, "GRN_DELETESB_STRPRD", CommandType.StoredProcedure, oraPrm, false);
                System.Diagnostics.Debug.WriteLine("GRN_DELETESB_STRPRD result is " + result);

                //region Fatihin CR53 16052018
                result = tool.ExecuteStored(connString2, "GRN_LOADPATHCON_STRPRD", CommandType.StoredProcedure, oraPrm, false);
                System.Diagnostics.Debug.WriteLine("GRN_LOADPATHCON_STRPRD result is " + result);
                //endregion

                
            }
            else if (integ == "nis")
            {
                /*using (OracleConnection connection = new OracleConnection(connString))
                {
                    OracleCommand command = new OracleCommand("SELECT COUNT(*) FROM GC_NETELEM WHERE JOB_ID = '" + job + "' AND G3E_FNO IN (9100,13000,9500,9600,9000)", connection);
                    connection.Open();
                    OracleDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                    {
                        var rawCount = reader["COUNT(*)"];
                        count = Convert.ToInt32(rawCount);
                    }
                }*/

                /*oraPrm = new OracleParameter[1];
                oraPrm[0] = new OracleParameter("schemeName", OracleDbType.Varchar2);
                oraPrm[0].Value = job;*/
                /*result = tool.ExecuteStored(connString2, "NIS_CREATE_STRPRD", CommandType.StoredProcedure, oraPrm, false);
                System.Diagnostics.Debug.WriteLine("NIS_CREATE_STRPRD result is " + result);*/
                result = tool.ExecuteStored(connString2, "NIS_CREATEFCOSP_STRPRD", CommandType.StoredProcedure, oraPrm, false);
                System.Diagnostics.Debug.WriteLine("NIS_CREATEFCOSP_STRPRD result is " + result);

                result = tool.ExecuteStored(connString2, "NIS_CREATEFUOSP_STRPRD", CommandType.StoredProcedure, oraPrm, false);
                System.Diagnostics.Debug.WriteLine("NIS_CREATEFUOSP_STRPRD result is " + result);

                result = tool.ExecuteStored(connString2, "NIS_CREATENEOSP_STRPRD", CommandType.StoredProcedure, oraPrm, false);
                System.Diagnostics.Debug.WriteLine("NIS_CREATENEOSP_STRPRD result is " + result);

                result = tool.ExecuteStored(connString2, "NIS_CREATESBOSP_STRPRD", CommandType.StoredProcedure, oraPrm, false);
                System.Diagnostics.Debug.WriteLine("NIS_CREATESBOSP_STRPRD result is " + result);

                result = tool.ExecuteStored(connString2, "NIS_CREATECARDOSP_STRPRD", CommandType.StoredProcedure, oraPrm, false);
                System.Diagnostics.Debug.WriteLine("NIS_CREATECARDOSP_STRPRD result is " + result);

                result = tool.ExecuteStored(connString2, "NIS_CREATECSOSP_STRPRD", CommandType.StoredProcedure, oraPrm, false);
                System.Diagnostics.Debug.WriteLine("NIS_CREATECSOSP_STRPRD result is " + result);

                result = tool.ExecuteStored(connString2, "NIS_CREATEISP_STRPRD", CommandType.StoredProcedure, oraPrm, false);
                System.Diagnostics.Debug.WriteLine("NIS_CREATENIS_STRPRD result is " + result);

                hasRelatedFNO = 1;
            }

            return Json(new
            {
                hasRelatedFno = hasRelatedFNO
            }, JsonRequestBehavior.AllowGet);
        }

        #region Marfu'ah ASBGated 280519
        public ActionResult checkASBGated(string typeOfInteg, string id, string ptt, string g3e_fno)
        {
            bool ASBExist = false;
            using (Entities9 ctxASB = new Entities9()) {

                int queryValidPTT = 0;
                int queryValidFeat = 0;
                int querycheckAllPTT = (from a in ctxASB.REF_GATED_ASB_PTT
                                        where a.PTT == "ALL"
                                        select a).Count();

                int querycheckAllFNO = (from a in ctxASB.REF_GATED_ASB_FEAT
                                        where a.G3E_FNO == "ALL"
                                        select a).Count();


                if (typeOfInteg.Trim() == "Load Site")
                {
                    queryValidPTT = (from a in ctxASB.REF_GATED_ASB_PTT
                                         join b in ctxASB.GRN_LOADSITE on a.PTT equals b.PTT
                                         where b.PTT.Trim() == ptt.Trim()
                                         select b).Count();

                    queryValidFeat = (from a in ctxASB.REF_GATED_ASB_FEAT
                                          join b in ctxASB.GRN_LOADSITE on a.G3E_FNO equals b.G3E_FNO
                                          where b.G3E_FNO.Trim() == g3e_fno.Trim()
                                          select b).Count();
                }

                else if (typeOfInteg.Trim() == "Load Equipment")
                {
                    queryValidPTT = (from a in ctxASB.REF_GATED_ASB_PTT
                                     join b in ctxASB.GRN_LOADEQUIP on a.PTT equals b.PTT
                                     where b.PTT.Trim() == ptt.Trim()
                                     select b).Count();

                    queryValidFeat = (from a in ctxASB.REF_GATED_ASB_FEAT
                                      join b in ctxASB.GRN_LOADEQUIP on a.G3E_FNO equals b.G3E_FNO
                                      where b.G3E_FNO.Trim() == g3e_fno.Trim()
                                      select b).Count();
                }

                if (querycheckAllPTT > 0 && querycheckAllFNO > 0)
                {
                    ASBExist = true;
                }

                if (queryValidPTT > 0 || queryValidFeat > 0 && querycheckAllFNO > 0)
                {
                   ASBExist = true;
                }

                else if (queryValidPTT > 0 || queryValidFeat > 0 && querycheckAllPTT > 0)
                    ASBExist = true;

                else
                   ASBExist = false;

                System.Diagnostics.Debug.WriteLine("ASBExist" + ASBExist);
            }
            return Json(new { id = id, isASBExist = ASBExist }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult checkASBGatedNIS(string typeOfInteg, string id, string ptt, string g3e_fno)
        {
            bool ASBExist = false;
            using (Entities9 ctxASB = new Entities9())
            {

                int queryValidPTT = 0;
                int queryValidFeat = 0;
                int querycheckAllPTT = (from a in ctxASB.REF_GATED_ASB_PTT
                                        where a.PTT == "ALL"
                                        select a).Count();

                int querycheckAllFNO = (from a in ctxASB.REF_GATED_ASB_FEAT
                                        where a.G3E_FNO == "ALL"
                                        select a).Count();


                if (typeOfInteg.Trim() == "Create Network Element")
                {
                    queryValidPTT = (from a in ctxASB.REF_GATED_ASB_PTT
                                     join b in ctxASB.NIS_CREATENE_OSP on a.PTT equals b.PTT
                                     where b.PTT.Trim() == ptt.Trim()
                                     select b).Count();

                    queryValidFeat = (from a in ctxASB.REF_GATED_ASB_FEAT
                                      join b in ctxASB.NIS_CREATENE_OSP on a.G3E_FNO equals b.G3E_FNO
                                      where b.G3E_FNO.Trim() == g3e_fno.Trim()
                                      select b).Count();
                }

                if (querycheckAllPTT > 0 && querycheckAllFNO > 0)
                {
                    ASBExist = true;
                }

                if (queryValidPTT > 0 || queryValidFeat > 0 && querycheckAllFNO > 0)
                {
                    ASBExist = true;
                }

                else if (queryValidPTT > 0 || queryValidFeat > 0 && querycheckAllPTT > 0)
                    ASBExist = true;

                else
                    ASBExist = false;

                System.Diagnostics.Debug.WriteLine("ASBExist" + ASBExist);
            }
            return Json(new { id = id, isASBExist = ASBExist }, JsonRequestBehavior.AllowGet);
        }




        #endregion

        #region Fatihin CR53 - 16 May 2018
        public ActionResult CheckJobResult(string typeOfInte, string id, int hasFnoRelated)
        {
            string LOADORCREATE = "";
            List<string> PATHCOMSUMER = new List<string>();
            int b = 0;
            if (hasFnoRelated == 1)
            {
                using (Entities9 ctxNeps = new Entities9())
                {
                    if (typeOfInte == "Load Site")
                    {
                        #region Marfu'ah 030519 CR-ASBGated
                        var querySITE = from a in ctxNeps.GRN_LOADSITE
                                        where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                        orderby a.SITENAME descending
                                        select a;

                        if (querySITE.Count() > 0)
                        {
                            foreach (var a in querySITE)
                            {
                                LOADORCREATE += "[" + a.FLAG + "|" + Convert.ToInt32(a.LOADSITE_ID) + "|" + a.G3E_FID + "|" + a.FEAT_STATE + "|" + a.SITEPTT + "|" + a.SITEEXC + "|" + a.SITENAME + "|" + a.SITEFLOOR + "|" +
                                                  a.SITENO + "|" + a.SITEBUILDING + "|" + a.ADDRESSSTREET + "|" + a.ADDRESSCITY + "|" +
                                                  a.ADDRESSSTATE + "|" + a.POSTCODE + "|" + a.STREETTYPE + "|" + a.COUNTY + "|" +
                                                  a.EQUIPMENTLOCATION + "|" + a.CABLINGTYPE + "|" + a.COPPEROWNBYTM + "|" +
                                                  a.FIBERTOPREMISEEXIST + "|" + a.TOTALSERVICEBOUNDARY + "|" + a.PTT + "|" + a.G3E_FNO;
                            }
                        }
                        #endregion
                    }

                    else if (typeOfInte == "Load Equipment")
                    {
                        #region Marfu'ah 030519 CR-ASBGated
                        var queryEQUIP = from a in ctxNeps.GRN_LOADEQUIP
                                         where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                         orderby a.SITE descending
                                         select a;

                        if (queryEQUIP.Count() > 0)
                        {
                            foreach (var a in queryEQUIP)
                            {
                                LOADORCREATE += "[" + a.FLAG + "|" + Convert.ToInt32(a.LOADEQUIP_ID) + "|" + a.G3E_FID + "|" + a.FEAT_STATE + "|" + a.EQUIPCAT + "|" + a.EQUIPVEND + "|" + a.EQUIMODEL + "|" + a.EXCDESC + "|" +
                                                  a.SITE + "|" + a.TEMPNAME + "|" + a.PATCHCORD + "|" + a.PROJECTNAME + "|" + a.DESIGN + "|" + a.PTT + "|" + a.G3E_FNO;
                            }
                        }
                        #endregion
                    }

                    else if (typeOfInte == "Load Service Boundary Address")
                    {
                        #region Marfu'ah 030519 CR-ASBGated
                        var querySERVICEBOUNDARY = from a in ctxNeps.GRN_LOADSERVBOUND
                                                   where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                                   orderby a.SERVICEBOUNDARYID descending
                                                   select a;

                        if (querySERVICEBOUNDARY.Count() > 0)
                        {
                            foreach (var a in querySERVICEBOUNDARY)
                            {
                                b++;
                                LOADORCREATE += "[" + a.FLAG + "|" + Convert.ToInt32(a.LOADSERVBOUND_ID) + "|" + a.G3E_ID + "|" + a.SERVICEBOUNDARYID + "_00" + b + "|" + a.SITENO + "|" +
                                                            a.SITEFLOOR + "|" + a.SITEBUILDING + "|" + a.STREETYPE + "|" + a.STREETNAME + "|" +
                                                            a.SECTION + "|" + a.POSTCODE + "|" + a.CITY + "|" + a.STATE + "|" + a.PREMISETYPE + "|" +
                                                            a.SITESDP + " | " + a.DDP;
                            }
                        }
                        #endregion
                    }

                    else if (typeOfInte == "Add Card/Splitter")
                    {
                        #region Marfu'ah 030519 CR-ASBGated
                        var queryADDCARD = from a in ctxNeps.GRN_ADDCARD
                                           where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                           orderby a.EQUIPMENTID descending
                                           select a;

                        if (queryADDCARD.Count() > 0)
                        {
                            foreach (var a in queryADDCARD)
                            {
                                LOADORCREATE += "[" + a.FLAG + "|" + Convert.ToInt32(a.ADDCARD_ID) + "|" + a.EQUIPMENTID + "|" + a.EQUIPVEND + "|" + a.EQUIPMODEL + "|" + a.TEMPLATENAME + "|" + a.SLOT_NO + "|" +
                                                  a.CARD_TYPE + "|" + a.TOTAL_PORT + "|" + a.PORTIN;
                            }
                        }
                        #endregion
                    }

                    else if (typeOfInte == "Delete Equipment")
                    {
                        #region Marfu'ah 030519 CR-ASBGated
                        var queryDELETEEQUIP = from a in ctxNeps.GRN_DELETEEQUIP
                                               where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper()
                                               orderby a.GRANITE_EQUIP_ID descending
                                               select a;

                        if (queryDELETEEQUIP.Count() > 0)
                        {
                            foreach (var a in queryDELETEEQUIP)
                            {
                                LOADORCREATE += "[" + a.FLAG + "|" + Convert.ToInt32(a.DELETEEQUIP_ID) + "|" + a.G3E_FID + "|" + a.G3E_FNO + "|" + a.JOB_ID + "|" + a.GRANITE_EQUIP_ID + "|" + a.EQUIPMENTID + "|" +
                                                  a.EQUIPCAT + "|" + a.SLOT + "|" + a.PORTIN;
                            }
                        }
                        #endregion
                    }

                    else if (typeOfInte == "Delete Path")
                    {
                        #region Marfu'ah 030519 CR-ASBGated
                        var queryDELETEPATH = from a in ctxNeps.GRN_DELETEPATH
                                              where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper()
                                              orderby a.JOB_ID descending
                                              select a;

                        if (queryDELETEPATH.Count() > 0)
                        {
                            foreach (var a in queryDELETEPATH)
                            {
                                LOADORCREATE += "[" + a.FLAG + "|" + Convert.ToInt32(a.DELETEPATH_ID) + "|" + a.G3E_FID + "|" + a.G3E_FNO + "|" + a.JOB_ID + "|" + a.PATHNAME + "|" + a.PATHTYPE;
                            }
                        }
                        #endregion
                    }

                    else if (typeOfInte == "Delete Service Boundary")
                    {
                        #region Marfu'ah 030519 CR-ASBGated
                        var queryDELETESB = from a in ctxNeps.GRN_DELETESB
                                            where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper()
                                            orderby a.JOB_ID descending
                                            select a;

                        if (queryDELETESB.Count() > 0)
                        {
                            foreach (var a in queryDELETESB)
                            {
                                LOADORCREATE += "[" + a.FLAG + "|" + Convert.ToInt32(a.DELETESB_ID) + "|" + a.G3E_FID + "|" + a.G3E_FNO + "|" + a.JOB_ID + "|" + a.SERVICEBOUNDARYID + "|" + a.SITEDP;
                            }
                        }
                        #endregion
                    }

                    else if (typeOfInte.Trim() == "Create Network Element")
                    {
                        //01042019 FATIHIN
                        #region EF

                        var queryCreateNe = from a in ctxNeps.NIS_CREATENE_OSP
                                            where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper()
                                            orderby a.NE_ID
                                            select a;

                        if (queryCreateNe.Count() > 0)
                        {
                            foreach (var a in queryCreateNe)
                            {
                                LOADORCREATE += "[" + a.FLAG + "|" + Convert.ToInt32(a.NE_ID) + "|" + a.FEATURE_STATE + "|" + a.LOCATIONDETAIL + "|" + a.AREACODE + "|" + a.AREADESCRIPTION + "|" + a.AREAAREACODE + "|" +
                                                  a.AREAARETCODE + "|" + a.EQUPLOCNTTNAME + "|" + a.EQUPEQUTABBREVIATION + "|" + a.EQUPINDEX + "|" + a.EQUPSTATUS + "|" + a.EQUPMANRABBREVIATION + "|" + a.EQUPEQUMMODEL
                                                   + "|" + a.PTT + "|" + a.G3E_FNO; 
                            }
                        }

                        var queryCreateNeISP = from a in ctxNeps.NIS_ISP_CREATENE
                                               where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                               orderby a.CREATENE_ID
                                               select a;

                        if (queryCreateNeISP.Count() > 0)
                        {
                            foreach (var a in queryCreateNeISP)
                            {
                                LOADORCREATE += "[" + a.FLAG + "|" + Convert.ToInt32(a.CREATENE_ID) + "|" + a.LOCNDETAIL + "|"+ a.AREACODE+ "|" + a.AREADESCRIPTION + "|" + a.AREAAREACODE + "|" +
                                              a.AREAARETCODE + "|" + a.EQUPLOCNTTNAME + "|" + a.EQUPEQUTABBREAVIATION + "|" + a.EQUPINDEX + "|" + a.EQUPSTATUS + "|" + a.EQUPMANRABBREAVIATION + "|"+ a.EQUPEQUMMODEL;

                                //LOADORCREATE += "[" + a.CREATENE_ID + "|" + a.SCHEME_NAME + "|" + a.G3E_FID + "|ISP|" + a.EQUPLOCNTTNAME + "|" + a.EQUPEQUTABBREAVIATION + "|" +
                                //                  a.EQUPINDEX + "|" + a.EQUPEQUMMODEL + "|" + a.EQUPMANRABBREAVIATION;
                            }
                        }
                        #endregion
                        #region
                        /*
                          var queryCreateNe = from a in ctxNeps.NIS_CREATENE
                                              where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper()
                                              orderby a.CREATENE_ID
                                              select a;

                          if (queryCreateNe.Count() > 0)
                          {
                              foreach (var a in queryCreateNe)
                              {
                                  LOADORCREATE += "[" + a.CREATENE_ID + "|" + a.JOB_ID + "|" + a.G3E_FID + "|" + a.G3E_FNO + "|" + a.LOCNTTNAME + "|" + a.EQUPABB + "|" +
                                                    a.INDEX1 + "|" + a.EQUPMODEL + "|" + a.MANRABB;
                              }
                          }

                          var queryCreateNeISP = from a in ctxNeps.NIS_ISP_CREATENE
                                                 where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                                 orderby a.CREATENE_ID
                                                 select a;

                          if (queryCreateNeISP.Count() > 0)
                          {
                              foreach (var a in queryCreateNeISP)
                              {
                                  LOADORCREATE += "[" + a.CREATENE_ID + "|" + a.SCHEME_NAME + "|" + a.G3E_FID + "|ISP|" + a.EQUPLOCNTTNAME + "|" + a.EQUPEQUTABBREAVIATION + "|" +
                                                    a.EQUPINDEX + "|" + a.EQUPEQUMMODEL + "|" + a.EQUPMANRABBREAVIATION;
                              }
                          }*/
                        #endregion
                    }

                    else if (typeOfInte.Trim() == "Create Card")
                    {
                        #region EF
                        var queryCreateCard = from a in ctxNeps.NIS_CREATECARD_OSP
                                              where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper()
                                              orderby a.CARD_ID
                                              select new { 
                                              a.CARD_ID,
                                              a.FLAG,
                                              a.CARDSLOT,
                                              a.CARDNAME,
                                              a.CARDMODEL,
                                              a.CARDSTATUS,
                                              a.CARDSERIAL,
                                              a.CARDCOUNTPORT,
                                              a.PORTSTARTNUM }
                                              ;

                        if (queryCreateCard.Count() > 0)
                        {
                            foreach (var a in queryCreateCard)
                            {
                                LOADORCREATE += "{" + a.FLAG + "|" + Convert.ToInt32(a.CARD_ID) + "|" + a.CARDSLOT + "|" + a.CARDNAME + "|" + a.CARDMODEL + "|" +
                                                    a.CARDSTATUS + "|" + a.CARDSERIAL + "|" + a.CARDCOUNTPORT + "|" + a.PORTSTARTNUM;

                            }
                        }

                        var queryCreateCardisp = from a in ctxNeps.NIS_ISP_CREATECARD
                                                 where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                                 orderby a.CREATECARD_ID
                                                 select new {
                                                 a.CREATECARD_ID,
                                                 a.FLAG,
                                                 a.SLOT,
                                                 a.CARDNAME,
                                                 a.CARDMODEL,
                                                 a.CARDSTATUS,
                                                 a.CARDCOUNTPORT,
                                                 a.PORTSTARTNUM};


                        if (queryCreateCardisp.Count() > 0)
                        {
                            foreach (var a in queryCreateCardisp)
                            {
                                LOADORCREATE += "{" + a.FLAG + "|" + Convert.ToInt32(a.CREATECARD_ID) + "|" + a.SLOT + "|" + a.CARDNAME + "|" +a.CARDMODEL + "|" + 
                                    a.CARDSTATUS + "|" + "-" + "|" + a.CARDCOUNTPORT + "|" + a.PORTSTARTNUM;
                                //LOADORCREATE += "[" + a.CREATECARD_ID + "|" + a.G3E_FID + "|" + a.ITEM_ID + "|||" + a.SCHEME_NAME + "|" +
                                //                    a.SLOT + "|" + a.CARDNAME + "|" + a.CARDMODEL + "|" + a.CARDCOUNTPORT + "|" + a.PORTSTARTNUM;
                            }
                            

                        }
                        #endregion
                        #region
                        /*
                          var queryCreateCard = from a in ctxNeps.NIS_CREATECARD
                                                where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                                orderby a.CREATECARD_ID
                                                select a;

                          if (queryCreateCard.Count() > 0)
                          {
                              foreach (var a in queryCreateCard)
                              {
                                  LOADORCREATE += "[" + a.CREATECARD_ID + "|" + a.G3E_FID + "|" + a.MSAN_FID + "|" + a.RT_CODE + "|" + a.MUX_CODE + "|" + a.SCHEME_NAME + "|" +
                                                      a.SLOT + "|" + a.CARDNAME + "|" + a.CARDMODEL + "|" + a.TOTALCOUNT + "|" + a.PORTSTARTNO;
                              }
                          }

                          var queryCreateCardisp = from a in ctxNeps.NIS_ISP_CREATECARD
                                                where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                                orderby a.CREATECARD_ID
                                                select a;

                          if (queryCreateCardisp.Count() > 0)
                          {
                              foreach (var a in queryCreateCardisp)
                              {
                                  LOADORCREATE += "[" + a.CREATECARD_ID + "|" + a.G3E_FID + "|" + a.ITEM_ID + "|||" + a.SCHEME_NAME + "|" +
                                                      a.SLOT + "|" + a.CARDNAME + "|" + a.CARDMODEL + "|" + a.CARDCOUNTPORT + "|" + a.PORTSTARTNUM;
                              }
                          }*/
                        #endregion
                    }

                    else if (typeOfInte.Trim() == "Create Frame Container")
                    {
                        #region EF
                        var queryCreateFC = from a in ctxNeps.NIS_CREATEFC_OSP
                                            where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper()
                                            orderby a.FC_ID
                                            select new
                                            {
                                                a.FLAG,
                                                a.FC_ID,
                                                a.LOCNTTNAME,
                                                a.FRANNAME,
                                                a.INDEXX,
                                                a.STATUS,
                                                a.LOCATIONDETAIL,
                                                a.AREACODE,
                                                a.AREADESCRIPTION,
                                                a.AREAAREACODE,
                                                a.AREAARETCODE
                                            };

                        if (queryCreateFC.Count() > 0)
                        {
                            foreach (var a in queryCreateFC)
                            {
                                LOADORCREATE += "[" + a.FLAG + "|" + Convert.ToInt32(a.FC_ID) + "|" + a.LOCNTTNAME + "|" + a.FRANNAME + "|" + a.INDEXX + "|" +
                                                  a.STATUS + "|" + a.LOCATIONDETAIL + "|" + a.AREACODE + "|" + a.AREADESCRIPTION + "|" + a.AREAAREACODE + "|" + a.AREAARETCODE;
                            }
                        }

                        var queryCreateFCISP = from a in ctxNeps.NIS_ISP_CREATEFC
                                               where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                               orderby a.CREATEFC_ID
                                               select a;

                        if (queryCreateFCISP.Count() > 0)
                        {
                            foreach (var a in queryCreateFCISP)
                            {
                                LOADORCREATE += "[" + a.FLAG + "|" + Convert.ToInt32(a.CREATEFC_ID) + "|" + a.LOCNTTNAME + "|" + a.FRANNAME + "|" + a.INDEXX + "|" +
                                                  a.STATUS + "|" + a.LOCATIONDETAILS + "|" + a.AREACODE + "|" + a.AREADESCRIPTION + "|" + a.AREAAREACODE + "|" + a.AREAARETCODE;
                            }
                        }
                        #endregion
                        #region
                        /*
                          var queryCreateFC = from a in ctxNeps.NIS_CREATEFC
                                              where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper()
                                              orderby a.CREATEFC_ID
                                              select a;

                          if (queryCreateFC.Count() > 0)
                          {
                              foreach (var a in queryCreateFC)
                              {
                                  LOADORCREATE += "[" + a.CREATEFC_ID + "|" + a.JOB_ID + "|" + a.G3E_FID + "|" + a.G3E_FNO + "|" + a.FRANNAME + "|" + a.INDEX1 + "|" +
                                                    a.LOCNTTNAME + "|" + a.LOCATION + "||" + a.DISTANCE + "|" + a.CABINET;
                              }
                          }

                          var queryCreateFCISP = from a in ctxNeps.NIS_ISP_CREATEFC
                                                 where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                                 orderby a.CREATEFC_ID
                                                 select a;

                          if (queryCreateFCISP.Count() > 0)
                          {
                              foreach (var a in queryCreateFCISP)
                              {
                                  LOADORCREATE += "[" + a.CREATEFC_ID + "|" + a.SCHEME_NAME + "|" + a.VFDITEM_ID + "|5500|" + a.FRANNAME + "|" + a.INDEXX + "|" +
                                                    a.LOCNTTNAME + "|" + a.LOCATIONDETAILS + "|" + a.STATUS + "||";
                              }
                          }*/
                        #endregion
                    }
                    else if (typeOfInte.Trim() == "Create Frame Unit")
                    {
                        #region EF
                        var queryCreateFU = from a in ctxNeps.NIS_CREATEFU_OSP
                                            where a.JOB_ID == id.Trim()
                                            orderby a.FU_ID
                                            select new {
                                                a.FLAG,
                                                a.FU_ID,
                                                a.FRACID,
                                                a.FRAUPOSITION,
                                                a.FRAUNAME,
                                                a.FRAUDISTANCE,
                                                a.FUPTMANRABBREAVIATION,
                                                a.TERMINATIONTYPE,
                                                a.PRODUCTTYPE,
                                                a.COUNTPAIR
                                            };

                        if (queryCreateFU.Count() > 0)
                        {
                            //System.Diagnostics.Debug.WriteLine(queryCreateFU.Count());
                            foreach (var a in queryCreateFU)
                            {
                                //System.Diagnostics.Debug.WriteLine("Fu SUCCESS");
                                LOADORCREATE += "[" + a.FLAG + "|" + Convert.ToInt32(a.FU_ID) + "|" + a.FRACID + "|" + a.FRAUPOSITION + "|" + a.FRAUNAME + "|" + a.FRAUDISTANCE + "|" +
                                                  a.FUPTMANRABBREAVIATION + "|" + a.TERMINATIONTYPE + "|" + a.PRODUCTTYPE + "|" + a.COUNTPAIR;

                            }
                        }

                       var queryCreateFUISP = from a in ctxNeps.NIS_ISP_CREATEFU
                                               where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                               orderby a.CREATEFU_ID
                                               select a;

                        if (queryCreateFUISP.Count() > 0)
                        {
                            foreach (var a in queryCreateFUISP)
                            {
                                LOADORCREATE += "[" + a.FLAG + "|" + Convert.ToInt32(a.CREATEFU_ID) +  "|" + a.FRACID + "|" + a.FRAUPOSITION + "|" + a.FRAUNAME + "|" +
                                                  a.FRAUDISTANCE + "|" + a.FUPTMANRABBREAVIATION + "|" + a.TERMINATIONTYPE + "|" + a.PRODUCTTYPE + "|" + a.COUNTPAIR;
                            }
                        }
                        #endregion
                        #region
                        /*
                          var queryCreateFUISP = from a in ctxNeps.NIS_ISP_CREATEFU
                                                 where id.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                                 orderby a.CREATEFU_ID
                                                 select a;

                          if (queryCreateFUISP.Count() > 0)
                          {
                              foreach (var a in queryCreateFUISP)
                              {
                                  LOADORCREATE += "[" + a.CREATEFU_ID + "|" + a.SCHEME_NAME + "|" + a.G3E_FID + "|" + a.FRACID + "|" + a.FRAUPOSITION + "|" + a.FRAUNAME + "|" +
                                                    a.FRAUDISTANCE + "|" + a.FUPTMANRABBREAVIATION + "|" + a.TERMINATIONTYPE + "|" + a.PRODUCTTYPE + "|" + a.COUNTPAIR;
                              }
                          }*/
                        #endregion
                    }
                    else if (typeOfInte.Trim() == "Create Service Boundary")
                    {
                        #region EF
                        var queryCreateSB = from a in ctxNeps.NIS_CREATESB_OSP
                                            where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper()
                                            orderby a.SB_ID
                                            select a;

                        if (queryCreateSB.Count() > 0)
                        {
                            foreach (var a in queryCreateSB)
                            {
                                LOADORCREATE += "[" + a.FLAG + "|" + Convert.ToInt32(a.SB_ID) + "|" + a.STREETNUMBER + "|" + a.STREETTYPE + "|" +
                                                  a.STREETNAME + "|" + a.SUBURB + "|" + a.CITY + "|" + a.STATABBREVIATION + "|" + a.COUNTRY;
                            }
                        }


                        #endregion

                    }
                    else if (typeOfInte.Trim() == "Create Cable Sheath")
                    {
                        #region EF
                        var queryCreateCS = from a in ctxNeps.NIS_CREATECS_OSP
                                            where id.Trim().ToUpper() == a.JOB_ID.Trim().ToUpper()
                                            orderby a.CS_ID
                                            select a;

                        if (queryCreateCS.Count() > 0)
                        {
                            foreach (var a in queryCreateCS)
                            {
                                LOADORCREATE += "[" + a.FLAG + "|" + Convert.ToInt32(a.CS_ID) + "|" + a.CABSNAME + "|" + a.CABSCASTTYPE + "|" + a.CABSSTATUS + "|" +
                                                  a.CSPTPRODUCTNAME + "|" + a.MINNUMBER + "|" + a.MAXNUMBER + "|" + a.PAIRCOUNT + "|" + a.LOCATIONA + "|" + a.DISTANCE;


                            }
                        }


                        #endregion

                    }

                    else if (typeOfInte.Trim() == "Load Path Consumer")
                    {
                        #region path ODP
                        /*var queryPATHCONSUMERODP = from a in ctxData.WV_LOAD_PATH_CONSUMER_ODP
                                                         where id.Trim().ToUpper() == a.JOBID.Trim().ToUpper()
                                                         orderby a.DPSITE, a.ACARD2, a.APORT2, a.ACARD3, a.APORT3
                                                         select new
                                                         {
                                                             a.ANAME,
                                                             a.ATYPE,
                                                             a.ACARD2,
                                                             a.APORT2,
                                                             a.ACARD3,
                                                             a.DPSITE,
                                                             a.ZNAME,
                                                             a.ZTYPE,
                                                             a.ZCARD,
                                                             a.ZPORT,
                                                             a.ACARD4,
                                                             a.APORT4,
                                                         };

                              if (queryPATHCONSUMERODP.Count() > 0)
                              {
                                  //System.Diagnostics.Debug.WriteLine("a = " + queryPATHCONSUMER.Count());
                                  foreach (var a in queryPATHCONSUMERODP)
                                  {
                                      //System.Diagnostics.Debug.WriteLine("a = " + a.ACARD2.Count());
                                      //System.Diagnostics.Debug.WriteLine("b = " + PATHCOMSUMER);
                                      PATHCOMSUMER.Add(a.ANAME + "|" + a.ATYPE + "|" + a.ACARD2 + "|" + a.APORT2 + "|" + a.ACARD3 + "|" + a.DPSITE + "|" + a.ZNAME + "|" + a.ZTYPE + "|" + a.ZCARD + "|" + a.ZPORT + "|" + a.ACARD4 + "|" + a.APORT4);
                                  }
                              }*/
                        #endregion
                        //26012019 FATIHIN
                        #region path consumer
                        var queryPATHCONSUMER = from a in ctxNeps.GRN_LOADPATHCON
                                                where id.Trim().ToUpper() == a.JOBID.Trim().ToUpper()
                                                orderby a.DPSITE, a.ACARD2, a.APORT2, a.ACARD3, a.APORT3
                                                select new
                                                {
                                                    a.ANAME,
                                                    a.FEAT_STATE,
                                                    a.FLAG_FS,
                                                    a.ATYPE,
                                                    a.ACARD2,
                                                    a.APORT2,
                                                    a.ACARD3,
                                                    a.DPSITE,
                                                    a.ZNAME,
                                                    a.ZTYPE,
                                                    a.ZCARD,
                                                    a.ZPORT,
                                                    a.ACARD4,
                                                    a.APORT4,
                                                    a.DP_TAMBAHAN,
                                                    a.ID,
                                                    a.FLAG,
                                                    a.DPPORT
                                                };

                        if (queryPATHCONSUMER.Count() > 0)
                        {
                            //System.Diagnostics.Debug.WriteLine("a = " + queryPATHCONSUMER.Count());
                            foreach (var a in queryPATHCONSUMER)
                            {
                                /*string add_dp = "";
                                if (a.DP_TAMBAHAN == "UPDATE")
                                {
                                    add_dp = "YES";
                                }
                                else if (a.DP_TAMBAHAN == "NEW")
                                {
                                    add_dp = "NO";
                                }*/
                                //System.Diagnostics.Debug.WriteLine("a = " + a.ACARD2.Count());
                                //System.Diagnostics.Debug.WriteLine("b = " + PATHCOMSUMER);
                                PATHCOMSUMER.Add(a.ID + "|" + a.FLAG + "|" + a.ANAME + "|" + a.FEAT_STATE + "|" + a.ATYPE + "|" + a.ACARD2 + "|" + a.APORT2 + "|" + a.ACARD3 + "|" + a.DPSITE + "|" + a.ZNAME + "|" + a.ZTYPE + "|" + a.ZCARD + "|" + a.ZPORT + "|" + a.ACARD4 + "|" + a.APORT4 + "|" + a.DPPORT + "|" + a.DP_TAMBAHAN);
                            }
                        }
                        #endregion
                    }
                }
            }

            return Json(new
            {
                LoadOrCreate = LOADORCREATE,
                pConsumer = PATHCOMSUMER
            }, JsonRequestBehavior.AllowGet);
        }
        #endregion

        #region Mubin CR59/CR60 29082018
        public ActionResult GraniteQnd(string reportType, string ptt, string exc, string equipcode)
        {
            string result = "";
            using (Entities9 ctxNeps = new Entities9())
            {
                if (reportType == "Granite QND FDC")
                {
                    #region
                    if (ptt == "" && exc == "")
                    {
                        var queryFdc = from a in ctxNeps.BI_GRNQND_FDC
                                       orderby a.ID descending
                                       select a;
                        if (queryFdc.Count() > 0)
                        {
                            foreach (var a in queryFdc)
                            {
                                result += "[" + a.ID + "|" + a.BATCH_ID + "|" + a.NODE_INST + "|" + a.NODE_NAME + "|" + a.PTT_ID + "|" + a.EXC_ABB + "|" + a.FDC_CODE + "|" +
                                                  a.NODE_TYPE + "|" + a.NODE_VENDOR + "|" + a.NODE_MODEL + "|" + a.NODE_STATUS + "|" + a.SHELF_INST + "|" + a.SHELF_NAME + "|" +
                                                  a.SHELF_STATUS + "|" + a.SLOT_NUMBER + "|" + a.CARD_INSTID + "|" + a.CARD_TYPE + "|" + a.CARD_STATUS + "|" + a.PORT_INST + "|" +
                                                  a.PORT_NUMBER + "|" + a.PORT_BANDWIDTH + "|" + a.PORT_STATUS + "|" + a.PATH_INST + "|" + a.PATH_NAME + "|" + a.PATH_BANDWIDTH + "|" +
                                                  a.PATH_STATUS + "|" + a.DATE_CREATED;
                            }
                        }
                    }
                    else if (exc != "")
                    {
                        var queryFdc = from a in ctxNeps.BI_GRNQND_FDC
                                       where a.EXC_ABB == exc && a.FDC_CODE == equipcode
                                       orderby a.ID descending
                                       select a;
                        if (queryFdc.Count() > 0)
                        {
                            foreach (var a in queryFdc)
                            {
                                result += "[" + a.ID + "|" + a.BATCH_ID + "|" + a.NODE_INST + "|" + a.NODE_NAME + "|" + a.PTT_ID + "|" + a.EXC_ABB + "|" + a.FDC_CODE + "|" +
                                                  a.NODE_TYPE + "|" + a.NODE_VENDOR + "|" + a.NODE_MODEL + "|" + a.NODE_STATUS + "|" + a.SHELF_INST + "|" + a.SHELF_NAME + "|" +
                                                  a.SHELF_STATUS + "|" + a.SLOT_NUMBER + "|" + a.CARD_INSTID + "|" + a.CARD_TYPE + "|" + a.CARD_STATUS + "|" + a.PORT_INST + "|" +
                                                  a.PORT_NUMBER + "|" + a.PORT_BANDWIDTH + "|" + a.PORT_STATUS + "|" + a.PATH_INST + "|" + a.PATH_NAME + "|" + a.PATH_BANDWIDTH + "|" +
                                                  a.PATH_STATUS + "|" + a.DATE_CREATED;
                            }
                        }
                    }
                    else if (ptt != "" && exc == "")
                    {
                        var queryFdc = from a in ctxNeps.BI_GRNQND_FDC
                                       where ptt == a.PTT_ID
                                       orderby a.ID descending
                                       select a;
                        if (queryFdc.Count() > 0)
                        {
                            foreach (var a in queryFdc)
                            {
                                result += "[" + a.ID + "|" + a.BATCH_ID + "|" + a.NODE_INST + "|" + a.NODE_NAME + "|" + a.PTT_ID + "|" + a.EXC_ABB + "|" + a.FDC_CODE + "|" +
                                                  a.NODE_TYPE + "|" + a.NODE_VENDOR + "|" + a.NODE_MODEL + "|" + a.NODE_STATUS + "|" + a.SHELF_INST + "|" + a.SHELF_NAME + "|" +
                                                  a.SHELF_STATUS + "|" + a.SLOT_NUMBER + "|" + a.CARD_INSTID + "|" + a.CARD_TYPE + "|" + a.CARD_STATUS + "|" + a.PORT_INST + "|" +
                                                  a.PORT_NUMBER + "|" + a.PORT_BANDWIDTH + "|" + a.PORT_STATUS + "|" + a.PATH_INST + "|" + a.PATH_NAME + "|" + a.PATH_BANDWIDTH + "|" +
                                                  a.PATH_STATUS + "|" + a.DATE_CREATED;
                            }
                        }
                    }
                    #endregion
                }

                else if (reportType == "Granite QND MSAN")
                {
                    #region
                    if (ptt == "" && exc == "")
                    {
                        var queryMsan = from a in ctxNeps.BI_GRNQND_MSAN
                                        orderby a.ID descending
                                        select a;
                        if (queryMsan.Count() > 0)
                        {
                            foreach (var a in queryMsan)
                            {
                                result += "[" + a.ID + "|" + a.BATCH_ID + "|" + a.NODE_INST + "|" + a.NODE_NAME + "|" + a.PTT_ID + "|" + a.EXC_ABB + "|" + a.MSAN_CODE + "|" +
                                                  a.NODE_TYPE + "|" + a.NODE_VENDOR + "|" + a.NODE_MODEL + "|" + a.NODE_STATUS + "|" + a.SHELF_INST + "|" + a.SHELF_NAME + "|" +
                                                  a.SHELF_STATUS + "|" + a.SLOT_NUMBER + "|" + a.CARD_INSTID + "|" + a.CARD_TYPE + "|" + a.CARD_STATUS + "|" + a.PORT_INST + "|" +
                                                  a.PORT_NUMBER + "|" + a.PORT_BANDWIDTH + "|" + a.PORT_STATUS + "|" + a.PATH_INST + "|" + a.PATH_NAME + "|" + a.PATH_BANDWIDTH + "|" +
                                                  a.PATH_STATUS + "|" + a.DATE_CREATED;
                            }
                        }
                    }
                    else if (exc != "")
                    {
                        var queryMsan = from a in ctxNeps.BI_GRNQND_MSAN
                                        where exc == a.EXC_ABB && a.MSAN_CODE == equipcode
                                        orderby a.ID descending
                                        select a;
                        if (queryMsan.Count() > 0)
                        {
                            foreach (var a in queryMsan)
                            {
                                result += "[" + a.ID + "|" + a.BATCH_ID + "|" + a.NODE_INST + "|" + a.NODE_NAME + "|" + a.PTT_ID + "|" + a.EXC_ABB + "|" + a.MSAN_CODE + "|" +
                                                  a.NODE_TYPE + "|" + a.NODE_VENDOR + "|" + a.NODE_MODEL + "|" + a.NODE_STATUS + "|" + a.SHELF_INST + "|" + a.SHELF_NAME + "|" +
                                                  a.SHELF_STATUS + "|" + a.SLOT_NUMBER + "|" + a.CARD_INSTID + "|" + a.CARD_TYPE + "|" + a.CARD_STATUS + "|" + a.PORT_INST + "|" +
                                                  a.PORT_NUMBER + "|" + a.PORT_BANDWIDTH + "|" + a.PORT_STATUS + "|" + a.PATH_INST + "|" + a.PATH_NAME + "|" + a.PATH_BANDWIDTH + "|" +
                                                  a.PATH_STATUS + "|" + a.DATE_CREATED;
                            }
                        }
                    }
                    else if (ptt != "" && exc == "")
                    {
                        var queryMsan = from a in ctxNeps.BI_GRNQND_MSAN
                                        where ptt == a.PTT_ID
                                        orderby a.ID descending
                                        select a;
                        if (queryMsan.Count() > 0)
                        {
                            foreach (var a in queryMsan)
                            {
                                result += "[" + a.ID + "|" + a.BATCH_ID + "|" + a.NODE_INST + "|" + a.NODE_NAME + "|" + a.PTT_ID + "|" + a.EXC_ABB + "|" + a.MSAN_CODE + "|" +
                                                  a.NODE_TYPE + "|" + a.NODE_VENDOR + "|" + a.NODE_MODEL + "|" + a.NODE_STATUS + "|" + a.SHELF_INST + "|" + a.SHELF_NAME + "|" +
                                                  a.SHELF_STATUS + "|" + a.SLOT_NUMBER + "|" + a.CARD_INSTID + "|" + a.CARD_TYPE + "|" + a.CARD_STATUS + "|" + a.PORT_INST + "|" +
                                                  a.PORT_NUMBER + "|" + a.PORT_BANDWIDTH + "|" + a.PORT_STATUS + "|" + a.PATH_INST + "|" + a.PATH_NAME + "|" + a.PATH_BANDWIDTH + "|" +
                                                  a.PATH_STATUS + "|" + a.DATE_CREATED;
                            }
                        }
                    }
                    #endregion
                }

                else if (reportType == "Granite QND ODP")
                {
                    #region
                    if (ptt == "" && exc == "")
                    {
                        var queryOdp = from a in ctxNeps.BI_GRNQND_ODP
                                       orderby a.ID descending
                                       select a;
                        if (queryOdp.Count() > 0)
                        {
                            foreach (var a in queryOdp)
                            {
                                result += "[" + a.ID + "|" + a.BATCH_ID + "|" + a.NODE_INST + "|" + a.NODE_NAME + "|" + a.ODP_CODE + "|" + a.PTT_ID + "|" + a.EXC_ABB + "|" +
                                                  a.FDC_CODE + "|" + a.DP_CODE + "|" + a.NODE_TYPE + "|" + a.NODE_VENDOR + "|" + a.NODE_MODEL + "|" + a.NODE_STATUS + "|" +
                                                  a.SLOT_NUMBER + "|" + a.CARD_INSTID + "|" + a.CARD_TYPE + "|" + a.CARD_STATUS + "|" + a.PORT_INST + "|" +
                                                  a.PORT_NUMBER + "|" + a.PORT_BANDWIDTH + "|" + a.PORT_STATUS + "|" + a.PATH_INST + "|" + a.PATH_NAME + "|" + a.PATH_BANDWIDTH + "|" +
                                                  a.PATH_STATUS + "|" + a.DATE_CREATED;
                            }
                        }
                    }
                    else if (exc != "")
                    {
                        var queryOdp = from a in ctxNeps.BI_GRNQND_ODP
                                       where exc == a.EXC_ABB && a.ODP_CODE == equipcode
                                       orderby a.ID descending
                                       select a;
                        if (queryOdp.Count() > 0)
                        {
                            foreach (var a in queryOdp)
                            {
                                result += "[" + a.ID + "|" + a.BATCH_ID + "|" + a.NODE_INST + "|" + a.NODE_NAME + "|" + a.ODP_CODE + "|" + a.PTT_ID + "|" + a.EXC_ABB + "|" +
                                                  a.FDC_CODE + "|" + a.DP_CODE + "|" + a.NODE_TYPE + "|" + a.NODE_VENDOR + "|" + a.NODE_MODEL + "|" + a.NODE_STATUS + "|" +
                                                  a.SLOT_NUMBER + "|" + a.CARD_INSTID + "|" + a.CARD_TYPE + "|" + a.CARD_STATUS + "|" + a.PORT_INST + "|" +
                                                  a.PORT_NUMBER + "|" + a.PORT_BANDWIDTH + "|" + a.PORT_STATUS + "|" + a.PATH_INST + "|" + a.PATH_NAME + "|" + a.PATH_BANDWIDTH + "|" +
                                                  a.PATH_STATUS + "|" + a.DATE_CREATED;
                            }
                        }
                    }
                    else if (ptt != "" && exc == "")
                    {
                        var queryOdp = from a in ctxNeps.BI_GRNQND_ODP
                                       where ptt == a.PTT_ID
                                       orderby a.ID descending
                                       select a;
                        if (queryOdp.Count() > 0)
                        {
                            foreach (var a in queryOdp)
                            {
                                result += "[" + a.ID + "|" + a.BATCH_ID + "|" + a.NODE_INST + "|" + a.NODE_NAME + "|" + a.ODP_CODE + "|" + a.PTT_ID + "|" + a.EXC_ABB + "|" +
                                                  a.FDC_CODE + "|" + a.DP_CODE + "|" + a.NODE_TYPE + "|" + a.NODE_VENDOR + "|" + a.NODE_MODEL + "|" + a.NODE_STATUS + "|" +
                                                  a.SLOT_NUMBER + "|" + a.CARD_INSTID + "|" + a.CARD_TYPE + "|" + a.CARD_STATUS + "|" + a.PORT_INST + "|" +
                                                  a.PORT_NUMBER + "|" + a.PORT_BANDWIDTH + "|" + a.PORT_STATUS + "|" + a.PATH_INST + "|" + a.PATH_NAME + "|" + a.PATH_BANDWIDTH + "|" +
                                                  a.PATH_STATUS + "|" + a.DATE_CREATED;
                            }
                        }
                    }
                    #endregion
                }

                else if (reportType == "Granite QND OLT")
                {
                    #region
                    if (ptt == "" && exc == "")
                    {
                        var queryOlt = from a in ctxNeps.BI_GRNQND_OLT
                                       orderby a.ID descending
                                       select a;
                        if (queryOlt.Count() > 0)
                        {
                            foreach (var a in queryOlt)
                            {
                                result += "[" + a.ID + "|" + a.BATCH_ID + "|" + a.NODE_INST + "|" + a.NODE_NAME + "|" + a.PTT_ID + "|" + a.EXC_ABB + "|" + a.OLT_CODE + "|" +
                                                  a.NODE_TYPE + "|" + a.NODE_VENDOR + "|" + a.NODE_MODEL + "|" + a.NODE_STATUS + "|" + a.SHELF_INST + "|" + a.SHELF_NAME + "|" +
                                                  a.SHELF_STATUS + "|" + a.SLOT_NUMBER + "|" + a.CARD_INSTID + "|" + a.CARD_TYPE + "|" + a.CARD_STATUS + "|" + a.PORT_INST + "|" +
                                                  a.PORT_NUMBER + "|" + a.PORT_BANDWIDTH + "|" + a.PORT_STATUS + "|" + a.PATH_INST + "|" + a.PATH_NAME + "|" + a.PATH_BANDWIDTH + "|" +
                                                  a.PATH_STATUS + "|" + a.DATE_CREATED;
                            }
                        }
                    }
                    else if (exc != "")
                    {
                        var queryOlt = from a in ctxNeps.BI_GRNQND_OLT
                                       where exc == a.EXC_ABB && a.OLT_CODE == equipcode
                                       orderby a.ID descending
                                       select a;
                        if (queryOlt.Count() > 0)
                        {
                            foreach (var a in queryOlt)
                            {
                                result += "[" + a.ID + "|" + a.BATCH_ID + "|" + a.NODE_INST + "|" + a.NODE_NAME + "|" + a.PTT_ID + "|" + a.EXC_ABB + "|" + a.OLT_CODE + "|" +
                                                  a.NODE_TYPE + "|" + a.NODE_VENDOR + "|" + a.NODE_MODEL + "|" + a.NODE_STATUS + "|" + a.SHELF_INST + "|" + a.SHELF_NAME + "|" +
                                                  a.SHELF_STATUS + "|" + a.SLOT_NUMBER + "|" + a.CARD_INSTID + "|" + a.CARD_TYPE + "|" + a.CARD_STATUS + "|" + a.PORT_INST + "|" +
                                                  a.PORT_NUMBER + "|" + a.PORT_BANDWIDTH + "|" + a.PORT_STATUS + "|" + a.PATH_INST + "|" + a.PATH_NAME + "|" + a.PATH_BANDWIDTH + "|" +
                                                  a.PATH_STATUS + "|" + a.DATE_CREATED;
                            }
                        }
                    }
                    else if (ptt != "" && exc == "")
                    {
                        var queryOlt = from a in ctxNeps.BI_GRNQND_OLT
                                       where ptt == a.PTT_ID
                                       orderby a.ID descending
                                       select a;
                        if (queryOlt.Count() > 0)
                        {
                            foreach (var a in queryOlt)
                            {
                                result += "[" + a.ID + "|" + a.BATCH_ID + "|" + a.NODE_INST + "|" + a.NODE_NAME + "|" + a.PTT_ID + "|" + a.EXC_ABB + "|" + a.OLT_CODE + "|" +
                                                  a.NODE_TYPE + "|" + a.NODE_VENDOR + "|" + a.NODE_MODEL + "|" + a.NODE_STATUS + "|" + a.SHELF_INST + "|" + a.SHELF_NAME + "|" +
                                                  a.SHELF_STATUS + "|" + a.SLOT_NUMBER + "|" + a.CARD_INSTID + "|" + a.CARD_TYPE + "|" + a.CARD_STATUS + "|" + a.PORT_INST + "|" +
                                                  a.PORT_NUMBER + "|" + a.PORT_BANDWIDTH + "|" + a.PORT_STATUS + "|" + a.PATH_INST + "|" + a.PATH_NAME + "|" + a.PATH_BANDWIDTH + "|" +
                                                  a.PATH_STATUS + "|" + a.DATE_CREATED;
                            }
                        }
                    }
                    #endregion
                }
            }

            return Json(new
            {
                repRes = result
            }, JsonRequestBehavior.AllowGet);
        }



        #endregion

        #region Fatihin CR53 - 16 May 2018
        public ActionResult UpdateFlagInteg(string flagIntegY, string flagIntegN, string updateY, string newN, string typeOfInte)
        {
            System.Diagnostics.Debug.WriteLine("UPDATE!!");
            Tools tool = new Tools();
            bool success = true;
            bool success2 = true;
            bool success3 = true;
            bool success4 = true;
            string sqlCmd = "";
            string sqlCmd2 = "";
            string sqlCmd3 = "";
            string sqlCmd4 = "";
            string IntegId = "";
            string IntegTbl = "";

            string[] flagY = flagIntegY.Split(',');
            string[] flagN = flagIntegN.Split(',');
            string[] dpY = updateY.Split(',');
            string[] dpN = newN.Split(',');

            using (Entities9 ctxNeps = new Entities9())
            {
                if (typeOfInte == "Load Site")
                {
                    IntegTbl = "GRN_LOADSITE";
                    IntegId = "LOADSITE_ID";
                }

                else if (typeOfInte == "Load Equipment")
                {
                    IntegTbl = "GRN_LOADEQUIP";
                    IntegId = "LOADEQUIP_ID";
                }

                else if (typeOfInte == "Load Service Boundary Address")
                {
                    IntegTbl = "GRN_LOADSERVBOUND";
                    IntegId = "LOADSERVBOUND_ID";
                }

                else if (typeOfInte == "Add Card/Splitter")
                {
                    IntegTbl = "GRN_ADDCARD";
                    IntegId = "ADDCARD_ID";
                }

                else if (typeOfInte == "Delete Equipment")
                {
                    IntegTbl = "GRN_DELETEEQUIP";
                    IntegId = "DELETEEQUIP_ID";
                }

                else if (typeOfInte == "Delete Path")
                {
                    IntegTbl = "GRN_DELETEPATH";
                    IntegId = "DELETEPATH_ID";
                }

                else if (typeOfInte == "Delete Service Boundary")
                {
                    IntegTbl = "GRN_DELETESB";
                    IntegId = "DELETESB_ID";
                }

                else if (typeOfInte == "Load Path Consumer")
                {
                    IntegTbl = "GRN_LOADPATHCON";
                    IntegId = "ID";
                }

                if (flagY[0] != "")
                {
                    sqlCmd += " UPDATE " + IntegTbl + " SET FLAG = 'Y' WHERE " + IntegId + " IN (";
                    foreach (var x in flagY)
                    {
                        sqlCmd += x;
                        var lastItem = flagY.LastOrDefault();
                        if (!x.Equals(lastItem))
                        {
                            sqlCmd += ",";
                        }
                    }
                    sqlCmd += ")";
                    System.Diagnostics.Debug.WriteLine("SQL Statement :" + sqlCmd);
                    //System.Diagnostics.Debug.WriteLine("Flag[0]:" + flagY[0]);
                    success = tool.ExecuteSql(ctxNeps, sqlCmd);
                    System.Diagnostics.Debug.WriteLine("Sql Success :" + success);
                }

                //UPDATE FLAG TO N
                if (flagN[0] != "")
                {
                    sqlCmd2 += " UPDATE " + IntegTbl + " SET FLAG = 'N' WHERE " + IntegId + " IN (";
                    foreach (var x in flagN)
                    {
                        sqlCmd2 += x;
                        var lastItem = flagN.LastOrDefault();
                        if (!x.Equals(lastItem))
                        {
                            sqlCmd2 += ",";
                        }
                    }
                    sqlCmd2 += ")";
                    System.Diagnostics.Debug.WriteLine("SQL Statement :" + sqlCmd2);
                    //System.Diagnostics.Debug.WriteLine("Flag[0]:" + flagN[0]);
                    success2 = tool.ExecuteSql(ctxNeps, sqlCmd2);
                }
                //UPDATE dptambahan TO update
                if (dpY[0] != "" && typeOfInte == "Load Path Consumer")
                {
                    sqlCmd3 += " UPDATE GRN_LOADPATHCON SET DP_TAMBAHAN = 'UPDATE' WHERE ID IN (";
                    foreach (var d in dpY)
                    {
                        sqlCmd3 += d;
                        var lastItem = dpY.LastOrDefault();
                        if (!d.Equals(lastItem))
                        {
                            sqlCmd3 += ",";
                        }
                    }
                    sqlCmd3 += ")";
                    System.Diagnostics.Debug.WriteLine("SQL Statement :" + sqlCmd3);
                    success3 = tool.ExecuteSql(ctxNeps, sqlCmd3);
                }
                //UPDATE dptambahan TO update
                if (dpN[0] != "" && typeOfInte == "Load Path Consumer")
                {
                    sqlCmd4 += " UPDATE GRN_LOADPATHCON SET DP_TAMBAHAN = 'NEW' WHERE  ID IN (";
                    foreach (var d in dpN)
                    {
                        sqlCmd4 += d;
                        var lastItem = dpN.LastOrDefault();
                        if (!d.Equals(lastItem))
                        {
                            sqlCmd4 += ",";
                        }
                    }
                    sqlCmd4 += ")";
                    System.Diagnostics.Debug.WriteLine("SQL Statement :" + sqlCmd4);
                    success2 = tool.ExecuteSql(ctxNeps, sqlCmd4);
                }

            }
            return Json(new
            {
                Success = success,
                Success2 = success2,
                Success3 = success3,
                Success4 = success4
            }, JsonRequestBehavior.AllowGet);
        }
        #endregion

        #region Fatihin EF NIS
        public ActionResult UpdateFlagIntegNIS(string flagIntegY, string flagIntegN, string typeOfInte, string OSPISP)
        {
            System.Diagnostics.Debug.WriteLine("UPDATE!!");
            Tools tool = new Tools();
            bool success = true;
            bool success2 = true;

            string sqlCmd = "";
            string sqlCmd2 = "";

            string IntegId = "";
            string IntegTbl = "";

            string[] flagY = flagIntegY.Split(',');
            string[] flagN = flagIntegN.Split(',');


            using (Entities9 ctxNeps = new Entities9())
            {
                if (typeOfInte.Trim() == "Create Frame Container")
                {
                    if (OSPISP == "OSP")
                    {
                        IntegTbl = "NIS_CREATEFC_OSP";
                        IntegId = "FC_ID";
                    }
                    else
                    {
                        IntegTbl = "NIS_ISP_CREATEFC";
                        IntegId = "CREATEFC_ID";
                    }
                   
                }

                else if (typeOfInte.Trim() == "Create Frame Unit")
                {
                    if (OSPISP == "OSP")
                    {
                        IntegTbl = "NIS_CREATEFU_OSP";
                        IntegId = "FU_ID";
                    }
                    else
                    {
                        IntegTbl = "NIS_ISP_CREATEFU";
                        IntegId = "CREATEFU_ID";
                    }
                    
                }

                else if (typeOfInte.Trim() == "Create Service Boundary")
                {
                    IntegTbl = "NIS_CREATESB_OSP";
                    IntegId = "SB_ID";
                }

                else if (typeOfInte.Trim() == "Create Network Element")
                {
                    if (OSPISP == "OSP")
                    {
                        IntegTbl = "NIS_CREATENE_OSP";
                        IntegId = "NE_ID";
                    }
                    else
                    {
                        IntegTbl = "NIS_ISP_CREATENE";
                        IntegId = "CREATENE_ID";
                    }
                    
                }

                else if (typeOfInte.Trim() == "Create Card")
                {
                    if (OSPISP == "OSP")
                    {
                        IntegTbl = "NIS_CREATECARD_OSP";
                        IntegId = "CARD_ID";
                    }
                    else
                    {
                        IntegTbl = "NIS_ISP_CREATECARD";
                        IntegId = "CREATECARD_ID";
                    }
                   
                }

                else if (typeOfInte.Trim() == "Create Cable Sheath")
                {
                    IntegTbl = "NIS_CREATECS_OSP";
                    IntegId = "CS_ID";
                }



                if (flagY[0] != "")
                {
                    sqlCmd += " UPDATE " + IntegTbl + " SET FLAG = 'Y' WHERE " + IntegId + " IN (";
                    foreach (var x in flagY)
                    {
                        sqlCmd += x;
                        var lastItem = flagY.LastOrDefault();
                        if (!x.Equals(lastItem))
                        {
                            sqlCmd += ",";
                        }
                    }
                    sqlCmd += ")";
                    System.Diagnostics.Debug.WriteLine("SQL Statement :" + sqlCmd);
                    //System.Diagnostics.Debug.WriteLine("Flag[0]:" + flagY[0]);
                    success = tool.ExecuteSql(ctxNeps, sqlCmd);
                    System.Diagnostics.Debug.WriteLine("Sql Success :" + success);
                }

                //UPDATE FLAG TO N
                if (flagN[0] != "")
                {
                    sqlCmd2 += " UPDATE " + IntegTbl + " SET FLAG = 'N' WHERE " + IntegId + " IN (";
                    foreach (var x in flagN)
                    {
                        sqlCmd2 += x;
                        var lastItem = flagN.LastOrDefault();
                        if (!x.Equals(lastItem))
                        {
                            sqlCmd2 += ",";
                        }
                    }
                    sqlCmd2 += ")";
                    System.Diagnostics.Debug.WriteLine("SQL Statement :" + sqlCmd2);
                    //System.Diagnostics.Debug.WriteLine("Flag[0]:" + flagN[0]);
                    success2 = tool.ExecuteSql(ctxNeps, sqlCmd2);
                }


            }
            return Json(new
            {
                Success = success,
                Success2 = success2
            }, JsonRequestBehavior.AllowGet);
        }
        #endregion

        public class TypeOfIntegration
        {
            public string type;
            public string status;

            public TypeOfIntegration()
            {
                this.status = "";
            }
        }

        //Fatihin 16052018
        public ActionResult CheckJobGRN(string job, int hasRelatedFno, string JOBTYPE)
        {
            TypeOfIntegration[] integType = new TypeOfIntegration[8];
            integType[0] = new TypeOfIntegration();
            integType[0].type = "Load Site";
            integType[1] = new TypeOfIntegration();
            integType[1].type = "Load Equipment";
            integType[2] = new TypeOfIntegration();
            integType[2].type = "Load Service Boundary Address";
            integType[3] = new TypeOfIntegration();
            integType[3].type = "Load Path Consumer";
            integType[4] = new TypeOfIntegration();
            integType[4].type = "Add Card/Splitter";
            integType[5] = new TypeOfIntegration();
            integType[5].type = "Delete Equipment";
            integType[6] = new TypeOfIntegration();
            integType[6].type = "Delete Path";
            integType[7] = new TypeOfIntegration();
            integType[7].type = "Delete Service Boundary";

            ViewBag.typeOfInteg = integType;
            ViewBag.schemeName = job;
            ViewBag.hasRelatedFno = hasRelatedFno;
            ViewBag.JOBTYPE = JOBTYPE;
            return View();
        }

        //fatihin 19042018
        public ActionResult CheckJobNIS(string job, int hasRelatedFno, string type)
        {
            TypeOfIntegration[] integType = new TypeOfIntegration[6];
            integType[0] = new TypeOfIntegration();
            integType[0].type = "Create Frame Container";
            integType[1] = new TypeOfIntegration();
            integType[1].type = "Create Frame Unit";
            integType[2] = new TypeOfIntegration();
            integType[2].type = "Create Service Boundary";
            integType[3] = new TypeOfIntegration();
            integType[3].type = "Create Network Element";
            integType[4] = new TypeOfIntegration();
            integType[4].type = "Create Card";
            integType[5] = new TypeOfIntegration();
            integType[5].type = "Create Cable Sheath";

            ViewBag.typeOfInteg = integType;
            ViewBag.schemeName = job;
            ViewBag.hasRelatedFno = hasRelatedFno;
            ViewBag.type = type;
            //ViewBag.JobType = JobType;
            return View();
        }
        // endRegion

        #region Mubin CR45-17052018
        public ActionResult GenerateXLSX(string targetJob)
        {
            using (Entities9 ctxNeps = new Entities9())
            {
                var querySITE = from a in ctxNeps.GRN_LOADSITE
                                join b in ctxNeps.GRN_LOADEQUIP on a.SITENAME equals b.SITE
                                where targetJob.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                orderby a.SITENAME descending
                                select new
                                {
                                    a.SITEEXC,
                                    a.SITETYPE,
                                    a.SITENAME,
                                    a.SITENO,
                                    a.SITEFLOOR,
                                    a.SITEBUILDING,
                                    a.STREETTYPE,
                                    a.ADDRESSSTREET,
                                    a.POSTCODE,
                                    a.ADDRESSCITY,
                                    a.ADDRESSSTATE,
                                    a.COUNTY,
                                    a.COUNTRY,
                                    a.ACCESSRESTRICTION,
                                    a.COMMENTS,
                                    a.CONTACT,
                                    a.TOTALSERVICEBOUNDARY,
                                    b.INSERVICE,
                                    b.PATCHCORD
                                };

                var siteDetail = (from a in ctxNeps.GRN_LOADSITE
                                  where targetJob.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                  orderby a.SITENAME ascending
                                  select new
                                  {
                                      a.SITEREGION,
                                      a.SITESTATE,
                                      a.SITEPTT,
                                      a.ADDRESSCITY
                                  }).FirstOrDefault();

                var querySB = from d in ctxNeps.GRN_LOADSERVBOUND
                              join e in ctxNeps.GRN_LOADSITE on d.SITESDP equals e.SITENAME
                              where targetJob.Trim().ToUpper() == d.SCHEME_NAME.Trim().ToUpper()
                              orderby d.SITESDP descending
                              select new
                              {
                                  d.SITENO,
                                  d.SITEFLOOR,
                                  d.SITEBUILDING,
                                  d.STREETYPE,
                                  d.STREETNAME,
                                  d.SECTION,
                                  d.POSTCODE,
                                  d.CITY,
                                  d.STATE,
                                  d.COUNTRY,
                                  d.PREMISETYPE,
                                  d.SITESDP,
                                  d.SMARTDEVELOPER,
                                  d.SMARTPROJECTID,
                                  e.COPPEROWNBYTM,
                                  e.FIBERTOPREMISEEXIST
                              };

                var queryOLT = from f in ctxNeps.WV_LOAD_PATHCONSUMER
                               where targetJob.Trim().ToUpper() == f.JOBID.Trim().ToUpper()
                               orderby f.DPSITE descending
                               select new
                               {
                                   f.ZNAME,
                                   f.ZCARD,
                                   f.ZPORT
                               };
                #region LOADSITE
                var queryLOADSITE = from a in ctxNeps.GRN_LOADSITE
                                    where targetJob.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                    orderby a.SITENAME descending
                                    select new
                                    {
                                        a.LOADSITE_ID,
                                        a.CLASS_NAME,
                                        a.SCHEME_NAME,
                                        a.G3E_FID,
                                        a.USERID,
                                        a.SITEREGION,
                                        a.SITESTATE,
                                        a.SITEPTT,
                                        a.SITEEXC,
                                        a.SITENAME,
                                        a.SITEDESC,
                                        a.SITETYPE,
                                        a.SITENO,
                                        a.SITEFLOOR,
                                        a.SITEBUILDING,
                                        a.ADDRESSSTREET,
                                        a.ADDRESSCITY,
                                        a.ADDRESSSTATE,
                                        a.POSTCODE,
                                        a.STREETTYPE,
                                        a.COUNTY,
                                        a.COUNTRY,
                                        a.SITECOMMENT,
                                        a.EQUIPMENTLOCATION,
                                        a.CABLINGTYPE,
                                        a.COPPEROWNBYTM,
                                        a.FIBERTOPREMISEEXIST,
                                        a.TOTALSERVICEBOUNDARY,
                                        a.ACCESSRESTRICTION,
                                        a.CONTACT,
                                        a.COMMENTS,
                                        a.GRN_STATUS,
                                        a.FLAG
                                    };
                #endregion

                #region LOADEQUIP
                var queryLOADEQUIP = from a in ctxNeps.GRN_LOADEQUIP
                                     where targetJob.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                     select new
                                     {
                                         a.LOADEQUIP_ID,
                                         a.CLASS_NAME,
                                         a.G3E_FID,
                                         a.SCHEME_NAME,
                                         a.USERID,
                                         a.EQUIPID,
                                         a.EQUIPCAT,
                                         a.EQUIPVEND,
                                         a.EQUIMODEL,
                                         a.REGION,
                                         a.STATE,
                                         a.EXCDESC,
                                         a.SITE,
                                         a.MNGTIP,
                                         a.TEMPNAME,
                                         a.INSERVICE,
                                         a.TAGGING,
                                         a.OUTDOORINDOORTAGGING,
                                         a.DPEXTENSIONFLAG,
                                         a.NETWORKNAME,
                                         a.GRN_STATUS,
                                         a.PATCHCORD,
                                         a.FLAG
                                     }
                                     ;
                #endregion

                #region LOADSERVBOUND
                var queryLOADSERVBOUND = from a in ctxNeps.GRN_LOADSERVBOUND
                                         where targetJob.Trim().ToUpper() == a.SCHEME_NAME.Trim().ToUpper()
                                         select new
                                         {
                                             a.LOADSERVBOUND_ID,
                                             a.G3E_FID,
                                             a.SCHEME_NAME,
                                             a.USERID,
                                             a.G3E_ID,
                                             a.SERVICEBOUNDARYID,
                                             a.SITENO,
                                             a.SITEFLOOR,
                                             a.SITEBUILDING,
                                             a.STREETYPE,
                                             a.STREETNAME,
                                             a.SECTION,
                                             a.POSTCODE,
                                             a.CITY,
                                             a.STATE,
                                             a.COUNTRY,
                                             a.PREMISETYPE,
                                             a.SERVCATEGORY,
                                             a.SITESDP,
                                             a.SITEDPEXC,
                                             a.DDP,
                                             a.SMARTDEVELOPER,
                                             a.SMARTPROJECTID,
                                             a.GRN_STATUS,
                                             a.COS,
                                             a.DROPCABLEDISTANCE,
                                             a.FLAG
                                         };
                #endregion

                #region LOADPATH
                var queryLOADPATH = from f in ctxNeps.WV_LOAD_PATHCONSUMER
                                    where targetJob.Trim().ToUpper() == f.JOBID.Trim().ToUpper()
                                    orderby f.DPSITE descending
                                    select new
                                    {
                                        f.ID,
                                        f.ANAME,
                                        f.ATYPE,
                                        f.ASITE,
                                        f.ACARD2,
                                        f.APORT2,
                                        f.ACARD3,
                                        f.APORT3,
                                        f.ZNAME,
                                        f.ZTYPE,
                                        f.ZSITE,
                                        f.ZCARD,
                                        f.ZPORT,
                                        f.DPNAME,
                                        f.DPSITE,
                                        f.DPPORT,
                                        f.JOBID,
                                        f.PROCESSID,
                                        f.PROCESSTIME,
                                        f.OUTPORTFDC,
                                        f.INPORTFDP,
                                        f.REMARKS,
                                        f.APORT4,
                                        f.ACARD4,
                                        f.DP_TAMBAHAN,
                                        f.FLAG

                                    }
                                    ;
                #endregion

                #region ADDCARD
                var queryADDCARD = from f in ctxNeps.GRN_ADDCARD
                                   where targetJob.Trim().ToUpper() == f.SCHEME_NAME.Trim().ToUpper()
                                   select new
                                   {
                                       f.ADDCARD_ID,
                                       f.CLASS_NAME,
                                       f.G3E_FID,
                                       f.SCHEME_NAME,
                                       f.USERID,
                                       f.EQUIPMENTID,
                                       f.EQUIPVEND,
                                       f.EQUIPMODEL,
                                       f.TEMPLATENAME,
                                       f.SLOT_NO,
                                       f.CARD_TYPE,
                                       f.TOTAL_PORT,
                                       f.PROJECT_NAME,
                                       f.PATCHCORD,
                                       f.DESIGN,
                                       f.ADDCARD,
                                       f.ADDSPLITTER,
                                       f.OWNER_FID,
                                       f.FLAG
                                   };
                #endregion

                using (Entities ctxNeps1 = new Entities())
                {
                    using (var xlsFile = new ExcelPackage())
                    {
                        #region site tab
                        if (querySITE.Count() > 0)
                        {
                            var prtID = (from a in ctxNeps1.WV_PRT_MAST
                                         where a.PRT_NAME == siteDetail.SITEPTT
                                         select a).FirstOrDefault();

                            var loadSite = xlsFile.Workbook.Worksheets.Add("Site");
                            loadSite.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            loadSite.Row(1).Style.WrapText = true;
                            loadSite.Row(1).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            loadSite.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            loadSite.Cells["A1:AF5"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            loadSite.Cells["A1:AF5"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            loadSite.Cells["A1:AF1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            loadSite.Cells["A1,E1:F1,I1:K1,Q1,T1:X1,Z1,AB1:AE1"].Style.Fill.BackgroundColor.SetColor(Color.Orange);
                            loadSite.Cells["B1:D1,G1:H1,L1:P1,R1:S1,Y1,AA1,AF1"].Style.Fill.BackgroundColor.SetColor(Color.Red);

                            loadSite.Cells["A1"].Value = "#";
                            loadSite.Cells["B1"].Value = "Site Name";
                            loadSite.Cells["C1"].Value = "Status";
                            loadSite.Cells["D1"].Value = "Exchange Abbr";
                            loadSite.Cells["E1"].Value = "Latitude (Degrees/Min/Sec)";
                            loadSite.Cells["F1"].Value = "Longitude (Degrees/Min/Sec)";
                            loadSite.Cells["G1"].Value = "Site Type";
                            loadSite.Cells["H1"].Value = "Parent Site ID";
                            loadSite.Cells["I1"].Value = "House Number, Unit Number";
                            loadSite.Cells["J1"].Value = "Floor Number";
                            loadSite.Cells["K1"].Value = "Building";
                            loadSite.Cells["L1"].Value = "Street Type";
                            loadSite.Cells["M1"].Value = "Street Name";
                            loadSite.Cells["N1"].Value = "Postcode";
                            loadSite.Cells["O1"].Value = "City";
                            loadSite.Cells["P1"].Value = "State";
                            loadSite.Cells["Q1"].Value = "Section / Taman";
                            loadSite.Cells["R1"].Value = "Country";
                            loadSite.Cells["S1"].Value = "Access Restrictions";
                            loadSite.Cells["T1"].Value = "Comments (Access Restrictions)";
                            loadSite.Cells["U1"].Value = "Contacts & Phone No";
                            loadSite.Cells["V1"].Value = "Comments (for SDF/DP etc)";
                            loadSite.Cells["W1"].Value = "Distance between ODF to FDC";
                            loadSite.Cells["W1:X1"].Merge = true;
                            loadSite.Cells["Y1"].Value = "Total Service Boundary";
                            loadSite.Cells["Z1"].Value = "Total Splitter in DB";
                            loadSite.Cells["AA1"].Value = "In Service Date";
                            loadSite.Cells["AB1"].Value = "Splitter In (For ODP Only)";
                            loadSite.Cells["AC1"].Value = "Splitter Out (For ODP Only)";
                            loadSite.Cells["AD1"].Value = "Smart Partnership";
                            loadSite.Cells["AE1"].Value = "Patch Cord";
                            loadSite.Cells["AF1"].Value = "Outdoor / Indoor Tagging";

                            loadSite.Cells["G2"].Value = "REGION";
                            loadSite.Cells["G3"].Value = "STATE";
                            loadSite.Cells["G4"].Value = "TELEKOM LOCAL OFFICE";
                            loadSite.Cells["G5"].Value = "CENTRAL OFFICE";

                            loadSite.Cells["H2"].Value = siteDetail.SITEREGION;
                            loadSite.Cells["H3"].Value = siteDetail.SITESTATE;
                            loadSite.Cells["H4"].Value = prtID.PRT_ID;
                            loadSite.Cells["H5"].Value = siteDetail.ADDRESSCITY;

                            int num = 6;
                            foreach (var c in querySITE)
                            {
                                loadSite.Cells["A" + num + ":AF" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                loadSite.Cells["A" + num + ":AF" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                loadSite.Cells["A" + num].Value = num - 5;
                                loadSite.Cells["C" + num].Value = "In Service";
                                loadSite.Cells["D" + num].Value = c.SITEEXC;
                                loadSite.Cells["G" + num].Value = c.SITETYPE;
                                loadSite.Cells["H" + num].Value = c.SITENAME;
                                loadSite.Cells["I" + num].Value = c.SITENO;
                                loadSite.Cells["J" + num].Value = c.SITEFLOOR;
                                loadSite.Cells["K" + num].Value = c.SITEBUILDING;
                                loadSite.Cells["L" + num].Value = c.STREETTYPE;
                                loadSite.Cells["M" + num].Value = c.ADDRESSSTREET;
                                loadSite.Cells["N" + num].Value = c.POSTCODE;
                                loadSite.Cells["O" + num].Value = c.ADDRESSCITY;
                                loadSite.Cells["P" + num].Value = c.ADDRESSSTATE;
                                loadSite.Cells["Q" + num].Value = c.COUNTY;
                                loadSite.Cells["R" + num].Value = c.COUNTRY;
                                loadSite.Cells["S" + num].Value = c.ACCESSRESTRICTION;
                                loadSite.Cells["T" + num].Value = c.COMMENTS;
                                loadSite.Cells["U" + num].Value = c.CONTACT;
                                loadSite.Cells["Y" + num].Value = c.TOTALSERVICEBOUNDARY;
                                loadSite.Cells["AA" + num].Value = c.INSERVICE;
                                loadSite.Cells["AE" + num++].Value = c.PATCHCORD;
                            }
                            loadSite.Cells.AutoFitColumns();
                            loadSite.Column(19).Width = 11;
                            loadSite.Column(20).Width = 11.11;
                            loadSite.Column(22).Width = 10.67;
                            loadSite.Column(30).Width = 10.44;
                        }
                        #endregion

                        #region sb tab
                        if (querySB.Count() > 0)
                        {
                            var loadSB = xlsFile.Workbook.Worksheets.Add("Service Boundary");
                            loadSB.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            loadSB.Cells["A1"].Style.WrapText = true;
                            loadSB.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            loadSB.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            loadSB.Row(2).Style.WrapText = true;
                            loadSB.Row(2).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            loadSB.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            loadSB.Cells["A1:AA2"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            loadSB.Cells["A1:AA2"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            loadSB.Cells["A1:AA2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            loadSB.Cells["A1:AA1,A2,C2:D2,G2,N2:P2,V2,X2:AA2"].Style.Fill.BackgroundColor.SetColor(Color.Orange);
                            loadSB.Cells["B2,E2:F2,H2:M2,Q2:U2,W2"].Style.Fill.BackgroundColor.SetColor(Color.Red);

                            loadSB.Cells["A1:A2"].Merge = true;
                            loadSB.Cells["B1"].Value = "Customer Premise";
                            loadSB.Cells["Q1"].Value = "DP Information";
                            loadSB.Cells["Z1"].Value = "Smart Partnership";
                            loadSB.Cells["B1,Q1,Z1"].Style.Font.Bold = true;
                            loadSB.Cells["B1:P1"].Merge = true;
                            loadSB.Cells["Q1:X1"].Merge = true;
                            loadSB.Cells["Z1:AA1"].Merge = true;
                            loadSB.Cells["B2"].Value = "House / Unit / Lot";
                            loadSB.Cells["C2"].Value = "Floor Number";
                            loadSB.Cells["D2"].Value = "Building Name";
                            loadSB.Cells["E2"].Value = "Street Type";
                            loadSB.Cells["F2"].Value = "Street Name";
                            loadSB.Cells["G2"].Value = "Section / Taman";
                            loadSB.Cells["H2"].Value = "Postcode";
                            loadSB.Cells["I2"].Value = "City";
                            loadSB.Cells["J2"].Value = "State";
                            loadSB.Cells["K2"].Value = "Country";
                            loadSB.Cells["L2"].Value = "Premise Type";
                            loadSB.Cells["M2"].Value = "Copper TM";
                            loadSB.Cells["N2"].Value = "Cable entry direction to the premise";
                            loadSB.Cells["O2"].Value = "Condition of underground cable";
                            loadSB.Cells["P2"].Value = "Route DP to Customer Premises";
                            loadSB.Cells["Q2"].Value = "FDC ID / MSAN ID";
                            loadSB.Cells["R2"].Value = "DP NO";
                            loadSB.Cells["S2"].Value = "DP Type";
                            loadSB.Cells["T2"].Value = "Status";
                            loadSB.Cells["U2"].Value = "Fiber Cable to Customer Premise Exist";
                            loadSB.Cells["V2"].Value = "DP Priority";
                            loadSB.Cells["W2"].Value = "DP Category";
                            loadSB.Cells["X2"].Value = "Comments";
                            loadSB.Cells["Y2"].Value = "Error";
                            loadSB.Cells["Z2"].Value = "Smart Developer";
                            loadSB.Cells["AA2"].Value = "Smart Project ID";

                            int num = 3;
                            foreach (var c in querySB)
                            {
                                loadSB.Cells["A" + num + ":AA" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                loadSB.Cells["A" + num + ":AA" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                loadSB.Cells["A" + num].Value = num - 2;
                                loadSB.Cells["B" + num].Value = c.SITENO;
                                loadSB.Cells["C" + num].Value = c.SITEFLOOR;
                                loadSB.Cells["D" + num].Value = c.SITEBUILDING;
                                loadSB.Cells["E" + num].Value = c.STREETYPE;
                                loadSB.Cells["F" + num].Value = c.STREETNAME;
                                loadSB.Cells["G" + num].Value = c.SECTION;
                                loadSB.Cells["H" + num].Value = c.POSTCODE;
                                loadSB.Cells["I" + num].Value = c.CITY;
                                loadSB.Cells["J" + num].Value = c.STATE;
                                loadSB.Cells["K" + num].Value = c.COUNTRY;
                                loadSB.Cells["L" + num].Value = c.PREMISETYPE;
                                loadSB.Cells["M" + num].Value = c.COPPEROWNBYTM;
                                loadSB.Cells["Q" + num].Value = c.SITESDP;
                                loadSB.Cells["T" + num].Value = "In Service";
                                loadSB.Cells["U" + num].Value = c.FIBERTOPREMISEEXIST;
                                loadSB.Cells["Z" + num].Value = c.SMARTDEVELOPER;
                                loadSB.Cells["AA" + num++].Value = c.SMARTPROJECTID;
                            }
                            loadSB.Cells.AutoFitColumns();
                            loadSB.Column(15).Width = 11.44;
                            loadSB.Column(24).Width = 10.78;
                        }
                        #endregion

                        #region oltOdf tab
                        var oltODF = xlsFile.Workbook.Worksheets.Add("OLT - ODF");
                        oltODF.Column(1).Width = 37.78;
                        oltODF.Column(2).Width = 9.78;
                        oltODF.Column(3).Width = 11.89;
                        oltODF.Column(4).Width = 9.22;
                        oltODF.Column(5).Width = 9.33;
                        oltODF.Column(6).Width = 9.78;
                        oltODF.Column(7).Width = 2.67;
                        oltODF.Column(8).Width = 13.67;
                        oltODF.Column(9).Width = 20.56;
                        oltODF.Column(10).Width = 20.78;
                        oltODF.Column(11).Width = 13.22;
                        oltODF.Column(12).Width = 13.89;
                        oltODF.Column(14).Width = 12.67;
                        oltODF.Column(15).Width = 14.56;
                        oltODF.Column(16).Width = 10.78;
                        oltODF.Column(17).Width = 14.67;
                        oltODF.Column(18).Width = 9.67;
                        oltODF.Column(19).Width = 9.33;
                        oltODF.Column(20).Width = 9.56;
                        oltODF.Column(21).Width = 21.67;
                        oltODF.Column(22).Width = 26.89;
                        oltODF.Column(23).Width = 16.33;
                        oltODF.Row(26).Style.WrapText = true;
                        oltODF.Row(26).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        oltODF.Cells["C2:F25,D26,J2:K25"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        oltODF.Cells["A1:K26,L26:X26"].Style.Font.Bold = true;
                        oltODF.Cells["A1:K26,L26:X26"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        oltODF.Cells["A1:F26,H1:K26,L25:X26"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        oltODF.Cells["A1:K1,A2:B25,G2:I25,A26:X26"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        oltODF.Cells["A1:F1,A2:B25,A26:F26"].Style.Fill.BackgroundColor.SetColor(Color.PaleTurquoise);
                        oltODF.Cells["H1:K1,H2:I25,H26:V26"].Style.Fill.BackgroundColor.SetColor(Color.Thistle);
                        oltODF.Cells["W26:X26"].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                        oltODF.Cells["G1:G26"].Style.Fill.BackgroundColor.SetColor(Color.Orange);

                        #region olt details
                        oltODF.Cells["A1:F1"].Merge = true;
                        oltODF.Cells["A2"].Value = "OLT ID";
                        oltODF.Cells["A3"].Value = "OLT Status";
                        oltODF.Cells["A4"].Value = "Equipment Region";
                        oltODF.Cells["A5"].Value = "State";
                        oltODF.Cells["A6"].Value = "PTT";
                        oltODF.Cells["A7"].Value = "Central Office Location";
                        oltODF.Cells["A8"].Value = "Vendor";
                        oltODF.Cells["A9"].Value = "Supplier";
                        oltODF.Cells["A10"].Value = "No of Shelf";
                        oltODF.Cells["A11"].Value = "In Service Date (COA Date)";
                        oltODF.Cells["A12"].Value = "Serial No (Chasis No)";
                        oltODF.Cells["A13"].Value = "GEMS ID (Asset No)";
                        oltODF.Cells["A14"].Value = "Equipment IP";
                        oltODF.Cells["A15"].Value = "Installer Name (Project Manager)";
                        oltODF.Cells["A16"].Value = "IP Address";
                        oltODF.Cells["A17"].Value = "MetroE Main Core No";
                        oltODF.Cells["A18"].Value = "MetroE Protection Core No";
                        oltODF.Cells["A19"].Value = "NPE Name";
                        oltODF.Cells["A20"].Value = "NPE Shelf/Slot/Subslot/Port";
                        oltODF.Cells["A21"].Value = "NPE Standby Shelf/Slot/Subslot/Port";
                        oltODF.Cells["A22"].Value = "OLT UPLINK Shelf/Slot/Port";
                        oltODF.Cells["A23"].Value = "OLT UPLINK Card Type";
                        oltODF.Cells["A24"].Value = "MetroE Fiber Code";
                        oltODF.Cells["A25"].Value = "Comments";

                        // olt details data here
                        for (int i = 2; i < 26; i++)
                        {
                            oltODF.Cells["A" + i + ":B" + i].Merge = true;
                            if (i >= 20 && i <= 22) continue;
                            oltODF.Cells["C" + i + ":F" + i].Merge = true;
                        }
                        oltODF.Cells["D22:E22"].Merge = true;
                        #endregion

                        #region olt table
                        oltODF.Cells["A26"].Value = "OLT ID";
                        oltODF.Cells["B26"].Value = "Shelf No";
                        oltODF.Cells["C26"].Value = "Card/Slot No";
                        oltODF.Cells["D26"].Value = "Port No";
                        oltODF.Cells["D26:F26"].Merge = true;

                        if (queryOLT.Count() > 0)
                        {
                            int num = 27;
                            foreach (var c in queryOLT)
                            {
                                var shortZname = c.ZNAME.Split('_');
                                oltODF.Cells["A" + num + ":X" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                oltODF.Cells["A" + num + ":F" + num + ",H" + num + ":X" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                oltODF.Cells["A" + num + ":X" + num].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                oltODF.Cells["A" + num + ":X" + num].Style.Fill.BackgroundColor.SetColor(Color.Orange);
                                oltODF.Cells["A" + num].Value = shortZname[1];
                                oltODF.Cells["C" + num].Value = c.ZCARD;
                                oltODF.Cells["D" + num].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                oltODF.Cells["D" + num].Value = c.ZPORT;
                                oltODF.Cells["D" + num + ":F" + num++].Merge = true;
                            }
                        }
                        #endregion

                        #region odf details
                        oltODF.Cells["H1"].Value = "ODF";
                        oltODF.Cells["H1"].Style.Font.Size = 24;
                        oltODF.Cells["H1:K1"].Merge = true;
                        oltODF.Cells["H2"].Value = "ODF ID";
                        oltODF.Cells["H3"].Value = "ODF Status";
                        oltODF.Cells["H4"].Value = "Equipment Region";
                        oltODF.Cells["H5"].Value = "State";
                        oltODF.Cells["H6"].Value = "PTT";
                        oltODF.Cells["H7"].Value = "Central Office Location";
                        oltODF.Cells["H8"].Value = "Vendor";
                        oltODF.Cells["H9"].Value = "Supplier";
                        oltODF.Cells["H10"].Value = "No of Shelf";
                        oltODF.Cells["H11"].Value = "In Service Date";
                        oltODF.Cells["H12"].Value = "Serial No";
                        oltODF.Cells["H13"].Value = "GEMS ID (Asset No)";
                        oltODF.Cells["H14"].Value = "Contact Person";
                        oltODF.Cells["H15"].Value = "Comments";

                        // odf details data here
                        for (int i = 2; i < 26; i++)
                        {
                            oltODF.Cells["H" + i + ":I" + i].Merge = true;
                            oltODF.Cells["J" + i + ":K" + i].Merge = true;
                        }
                        #endregion

                        #region odf table
                        oltODF.Cells["H26"].Value = "ODF ID";
                        oltODF.Cells["I26"].Value = "Diverse route connection to Distribution NE?";
                        oltODF.Cells["J26"].Value = "Diverse route connection to Distribution NE (SPUR)?";
                        oltODF.Cells["K26"].Value = "Main Cable Code";
                        oltODF.Cells["L26"].Value = "Main Cable core no";
                        oltODF.Cells["M26"].Value = "Main Shelf No";
                        oltODF.Cells["N26"].Value = "Main Card/Slot No";
                        oltODF.Cells["O26"].Value = "Main Capacity (no. of core)";
                        oltODF.Cells["P26"].Value = "Main ODF Port NO";
                        oltODF.Cells["Q26"].Value = "Protection Cable Code";
                        oltODF.Cells["R26"].Value = "Protection Cable core no";
                        oltODF.Cells["S26"].Value = "Protection Shelf No";
                        oltODF.Cells["T26"].Value = "Protection Card/Slot no";
                        oltODF.Cells["U26"].Value = "Protection Capacity (no. of core)";
                        oltODF.Cells["V26"].Value = "Protection ODF Port NO";
                        oltODF.Cells["W26"].Value = "Destination";
                        oltODF.Cells["X26"].Value = "1st Layer splitter No";
                        #endregion
                        #endregion

                        #region fdc tab
                        var loadFDC = xlsFile.Workbook.Worksheets.Add("FDC");
                        loadFDC.Column(1).Width = 16.56;
                        loadFDC.Column(2).Width = 16.11;
                        loadFDC.Column(3).Width = 8.33;
                        loadFDC.Column(4).Width = 14.11;
                        loadFDC.Column(5).Width = 9.33;
                        loadFDC.Column(6).Width = 10.33;
                        loadFDC.Column(7).Width = 20.78;
                        loadFDC.Column(8).Width = 9.56;
                        loadFDC.Column(9).Width = 11.67;
                        loadFDC.Column(10).Width = 12.56;
                        loadFDC.Column(11).Width = 15.11;
                        loadFDC.Row(16).Style.WrapText = true;
                        loadFDC.Row(16).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        loadFDC.Cells["E1:E14,B15:J16,A16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        loadFDC.Cells["B1:E14,A16:J16"].Style.Font.Bold = true;
                        loadFDC.Cells["A1:J16,M1:M14"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        loadFDC.Cells["B1:M14,A15:J16"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        loadFDC.Cells["B1:D14,E7:M8,B15:J16,A16"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        loadFDC.Cells["B1:D14"].Style.Fill.BackgroundColor.SetColor(Color.PaleTurquoise);
                        loadFDC.Cells["B16:D16"].Style.Fill.BackgroundColor.SetColor(Color.Thistle);
                        loadFDC.Cells["B15:J15,E16:J16,A16"].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                        loadFDC.Cells["E7:M8"].Style.Fill.BackgroundColor.SetColor(Color.Orange);

                        #region fdc details
                        loadFDC.Cells["B1"].Value = "FDC ID";
                        loadFDC.Cells["B2"].Value = "FDC Status";
                        loadFDC.Cells["B3"].Value = "Equipment Region";
                        loadFDC.Cells["B4"].Value = "State";
                        loadFDC.Cells["B5"].Value = "PTT";
                        loadFDC.Cells["B6"].Value = "Central Office Location";
                        loadFDC.Cells["B7"].Value = "Vendor";
                        loadFDC.Cells["B8"].Value = "Supplier";
                        loadFDC.Cells["B9"].Value = "FDC Model";
                        loadFDC.Cells["B10"].Value = "In Service Date";
                        loadFDC.Cells["B11"].Value = "Serial No";
                        loadFDC.Cells["B12"].Value = "GEMS ID (Asset No)";
                        loadFDC.Cells["B13"].Value = "Installer Name (Project Manager)";
                        loadFDC.Cells["B14"].Value = "Comments";

                        // fdc details data here
                        for (int i = 1; i < 15; i++)
                        {
                            loadFDC.Cells["B" + i + ":D" + i].Merge = true;
                            loadFDC.Cells["E" + i + ":M" + i].Merge = true;
                        }
                        #endregion

                        #region eside table
                        loadFDC.Cells["B15"].Value = "Eside (A Side)";
                        loadFDC.Cells["B15"].Style.Font.Size = 24;
                        loadFDC.Cells["B15:J15"].Merge = true;
                        loadFDC.Cells["A16"].Value = "FDC ID";
                        loadFDC.Cells["B16"].Value = "ODF ID";
                        loadFDC.Cells["C16"].Value = "ODF Shelf No";
                        loadFDC.Cells["D16"].Value = "ODF Port NO";
                        loadFDC.Cells["E16"].Value = "Port No";
                        loadFDC.Cells["F16"].Value = "Status";
                        loadFDC.Cells["G16"].Value = "Cable Code";
                        loadFDC.Cells["H16"].Value = "Core No";
                        loadFDC.Cells["I16"].Value = "Splitter No / Type";
                        loadFDC.Cells["J16"].Value = "In Service Date";
                        #endregion
                        #endregion

                        #region ipmsan tab
                        var loadIPMSAN = xlsFile.Workbook.Worksheets.Add("IPMSAN");
                        loadIPMSAN.Column(2).Width = 23.89;
                        loadIPMSAN.Column(3).Width = 16.11;
                        loadIPMSAN.Column(4).Width = 14.89;
                        loadIPMSAN.Column(5).Width = 18.33;
                        loadIPMSAN.Column(6).Width = 17.89;
                        loadIPMSAN.Column(7).Width = 9.22;
                        loadIPMSAN.Column(8).Width = 12.89;
                        loadIPMSAN.Column(9).Width = 22.33;
                        loadIPMSAN.Column(10).Width = 11.67;
                        loadIPMSAN.Column(11).Width = 17.33;
                        loadIPMSAN.Column(12).Width = 16.89;
                        loadIPMSAN.Column(13).Width = 9.67;
                        loadIPMSAN.Column(14).Width = 8.56;
                        loadIPMSAN.Column(15).Width = 8.33;
                        loadIPMSAN.Column(16).Width = 20.56;
                        loadIPMSAN.Column(17).Width = 19.22;
                        loadIPMSAN.Column(18).Width = 8.33;
                        loadIPMSAN.Column(19).Width = 10.67;
                        loadIPMSAN.Column(20).Width = 11.67;
                        loadIPMSAN.Column(21).Width = 10.22;
                        loadIPMSAN.Column(22).Width = 14.78;
                        loadIPMSAN.Column(23).Width = 13.67;
                        loadIPMSAN.Column(24).Width = 11.11;
                        loadIPMSAN.Column(25).Width = 25.89;
                        loadIPMSAN.Column(26).Width = 8.33;
                        loadIPMSAN.Column(27).Width = 32.56;
                        loadIPMSAN.Row(26).Style.WrapText = true;
                        loadIPMSAN.Row(26).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        loadIPMSAN.Cells["E1:O25,B26:AA28"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        loadIPMSAN.Cells["B1:O25,B26:AA28"].Style.Font.Bold = true;
                        loadIPMSAN.Cells["A1:O25,A26:AA28"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        loadIPMSAN.Cells["B1:O24,B25:AA28"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        loadIPMSAN.Cells["B1:D25,E1,E10,E15,E24,B26:AA28"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        loadIPMSAN.Cells["B1:D25"].Style.Fill.BackgroundColor.SetColor(Color.PaleTurquoise);
                        loadIPMSAN.Cells["N26:W27,O28,Q28:R28,T28,V28:W28"].Style.Fill.BackgroundColor.SetColor(Color.Thistle);
                        loadIPMSAN.Cells["I26,K26:M28,X26,AA26"].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                        loadIPMSAN.Cells["E1,E10,E15,E24,B26:H28,J26,N28,P28,S28,U28,Y26:Z28"].Style.Fill.BackgroundColor.SetColor(Color.Red);

                        #region ipmsan details
                        loadIPMSAN.Cells["B1"].Value = "IPMSAN ID";
                        loadIPMSAN.Cells["B2"].Value = "IPMSAN Status";
                        loadIPMSAN.Cells["B3"].Value = "Equipment Region";
                        loadIPMSAN.Cells["B4"].Value = "State";
                        loadIPMSAN.Cells["B5"].Value = "PTT";
                        loadIPMSAN.Cells["B6"].Value = "Central Office Location";
                        loadIPMSAN.Cells["B7"].Value = "Vendor";
                        loadIPMSAN.Cells["B8"].Value = "Supplier";
                        loadIPMSAN.Cells["B9"].Value = "IPMSAN Model";
                        loadIPMSAN.Cells["B10"].Value = "No of Shelf";
                        loadIPMSAN.Cells["B11"].Value = "In Service Date";
                        loadIPMSAN.Cells["B12"].Value = "Serial No";
                        loadIPMSAN.Cells["B13"].Value = "GEMS ID (Asset No)";
                        loadIPMSAN.Cells["B14"].Value = "Installer Name (Project Manager)";
                        loadIPMSAN.Cells["B15"].Value = "IP Address";
                        loadIPMSAN.Cells["B16"].Value = "MetroE Main Core No";
                        loadIPMSAN.Cells["B17"].Value = "MetroE Protection Core No";
                        loadIPMSAN.Cells["B18"].Value = "NPE Name";
                        loadIPMSAN.Cells["B19"].Value = "NPE Shelf/Slot/Subslot/Port";
                        loadIPMSAN.Cells["B20"].Value = "NPE Standby Shelf/Slot/Subslot/Port";
                        loadIPMSAN.Cells["B21"].Value = "OLT UPLINK Shelf/Slot/Port";
                        loadIPMSAN.Cells["B22"].Value = "OLT UPLINK Card Type";
                        loadIPMSAN.Cells["B23"].Value = "MetroE Fiber Code";
                        loadIPMSAN.Cells["B24"].Value = "Origin Equipment ID";
                        loadIPMSAN.Cells["B25"].Value = "Comments";

                        // ipmsan details data here
                        for (int i = 1; i < 26; i++)
                        {
                            loadIPMSAN.Cells["B" + i + ":D" + i].Merge = true;
                            if (i >= 19 && i <= 21)
                            {
                                loadIPMSAN.Cells["E" + i + ":G" + i].Merge = true;
                                loadIPMSAN.Cells["H" + i + ":I" + i].Merge = true;
                                loadIPMSAN.Cells["J" + i + ":L" + i].Merge = true;
                                loadIPMSAN.Cells["M" + i + ":O" + i].Merge = true;
                            }
                            else
                                loadIPMSAN.Cells["E" + i + ":O" + i].Merge = true;
                        }
                        #endregion

                        #region ipmsan table
                        loadIPMSAN.Cells["B26"].Value = "IPMSAN ID";
                        loadIPMSAN.Cells["C26"].Value = "Port No Origin";
                        loadIPMSAN.Cells["D26"].Value = "Shelf No RT";
                        loadIPMSAN.Cells["E26"].Value = "Slot No RT";
                        loadIPMSAN.Cells["F26"].Value = "Port No RT";
                        loadIPMSAN.Cells["G26"].Value = "VCI No";
                        loadIPMSAN.Cells["H26"].Value = "Card Type";
                        loadIPMSAN.Cells["I26"].Value = "Status";
                        loadIPMSAN.Cells["J26"].Value = "SERVICE CARD";
                        loadIPMSAN.Cells["K26"].Value = "Card Role";
                        loadIPMSAN.Cells["L26"].Value = "Port Model";
                        loadIPMSAN.Cells["M26"].Value = "Card Model";
                        loadIPMSAN.Cells["N26"].Value = "ODF Shelf No";
                        loadIPMSAN.Cells["N26:W26"].Merge = true;
                        loadIPMSAN.Cells["N27"].Value = "Port No";
                        loadIPMSAN.Cells["N27:R27"].Merge = true;
                        loadIPMSAN.Cells["S27"].Value = "Status";
                        loadIPMSAN.Cells["S27:W27"].Merge = true;
                        loadIPMSAN.Cells["N28"].Value = "Cable Code";
                        loadIPMSAN.Cells["O28"].Value = "Core No";
                        loadIPMSAN.Cells["P28"].Value = "Splitter No / Type";
                        loadIPMSAN.Cells["Q28"].Value = "In Service Date";
                        loadIPMSAN.Cells["R28"].Value = "FDC ID";
                        loadIPMSAN.Cells["S28"].Value = "ODF ID";
                        loadIPMSAN.Cells["T28"].Value = "ODF Shelf No";
                        loadIPMSAN.Cells["U28"].Value = "ODF Port NO";
                        loadIPMSAN.Cells["V28"].Value = "Port No";
                        loadIPMSAN.Cells["W28"].Value = "Status";
                        loadIPMSAN.Cells["X26"].Value = "Cable Code";
                        loadIPMSAN.Cells["Y26"].Value = "Core No";
                        loadIPMSAN.Cells["Z26"].Value = "Splitter No / Type";
                        loadIPMSAN.Cells["AA26"].Value = "In Service Date";
                        #endregion
                        #endregion

                        #region loadequip
                        if (queryLOADEQUIP.Count() > 0)
                        {
                            var LoadEquip = xlsFile.Workbook.Worksheets.Add("Load Equipment");
                            LoadEquip.Cells.Style.Font.Name = "Arial";
                            LoadEquip.Cells.Style.Font.Size = 8;
                            for (int i = 1; i < 2; i++)
                                LoadEquip.Column(i).Width = 11.00; //A B 
                            LoadEquip.Column(3).Width = 8.50; //C
                            LoadEquip.Column(4).Width = 19.00; //D
                            LoadEquip.Column(5).Width = 8.00; //E
                            LoadEquip.Column(6).Width = 16.00; //F
                            LoadEquip.Column(7).Width = 8.00; //G
                            LoadEquip.Column(8).Width = 10.00; //H
                            LoadEquip.Column(9).Width = 8.00; //I
                            LoadEquip.Column(10).Width = 8.00; //J
                            LoadEquip.Column(11).Width = 19.00; //K
                            LoadEquip.Column(12).Width = 16.00; //L
                            LoadEquip.Column(13).Width = 14.00; //M
                            LoadEquip.Column(14).Width = 11.00; //N
                            LoadEquip.Column(15).Width = 43.00; //O
                            LoadEquip.Column(16).Width = 9.00; //P
                            LoadEquip.Column(17).Width = 9.00; //Q
                            LoadEquip.Column(18).Width = 21.00; //R
                            LoadEquip.Column(19).Width = 16.00; //S
                            LoadEquip.Column(20).Width = 13.00; //T3
                            LoadEquip.Column(21).Width = 12.00; //U
                            LoadEquip.Column(22).Width = 10.00; //V
                            LoadEquip.Column(23).Width = 10.00; //W


                            LoadEquip.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            LoadEquip.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                            LoadEquip.Cells["A1:W1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                            LoadEquip.Cells["A1:W1"].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                            LoadEquip.Cells["A1:W1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                            LoadEquip.Cells["A1:W1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                            LoadEquip.Cells["A1:W1"].Style.Font.Bold = true;

                            LoadEquip.Cells["A1"].Value = "LOADEQUIP_ID";
                            LoadEquip.Cells["B1"].Value = "CLASS_NAME";
                            LoadEquip.Cells["C1"].Value = "G3E_FID";
                            LoadEquip.Cells["D1"].Value = "SCHEME_NAME";
                            LoadEquip.Cells["E1"].Value = "USERID";
                            LoadEquip.Cells["F1"].Value = "EQUIPID";
                            LoadEquip.Cells["G1"].Value = "EQUIPCAT";
                            LoadEquip.Cells["H1"].Value = "EQUIPVEND";
                            LoadEquip.Cells["I1"].Value = "EQUIMODEL";
                            LoadEquip.Cells["J1"].Value = "REGION";
                            LoadEquip.Cells["K1"].Value = "STATE";
                            LoadEquip.Cells["L1"].Value = "EXCDESC";
                            LoadEquip.Cells["M1"].Value = "SITE";
                            LoadEquip.Cells["N1"].Value = "MNGTIP";
                            LoadEquip.Cells["O1"].Value = "TEMPNAME";
                            LoadEquip.Cells["P1"].Value = "INSERVICE";
                            LoadEquip.Cells["Q1"].Value = "TAGGING";
                            LoadEquip.Cells["R1"].Value = "OUTDOORINDOORTAGGING";
                            LoadEquip.Cells["S1"].Value = "DPEXTENSIONFLAG";
                            LoadEquip.Cells["T1"].Value = "NETWORKNAME";
                            LoadEquip.Cells["U1"].Value = "GRN_STATUS";
                            LoadEquip.Cells["V1"].Value = "PATCHCORD";
                            LoadEquip.Cells["W1"].Value = "FLAG";

                            int num = 2;
                            foreach (var c in queryLOADEQUIP)
                            {
                                LoadEquip.Cells["A" + num + ":W" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                LoadEquip.Cells["A" + num + ":W" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                LoadEquip.Cells["A" + num].Value = c.LOADEQUIP_ID;
                                LoadEquip.Cells["B" + num].Value = c.CLASS_NAME;
                                LoadEquip.Cells["C" + num].Value = c.G3E_FID;
                                LoadEquip.Cells["D" + num].Value = c.SCHEME_NAME;
                                LoadEquip.Cells["E" + num].Value = c.USERID;
                                LoadEquip.Cells["F" + num].Value = c.EQUIPID;
                                LoadEquip.Cells["G" + num].Value = c.EQUIPCAT;
                                LoadEquip.Cells["H" + num].Value = c.EQUIPVEND;
                                LoadEquip.Cells["I" + num].Value = c.EQUIMODEL;
                                LoadEquip.Cells["J" + num].Value = c.REGION;
                                LoadEquip.Cells["K" + num].Value = c.STATE;
                                LoadEquip.Cells["L" + num].Value = c.EXCDESC;
                                LoadEquip.Cells["M" + num].Value = c.SITE;
                                LoadEquip.Cells["N" + num].Value = c.MNGTIP;
                                LoadEquip.Cells["O" + num].Value = c.TEMPNAME;
                                LoadEquip.Cells["P" + num].Value = c.INSERVICE;
                                LoadEquip.Cells["Q" + num].Value = c.TAGGING;
                                LoadEquip.Cells["R" + num].Value = c.OUTDOORINDOORTAGGING;
                                LoadEquip.Cells["S" + num].Value = c.DPEXTENSIONFLAG;
                                LoadEquip.Cells["T" + num].Value = c.NETWORKNAME;
                                LoadEquip.Cells["U" + num].Value = c.GRN_STATUS;
                                LoadEquip.Cells["V" + num].Value = c.PATCHCORD;
                                LoadEquip.Cells["W" + num++].Value = c.FLAG;

                            }
                        }
                        #endregion

                        #region loadsite
                        if (queryLOADSITE.Count() > 0)
                        {
                            var LSite = xlsFile.Workbook.Worksheets.Add("Load Site");
                            LSite.Cells.Style.Font.Name = "Arial";
                            LSite.Cells.Style.Font.Size = 8;
                            LSite.Column(1).Width = 10.00; //A.
                            LSite.Column(2).Width = 10.00; //B
                            LSite.Column(3).Width = 19.00; //C
                            for (int i = 4; i < 7; i++)
                                LSite.Column(i).Width = 9.00; //D E F
                            LSite.Column(7).Width = 20.00; //G
                            LSite.Column(8).Width = 13.00; //H
                            LSite.Column(9).Width = 8.00; //I
                            LSite.Column(10).Width = 13.00; //J
                            for (int i = 11; i < 15; i++)
                                LSite.Column(i).Width = 8.00; //K L F M N
                            LSite.Column(15).Width = 10.00; //O
                            LSite.Column(16).Width = 16.00; //P
                            LSite.Column(17).Width = 14.00; //Q
                            LSite.Column(18).Width = 19.00; //R
                            LSite.Column(19).Width = 8.00; //S
                            LSite.Column(20).Width = 9.00; //T
                            LSite.Column(21).Width = 18.00; //U
                            LSite.Column(22).Width = 8.00; //V
                            LSite.Column(23).Width = 11.00; //W
                            LSite.Column(24).Width = 18.00; //X
                            LSite.Column(25).Width = 15.00; //Y
                            LSite.Column(26).Width = 13.50; //Z
                            LSite.Column(27).Width = 16.00; //AA
                            LSite.Column(28).Width = 19.00; //AB
                            LSite.Column(29).Width = 16.00; //AC
                            for (int i = 30; i < 34; i++)
                                LSite.Column(i).Width = 8.00; //AD AE AF AG

                            LSite.Cells["A1"].Value = "LOADSITE_ID";
                            LSite.Cells["B1"].Value = "CLASS_NAME";
                            LSite.Cells["C1"].Value = "SCHEME_NAME";
                            LSite.Cells["D1"].Value = "G3E_FID";
                            LSite.Cells["E1"].Value = "USERID";
                            LSite.Cells["F1"].Value = "SITEREGION";
                            LSite.Cells["G1"].Value = "SITESTATE";
                            LSite.Cells["H1"].Value = "SITEPTT";
                            LSite.Cells["I1"].Value = "SITEEXC";
                            LSite.Cells["J1"].Value = "SITENAME";
                            LSite.Cells["K1"].Value = "SITEDESC";
                            LSite.Cells["L1"].Value = "SITETYPE";
                            LSite.Cells["M1"].Value = "SITENO";
                            LSite.Cells["N1"].Value = "SITEFLOOR";
                            LSite.Cells["O1"].Value = "SITEBUILDING";
                            LSite.Cells["P1"].Value = "ADDRESSSTREET";
                            LSite.Cells["Q1"].Value = "ADDRESSCITY";
                            LSite.Cells["R1"].Value = "ADDRESSSTATE";
                            LSite.Cells["S1"].Value = "POSTCODE";
                            LSite.Cells["T1"].Value = "STREETTYPE";
                            LSite.Cells["U1"].Value = "COUNTY";
                            LSite.Cells["V1"].Value = "COUNTRY";
                            LSite.Cells["W1"].Value = "SITECOMMENT";
                            LSite.Cells["X1"].Value = "EQUIPMENTLOCATION";
                            LSite.Cells["Y1"].Value = "CABLINGTYPE";
                            LSite.Cells["Z1"].Value = "COPPEROWNBYTM";
                            LSite.Cells["AA1"].Value = "FIBERTOPREMISEEXIST";
                            LSite.Cells["AB1"].Value = "TOTALSERVICEBOUNDARY";
                            LSite.Cells["AC1"].Value = "ACCESSRESTRICTION";
                            LSite.Cells["AD1"].Value = "CONTACT";
                            LSite.Cells["AE1"].Value = "COMMENTS";
                            LSite.Cells["AF1"].Value = "GRN_STATUS";
                            LSite.Cells["AG1"].Value = "FLAG";

                            LSite.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            LSite.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                            LSite.Cells["A1:AG1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                            LSite.Cells["A1:AG1"].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                            LSite.Cells["A1:AG1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                            LSite.Cells["A1:AG1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                            LSite.Cells["A1:AG1"].Style.Font.Bold = true;

                            int num = 2;
                            foreach (var c in queryLOADSITE)
                            {
                                LSite.Cells["A" + num + ":AG" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                LSite.Cells["A" + num + ":AG" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                LSite.Cells["A" + num].Value = c.LOADSITE_ID;
                                LSite.Cells["B" + num].Value = c.CLASS_NAME;
                                LSite.Cells["C" + num].Value = c.SCHEME_NAME;
                                LSite.Cells["D" + num].Value = c.G3E_FID;
                                LSite.Cells["E" + num].Value = c.USERID;
                                LSite.Cells["F" + num].Value = c.SITEREGION;
                                LSite.Cells["G" + num].Value = c.SITESTATE;
                                LSite.Cells["H" + num].Value = c.SITEPTT;
                                LSite.Cells["I" + num].Value = c.SITEEXC;
                                LSite.Cells["J" + num].Value = c.SITENAME;
                                LSite.Cells["K" + num].Value = c.SITEDESC;
                                LSite.Cells["L" + num].Value = c.SITETYPE;
                                LSite.Cells["M" + num].Value = c.SITENO;
                                LSite.Cells["N" + num].Value = c.SITEFLOOR;
                                LSite.Cells["O" + num].Value = c.SITEBUILDING;
                                LSite.Cells["P" + num].Value = c.ADDRESSSTREET;
                                LSite.Cells["Q" + num].Value = c.ADDRESSCITY;
                                LSite.Cells["R" + num].Value = c.ADDRESSSTATE;
                                LSite.Cells["S" + num].Value = c.POSTCODE;
                                LSite.Cells["T" + num].Value = c.STREETTYPE;
                                LSite.Cells["U" + num].Value = c.COUNTY;
                                LSite.Cells["V" + num].Value = c.COUNTRY;
                                LSite.Cells["W" + num].Value = c.SITECOMMENT;
                                LSite.Cells["X" + num].Value = c.EQUIPMENTLOCATION;
                                LSite.Cells["Y" + num].Value = c.CABLINGTYPE;
                                LSite.Cells["Z" + num].Value = c.COPPEROWNBYTM;
                                LSite.Cells["AA" + num].Value = c.FIBERTOPREMISEEXIST;
                                LSite.Cells["AB" + num].Value = c.TOTALSERVICEBOUNDARY;
                                LSite.Cells["AC" + num].Value = c.ACCESSRESTRICTION;
                                LSite.Cells["AD" + num].Value = c.CONTACT;
                                LSite.Cells["AE" + num].Value = c.COMMENTS;
                                LSite.Cells["AF" + num].Value = c.GRN_STATUS;
                                LSite.Cells["AG" + num++].Value = c.FLAG;
                            }

                        }
                        #endregion

                        #region loadservbound
                        if (queryLOADSERVBOUND.Count() > 0)
                        {
                            var LServBound = xlsFile.Workbook.Worksheets.Add("Load Service Boundary");
                            LServBound.Cells.Style.Font.Name = "Arial";
                            LServBound.Cells.Style.Font.Size = 8;
                            LServBound.Column(1).Width = 14.78; //A
                            LServBound.Column(2).Width = 8.00; //B
                            LServBound.Column(3).Width = 20.00; //C
                            LServBound.Column(4).Width = 6.00; //D
                            LServBound.Column(5).Width = 6.00; //E
                            LServBound.Column(6).Width = 14.00; //F
                            for (int i = 7; i < 11; i++)
                                LServBound.Column(i).Width = 10.00; //G H I J
                            LServBound.Column(11).Width = 16.00; //K
                            LServBound.Column(12).Width = 16.00; //L
                            LServBound.Column(13).Width = 8.00; //M
                            LServBound.Column(14).Width = 12.00; //N
                            LServBound.Column(15).Width = 19.00; //O
                            LServBound.Column(16).Width = 8.00; //P
                            LServBound.Column(17).Width = 21.00; //Q
                            LServBound.Column(18).Width = 16.00; //R
                            LServBound.Column(19).Width = 13.00; //S
                            LServBound.Column(20).Width = 8.00; //T
                            LServBound.Column(21).Width = 13.00; //U
                            LServBound.Column(22).Width = 13.50; //V
                            LServBound.Column(23).Width = 13.50; //W
                            LServBound.Column(24).Width = 10.00; //X
                            LServBound.Column(25).Width = 8.00; //Y
                            LServBound.Column(26).Width = 17.00; //Z
                            LServBound.Column(27).Width = 7.00; //AA

                            LServBound.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            LServBound.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                            LServBound.Cells["A1:AA1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                            LServBound.Cells["A1:AA1"].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                            LServBound.Cells["A1:AA1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                            LServBound.Cells["A1:AA1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                            LServBound.Cells["A1:AA1"].Style.Font.Bold = true;


                            LServBound.Cells["A1"].Value = "LOADSERVBOUND_ID";
                            LServBound.Cells["B1"].Value = "G3E_FID";
                            LServBound.Cells["C1"].Value = "SCHEME_NAME";
                            LServBound.Cells["D1"].Value = "USERID";
                            LServBound.Cells["E1"].Value = "G3E_ID";
                            LServBound.Cells["F1"].Value = "SERVICEBOUNDARYID";
                            LServBound.Cells["G1"].Value = "SITENO";
                            LServBound.Cells["H1"].Value = "SITEFLOOR";
                            LServBound.Cells["I1"].Value = "SITEBUILDING";
                            LServBound.Cells["J1"].Value = "STREETYPE";
                            LServBound.Cells["K1"].Value = "STREETNAME";
                            LServBound.Cells["L1"].Value = "SECTION";
                            LServBound.Cells["M1"].Value = "POSTCODE";
                            LServBound.Cells["N1"].Value = "CITY";
                            LServBound.Cells["O1"].Value = "STATE";
                            LServBound.Cells["P1"].Value = "COUNTRY";
                            LServBound.Cells["Q1"].Value = "PREMISETYPE";
                            LServBound.Cells["R1"].Value = "SERVCATEGORY";
                            LServBound.Cells["S1"].Value = "SITESDP";
                            LServBound.Cells["T1"].Value = "SITEDPEXC";
                            LServBound.Cells["U1"].Value = "DDP";
                            LServBound.Cells["V1"].Value = "SMARTDEVELOPER";
                            LServBound.Cells["W1"].Value = "SMARTPROJECTID";
                            LServBound.Cells["X1"].Value = "GRN_STATUS";
                            LServBound.Cells["Y1"].Value = "COS";
                            LServBound.Cells["Z1"].Value = "DROPCABLEDISTANCE";
                            LServBound.Cells["AA1"].Value = "FLAG";

                            int num = 2;
                            foreach (var c in queryLOADSERVBOUND)
                            {
                                LServBound.Cells["A" + num + ":AA" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                LServBound.Cells["A" + num + ":AA" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                                LServBound.Cells["A" + num].Value = c.LOADSERVBOUND_ID;
                                LServBound.Cells["B" + num].Value = c.G3E_FID;
                                LServBound.Cells["C" + num].Value = c.SCHEME_NAME;
                                LServBound.Cells["D" + num].Value = c.USERID;
                                LServBound.Cells["E" + num].Value = c.G3E_ID;
                                LServBound.Cells["F" + num].Value = c.SERVICEBOUNDARYID;
                                LServBound.Cells["G" + num].Value = c.SITENO;
                                LServBound.Cells["H" + num].Value = c.SITEFLOOR;
                                LServBound.Cells["I" + num].Value = c.SITEBUILDING;
                                LServBound.Cells["J" + num].Value = c.STREETYPE;
                                LServBound.Cells["K" + num].Value = c.STREETNAME;
                                LServBound.Cells["L" + num].Value = c.SECTION;
                                LServBound.Cells["M" + num].Value = c.POSTCODE;
                                LServBound.Cells["N" + num].Value = c.CITY;
                                LServBound.Cells["O" + num].Value = c.STATE;
                                LServBound.Cells["P" + num].Value = c.COUNTRY;
                                LServBound.Cells["Q" + num].Value = c.PREMISETYPE;
                                LServBound.Cells["R" + num].Value = c.SERVCATEGORY;
                                LServBound.Cells["S" + num].Value = c.SITESDP;
                                LServBound.Cells["T" + num].Value = c.SITEDPEXC;
                                LServBound.Cells["U" + num].Value = c.DDP;
                                LServBound.Cells["V" + num].Value = c.SMARTDEVELOPER;
                                LServBound.Cells["W" + num].Value = c.SMARTPROJECTID;
                                LServBound.Cells["X" + num].Value = c.GRN_STATUS;
                                LServBound.Cells["Y" + num].Value = c.COS;
                                LServBound.Cells["Z" + num].Value = c.DROPCABLEDISTANCE;
                                LServBound.Cells["AA" + num++].Value = c.FLAG;
                            }

                        }
                        #endregion

                        #region loadpath
                        if (queryLOADPATH.Count() > 1)
                        {
                            var LPath = xlsFile.Workbook.Worksheets.Add("Load Path Consumer");
                            LPath.Cells.Style.Font.Name = "Arial";
                            LPath.Cells.Style.Font.Size = 8;

                            LPath.Column(16).Width = 15.00; //N
                            LPath.Column(17).Width = 15.00; //O
                            LPath.Column(19).Width = 15.00; //Q
                            for (int i = 20; i < 26; i++)//R S T U V
                                LPath.Column(i).Width = 10.00;

                            LPath.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            LPath.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                            LPath.Cells["A1:Z1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                            LPath.Cells["A1:Z1"].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                            LPath.Cells["A1:Z1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                            LPath.Cells["A1:Z1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                            LPath.Cells["A1:Z1"].Style.Font.Bold = true;


                            LPath.Cells["A1"].Value = "ID";
                            LPath.Cells["B1"].Value = "ANAME";
                            LPath.Cells["C1"].Value = "ATYPE";
                            LPath.Cells["D1"].Value = "ASITE";
                            LPath.Cells["E1"].Value = "ACARD2";
                            LPath.Cells["F1"].Value = "APORT2";
                            LPath.Cells["G1"].Value = "ACARD3";
                            LPath.Cells["H1"].Value = "APORT3";
                            LPath.Cells["I1"].Value = "ZNAME";
                            LPath.Cells["J1"].Value = "ZTYPE";
                            LPath.Cells["K1"].Value = "ZSITE";
                            LPath.Cells["L1"].Value = "ZCARD";
                            LPath.Cells["M1"].Value = "ZPORT";
                            LPath.Cells["N1"].Value = "DPNAME";
                            LPath.Cells["O1"].Value = "DPSITE";
                            LPath.Cells["P1"].Value = "DPPORT";
                            LPath.Cells["Q1"].Value = "JOBID";
                            LPath.Cells["R1"].Value = "PROCESSID";
                            LPath.Cells["S1"].Value = "PROCESSTIME";
                            LPath.Cells["T1"].Value = "OUTPORTFDC";
                            LPath.Cells["U1"].Value = "INPORTFDP";
                            LPath.Cells["V1"].Value = "REMARKS";
                            LPath.Cells["W1"].Value = "APORT4";
                            LPath.Cells["X1"].Value = "ACARD4";
                            LPath.Cells["Y1"].Value = "DP_TAMBAHAN";
                            LPath.Cells["Z1"].Value = "FLAG";

                            int num = 2;
                            foreach (var c in queryLOADPATH)
                            {
                                LPath.Cells["A" + num + ":Z" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                LPath.Cells["A" + num + ":Z" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                                LPath.Cells["A" + num].Value = c.ID;
                                LPath.Cells["B" + num].Value = c.ANAME;
                                LPath.Cells["C" + num].Value = c.ATYPE;
                                LPath.Cells["D" + num].Value = c.ASITE;
                                LPath.Cells["E" + num].Value = c.ACARD2;
                                LPath.Cells["F" + num].Value = c.APORT2;
                                LPath.Cells["G" + num].Value = c.ACARD3;
                                LPath.Cells["H" + num].Value = c.APORT3;
                                LPath.Cells["I" + num].Value = c.ZNAME;
                                LPath.Cells["J" + num].Value = c.ZTYPE;
                                LPath.Cells["K" + num].Value = c.ZSITE;
                                LPath.Cells["L" + num].Value = c.ZCARD;
                                LPath.Cells["M" + num].Value = c.ZPORT;
                                LPath.Cells["N" + num].Value = c.DPNAME;
                                LPath.Cells["O" + num].Value = c.DPSITE;
                                LPath.Cells["P" + num].Value = c.DPPORT;
                                LPath.Cells["Q" + num].Value = c.JOBID;
                                LPath.Cells["R" + num].Value = c.PROCESSID;
                                LPath.Cells["S" + num].Value = c.PROCESSTIME;
                                LPath.Cells["T" + num].Value = c.OUTPORTFDC;
                                LPath.Cells["U" + num].Value = c.INPORTFDP;
                                LPath.Cells["V" + num].Value = c.REMARKS;
                                LPath.Cells["W" + num].Value = c.APORT4;
                                LPath.Cells["X" + num].Value = c.ACARD4;
                                LPath.Cells["Y" + num].Value = c.DP_TAMBAHAN;
                                LPath.Cells["Z" + num++].Value = c.FLAG;
                            }

                        }
                        #endregion

                        #region addcard
                        if (queryADDCARD.Count() > 0)
                        {
                            var ACard = xlsFile.Workbook.Worksheets.Add("Add Card");
                            ACard.Cells.Style.Font.Name = "Arial";
                            ACard.Cells.Style.Font.Size = 8;

                            ACard.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ACard.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                            ACard.Cells["A1:S1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                            ACard.Cells["A1:S1"].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                            ACard.Cells["A1:S1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                            ACard.Cells["A1:S1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                            ACard.Cells["A1:S1"].Style.Font.Bold = true;

                            ACard.Column(1).Width = 9.00; //A
                            ACard.Column(2).Width = 11.00; //B
                            ACard.Column(4).Width = 20.00; //D
                            ACard.Column(6).Width = 20.00; //F
                            ACard.Column(8).Width = 10.00; //H
                            ACard.Column(9).Width = 12.00; //I
                            ACard.Column(13).Width = 12.00; //M
                            ACard.Column(17).Width = 13.00; //Q

                            ACard.Cells["A1"].Value = "ADDCARD_ID";
                            ACard.Cells["B1"].Value = "CLASS_NAME";
                            ACard.Cells["C1"].Value = "G3E_FID";
                            ACard.Cells["D1"].Value = "SCHEME_NAME";
                            ACard.Cells["E1"].Value = "USERID";
                            ACard.Cells["F1"].Value = "EQUIPMENTID";
                            ACard.Cells["G1"].Value = "EQUIPVEND";
                            ACard.Cells["H1"].Value = "EQUIPMODEL";
                            ACard.Cells["I1"].Value = "TEMPLATENAME";
                            ACard.Cells["J1"].Value = "SLOT_NO";
                            ACard.Cells["K1"].Value = "CARD_TYPE";
                            ACard.Cells["L1"].Value = "TOTAL_PORT";
                            ACard.Cells["M1"].Value = "PROJECT_NAME";
                            ACard.Cells["N1"].Value = "PATCHCORD";
                            ACard.Cells["O1"].Value = "DESIGN";
                            ACard.Cells["P1"].Value = "ADDCARD";
                            ACard.Cells["Q1"].Value = "ADDSPLITTER";
                            ACard.Cells["R1"].Value = "OWNER_FID";
                            ACard.Cells["S1"].Value = "FLAG";

                            int num = 2;
                            foreach (var c in queryADDCARD)
                            {
                                ACard.Cells["A" + num + ":S" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                ACard.Cells["A" + num + ":S" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                                ACard.Cells["A" + num].Value = c.ADDCARD_ID;
                                ACard.Cells["B" + num].Value = c.CLASS_NAME;
                                ACard.Cells["C" + num].Value = c.G3E_FID;
                                ACard.Cells["D" + num].Value = c.SCHEME_NAME;
                                ACard.Cells["E" + num].Value = c.USERID;
                                ACard.Cells["F" + num].Value = c.EQUIPMENTID;
                                ACard.Cells["G" + num].Value = c.EQUIPVEND;
                                ACard.Cells["H" + num].Value = c.EQUIPMODEL;
                                ACard.Cells["I" + num].Value = c.TEMPLATENAME;
                                ACard.Cells["J" + num].Value = c.SLOT_NO;
                                ACard.Cells["K" + num].Value = c.CARD_TYPE;
                                ACard.Cells["L" + num].Value = c.TOTAL_PORT;
                                ACard.Cells["M" + num].Value = c.PROJECT_NAME;
                                ACard.Cells["N" + num].Value = c.PATCHCORD;
                                ACard.Cells["O" + num].Value = c.DESIGN;
                                ACard.Cells["P" + num].Value = c.ADDCARD;
                                ACard.Cells["Q" + num].Value = c.ADDSPLITTER;
                                ACard.Cells["R" + num].Value = c.OWNER_FID;
                                ACard.Cells["S" + num++].Value = c.FLAG;
                            }
                        }
                        #endregion



                        var stream = new MemoryStream();
                        xlsFile.SaveAs(stream);

                        string fileName = "CheckJobGRN_" + targetJob + ".xlsx";
                        string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                        stream.Position = 0;
                        return File(stream, contentType, fileName);
                    }
                }
            }
        }
        #endregion

        //public ActionResult PATHCOMSUMER(string id)
        //{
        //    //string PATHCOMSUMER = "";
        //    List<string> PATHCOMSUMER = new List<string>();
        //    using (Entities ctxNeps = new Entities())
        //    {
        //        var queryPATHCONSUMER = from a in ctxNeps.WV_LOAD_PATH_CONSUMER
        //                                where id.Trim().ToUpper() == a.JOBID.Trim().ToUpper()
        //                                select new
        //                                {
        //                                    a.ANAME,
        //                                    a.ATYPE,
        //                                    a.ACARD2,
        //                                    a.APORT2,
        //                                    a.ACARD3,
        //                                    a.APORT3,
        //                                    a.ZNAME,
        //                                    a.ZTYPE,
        //                                    a.ZCARD,
        //                                    a.ZPORT,
        //                                };

        //        if (queryPATHCONSUMER.Count() > 0)
        //        {
        //            foreach (var a in queryPATHCONSUMER)
        //            {
        //                PATHCOMSUMER.Add("[" + a.ANAME + "|" + a.ATYPE + "|" + a.ACARD2 + "|" + a.APORT2 + "|" + a.ACARD3 + "|" + a.APORT3 + "|" + a.ZNAME + "|" + a.ZTYPE + "|" + a.ZCARD + "|" + a.ZPORT);
        //            }
        //        }

        //        System.Diagnostics.Debug.WriteLine("a = " + PATHCOMSUMER);

        //        return Json(new
        //        {
        //            pConsumer = PATHCOMSUMER
        //        }, JsonRequestBehavior.AllowGet);

        //    }
        //}
    }
}
