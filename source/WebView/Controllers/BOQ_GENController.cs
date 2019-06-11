using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.OleDb;
using System.Linq;
using System.Web.Mvc;
using PagedList;
using Oracle.DataAccess.Client;
using NEPS.BOQ.Utilities;
using NEPS.BOQ.Classes.ContractBOQ;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;
using System.IO;

namespace WebView.Controllers
{

    public class BOQ_GENController : Controller
    {
        private string connString = ConfigurationManager.AppSettings.Get("connString");

        // GET: /BOQ_GEN/
        public ActionResult BOQ_GEN(int? page, string searchName, string searchIDName, string searchOwner, string searchIDOwner, string searchNo, string searchIDNo, string searchType, string searchIDType, string searchExc, string searchIDExc, string searchYear, string searchIDYear)
        {


            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            string user = User.Identity.Name;
            //WebService._base.OSPBOQ_MAIN_EXCEL BOQ_MAIN = new WebService._base.OSPBOQ_MAIN_EXCEL();
            //BOQ_MAIN = myWebService.GetOSPBOQ_MAIN_Excel(0, 100, "");

            WebService._base.OSPBOQ_GEN BOQ_GEN = new WebService._base.OSPBOQ_GEN();

            if (searchNo != null  && searchType != null && searchExc != null && searchYear != null && searchName != null && searchOwner != null)
            {
                ViewBag.searchName = searchName;
                ViewBag.searchNo = searchNo;
                ViewBag.searchType = searchType;
                ViewBag.searchExc = searchExc;
                ViewBag.searchYear = searchYear;
                ViewBag.searchOwner = searchOwner;

                if (searchNo.Equals("Select") && searchType.Equals("Select") && searchExc.Equals("Select") && searchYear.Equals("Select")  && searchName.Equals("Select") && searchOwner.Equals("Select"))
                    BOQ_GEN = myWebService.GetOSPBOQ_GEN(user,"JKH", 0, 100, null, null, null, null, null, null, null, null, null, null, null, null);
                else
                {
                    BOQ_GEN = myWebService.GetOSPBOQ_GEN(user,"JKH", 0, 100, searchName, searchIDName, searchOwner, searchIDOwner, searchNo, searchIDNo, searchType, searchIDType, searchExc, searchIDExc, searchYear, searchIDYear);
                }
            }

            else
            {
                BOQ_GEN = myWebService.GetOSPBOQ_GEN(user,"JKH", 0, 100,  null, null, null, null, null, null, null, null, null, null, null, null);
                
                ViewBag.searchName = "Select";
                ViewBag.searchNo = "Select";
                ViewBag.searchType = "Select";
                ViewBag.searchExc = "Select";
                ViewBag.searchYear = "Select";
                ViewBag.searchOwner = "Select";
            }


            //using (Entities ctxData = new Entities())
            //{
            //    List<SelectListItem> list = new List<SelectListItem>();
            //    var query = from p in ctxData.WV_PU_MAST
            //                where p.NETWORK_FLAG == "U"
            //                select new { Text = p.PU_ID, Value = p.PU_ID };

            //    list.Add(new SelectListItem() { Text = "", Value = "Select" });
            //    foreach (var a in query.Distinct().OrderBy(it => it.Text))
            //    {
            //        if (a.Value != null)
            //            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
            //    }
            //    ViewBag.x_PU_ID = list;
            //}

            ViewData["data2"] = BOQ_GEN.BOQ_GENList;
            //System.Diagnostics.Debug.WriteLine("SEARCHID : " + searchID);

           


            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.G3E_JOB
                            where p.G3E_OWNER == user
                            orderby p.G3E_OWNER ascending
                            select new { Text = p.G3E_OWNER, Value = p.G3E_OWNER };

                List<SelectListItem> list = new List<SelectListItem>();
                if (searchOwner == "Select" || searchOwner == null)
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }
                else
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    list.Add(new SelectListItem() { Text = searchOwner.ToString(), Value = searchOwner.ToString(), Selected = true });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null && a.Value != searchOwner.ToString())
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }

                ViewBag.x_G3E_OWNER = list;
            }

            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.G3E_JOB
                            where p.SCHEME_NAME != null && (p.JOB_TYPE == "Civil" || p.JOB_TYPE == "E/Side" || p.JOB_TYPE == "D/Side" || p.JOB_TYPE == "Fiber E/Side" || p.JOB_TYPE == "HSBB E/Side" || p.JOB_TYPE == "HSBB D/Side" || p.JOB_TYPE == "Fiber Trunk" || p.JOB_TYPE == "Fiber Junction" || p.JOB_TYPE == "Others") && p.G3E_OWNER == user
                            orderby p.SCHEME_NAME ascending
                            select new { Text = p.SCHEME_NAME, Value = p.SCHEME_NAME };

                List<SelectListItem> list = new List<SelectListItem>();
                if (searchName == "Select" || searchName == null)
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }
                else
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    list.Add(new SelectListItem() { Text = searchName.ToString(), Value = searchName.ToString(), Selected = true });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null && a.Value != searchName.ToString())
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }

                ViewBag.x_SCHEME_NAME = list;
            }

            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.G3E_JOB
                            where (p.JOB_TYPE == "Civil" || p.JOB_TYPE == "E/Side" || p.JOB_TYPE == "D/Side" || p.JOB_TYPE == "Fiber E/Side" || p.JOB_TYPE == "HSBB E/Side" || p.JOB_TYPE == "HSBB D/Side" || p.JOB_TYPE == "Fiber Trunk" || p.JOB_TYPE == "Fiber Junction" || p.JOB_TYPE == "Others") && p.G3E_OWNER == user
                            orderby p.SCH_NO ascending
                            select new { Text = p.SCH_NO, Value = p.SCH_NO };

                List<SelectListItem> list = new List<SelectListItem>();
                if (searchNo == "Select" || searchNo == null)
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }
                else
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    list.Add(new SelectListItem() { Text = searchNo.ToString(), Value = searchNo.ToString(), Selected = true });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null && a.Value != searchNo.ToString())
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }

                ViewBag.x_SCH_NO = list;
            }

            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.G3E_JOB
                            where (p.JOB_TYPE == "Civil" || p.JOB_TYPE == "E/Side" || p.JOB_TYPE == "D/Side" || p.JOB_TYPE == "Fiber E/Side" || p.JOB_TYPE == "HSBB E/Side" || p.JOB_TYPE == "HSBB D/Side" || p.JOB_TYPE == "Fiber Trunk" || p.JOB_TYPE == "Fiber Junction" || p.JOB_TYPE == "Others") && p.G3E_OWNER == user
                            orderby p.JOB_TYPE ascending
                            select new { Text = p.JOB_TYPE, Value = p.JOB_TYPE };

                List<SelectListItem> list = new List<SelectListItem>();
                if (searchType == "Select" || searchType == null)
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }
                else
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    list.Add(new SelectListItem() { Text = searchType.ToString(), Value = searchType.ToString(), Selected = true });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null && a.Value != searchType.ToString())
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }

                ViewBag.x_SCH_TYPE = list;
            }


            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.G3E_JOB
                            where (p.JOB_TYPE == "Civil" || p.JOB_TYPE == "E/Side" || p.JOB_TYPE == "D/Side" || p.JOB_TYPE == "Fiber E/Side" || p.JOB_TYPE == "HSBB E/Side" || p.JOB_TYPE == "HSBB D/Side" || p.JOB_TYPE == "Fiber Trunk" || p.JOB_TYPE == "Fiber Junction" || p.JOB_TYPE == "Others") && p.G3E_OWNER == user
                            orderby p.EXC_ABB ascending
                            select new { Text = p.EXC_ABB, Value = p.EXC_ABB };

                List<SelectListItem> list = new List<SelectListItem>();
                if (searchExc == "Select" || searchExc == null)
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }
                else
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    list.Add(new SelectListItem() { Text = searchExc.ToString(), Value = searchExc.ToString(), Selected = true });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null && a.Value != searchExc.ToString())
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }

                ViewBag.x_EXC_ABB = list;
            }


            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.G3E_JOB
                            where (p.JOB_TYPE == "Civil" || p.JOB_TYPE == "E/Side" || p.JOB_TYPE == "D/Side" || p.JOB_TYPE == "Fiber E/Side" || p.JOB_TYPE == "HSBB E/Side" || p.JOB_TYPE == "HSBB D/Side" || p.JOB_TYPE == "Fiber Trunk" || p.JOB_TYPE == "Fiber Junction" || p.JOB_TYPE == "Others") && p.G3E_OWNER == user
                            orderby p.YEAR_INSTALL ascending
                            select new { Text = p.YEAR_INSTALL, Value = p.YEAR_INSTALL };

                List<SelectListItem> list = new List<SelectListItem>();
                if (searchYear == "Select" || searchYear == null)
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }
                else
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    list.Add(new SelectListItem() { Text = searchYear.ToString(), Value = searchYear.ToString(), Selected = true });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null && a.Value.ToString() != searchYear.ToString())
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }

                ViewBag.x_YEAR_INSTALL = list;
            }

            //return View();
            int pageSize = 5;
            int pageNumber = (page ?? 1);
            //BOQ_MAIN.BOQ_MAIN_Excel.ToPagedList(pageNumber, pageSize);
            return View(BOQ_GEN.BOQ_GENList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult BOQ_GEN_CONTRACT(int? page,  string searchName, string searchIDName, string searchOwner, string searchIDOwner, string searchNo, string searchIDNo, string searchType, string searchIDType, string searchExc, string searchIDExc, string searchYear, string searchIDYear)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            string user = User.Identity.Name;
            //WebService._base.OSPBOQ_MAIN_EXCEL BOQ_MAIN = new WebService._base.OSPBOQ_MAIN_EXCEL();
            //BOQ_MAIN = myWebService.GetOSPBOQ_MAIN_Excel(0, 100, "");

            WebService._base.OSPBOQ_GEN BOQ_GEN = new WebService._base.OSPBOQ_GEN();
            /*string searchIDID = "";
            string searchID = "";*/
            if (searchNo != null && searchType != null && searchExc != null && searchYear != null &&  searchName != null && searchOwner != null)
            {
                ViewBag.searchName = searchName;
                ViewBag.searchNo = searchNo;
                ViewBag.searchType = searchType;
                ViewBag.searchExc = searchExc;
                ViewBag.searchYear = searchYear;
                ViewBag.searchOwner = searchOwner;

                if (searchNo.Equals("Select") && searchType.Equals("Select") && searchExc.Equals("Select") && searchYear.Equals("Select") && searchName.Equals("Select") && searchOwner.Equals("Select"))
                    BOQ_GEN = myWebService.GetOSPBOQ_GEN(user,"CONTRACT", 0, 100,  null, null, null, null, null, null, null, null, null, null, null, null);
                else
                {
                    BOQ_GEN = myWebService.GetOSPBOQ_GEN(user,"CONTRACT", 0, 100, searchName, searchIDName, searchOwner, searchIDOwner, searchNo, searchIDNo, searchType, searchIDType, searchExc, searchIDExc, searchYear, searchIDYear);                   
                }
            }
            else
            {
                BOQ_GEN = myWebService.GetOSPBOQ_GEN(user,"CONTRACT", 0, 100, null, null, null, null, null, null, null, null, null, null, null, null);
                
                ViewBag.searchName = "Select";
                ViewBag.searchNo = "Select";
                ViewBag.searchType = "Select";
                ViewBag.searchExc = "Select";
                ViewBag.searchYear = "Select";
                ViewBag.searchOwner = "Select";
            }


            using (Entities ctxData = new Entities())
            {
                List<SelectListItem> list = new List<SelectListItem>();
                var query = from p in ctxData.WV_PU_MAST
                            where p.NETWORK_FLAG == "U"
                            select new { Text = p.PU_ID, Value = p.PU_ID + "/" + p.BILL_RATE };

                list.Add(new SelectListItem() { Text = "", Value = "Select" });
                //foreach (var a in query.Distinct().OrderBy(it => it.Value))
                //{
                //    if (a.Value != null)
                //        list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                //}
                ViewBag.x_PU_ID = list;
            }

            ViewData["data2"] = BOQ_GEN.BOQ_GENList;
           

            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.G3E_JOB
                            where p.G3E_OWNER == user
                            orderby p.G3E_OWNER ascending
                            select new { Text = p.G3E_OWNER, Value = p.G3E_OWNER };

                var query2 = from p in ctxData.WV_ISP_JOB
                            where  p.G3E_OWNER == user
                            orderby p.G3E_OWNER ascending
                            select new { Text = p.G3E_OWNER, Value = p.G3E_OWNER };

                List<SelectListItem> list = new List<SelectListItem>();
                int check = 0;
                if (searchOwner == "Select" || searchOwner == null)
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                        check = 1;

                    }
                    if (check == 0)
                    {
                        foreach (var a in query2.Distinct().OrderBy(it => it.Value))
                        {
                            if (a.Value != null)
                                list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                        }
                    }
                }
                else
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    list.Add(new SelectListItem() { Text = searchOwner.ToString(), Value = searchOwner.ToString(), Selected = true });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null && a.Value != searchOwner.ToString())
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                        check = 1;
                    }
                    if (check == 0)
                    {
                        foreach (var a in query2.Distinct().OrderBy(it => it.Value))
                        {
                            if (a.Value != null && a.Value != searchOwner.ToString())
                                list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                        }
                    }
                }

                ViewBag.x_G3E_OWNER = list;
            }

            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.G3E_JOB
                            where (p.JOB_TYPE == "Copper (Equip)" || p.JOB_TYPE == "Fiber (Equip)" || p.JOB_TYPE == "HSBB (Equip)" || p.JOB_TYPE == "ISP") && p.G3E_OWNER == user
                            orderby p.SCHEME_NAME ascending
                            select new { Text = p.SCHEME_NAME, Value = p.SCHEME_NAME };

                var query2 = from p in ctxData.WV_ISP_JOB
                             where (p.JOB_TYPE == "ISP" || p.JOB_TYPE == "METROE" || p.JOB_TYPE == "IPCORE" || p.JOB_TYPE == "NGND" || p.JOB_TYPE == "NGND(AG)" || p.JOB_TYPE == "NGND(MG)"||
                             p.JOB_TYPE == "CID (BRAS)" || p.JOB_TYPE == "CID (Global)" || p.JOB_TYPE == "CID (IPCORE)" || p.JOB_TYPE == "CID (MetroE)") && p.G3E_OWNER == user
                            orderby p.SCHEME_NAME ascending
                            select new { Text = p.SCHEME_NAME, Value = p.SCHEME_NAME };


                List<SelectListItem> list = new List<SelectListItem>();
                if (searchName == "Select" || searchName == null)
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                    foreach (var a in query2.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }
                else
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    list.Add(new SelectListItem() { Text = searchName.ToString(), Value = searchName.ToString(), Selected = true });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null && a.Value != searchName.ToString())
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                    foreach (var a in query2.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null && a.Value != searchName.ToString())
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }

                ViewBag.x_SCHEME_NAME = list;
            }

            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.G3E_JOB
                            where (p.JOB_TYPE == "Copper (Equip)" || p.JOB_TYPE == "Fiber (Equip)" || p.JOB_TYPE == "HSBB (Equip)" || p.JOB_TYPE == "ISP") && p.G3E_OWNER == user
                            orderby p.SCH_NO ascending
                            select new { Text = p.SCH_NO, Value = p.SCH_NO };

                var query2 = from p in ctxData.WV_ISP_JOB
                             where (p.JOB_TYPE == "ISP" || p.JOB_TYPE == "METROE" || p.JOB_TYPE == "IPCORE" || p.JOB_TYPE == "NGND" || p.JOB_TYPE == "NGND(AG)" || p.JOB_TYPE == "NGND(MG)" ||
                             p.JOB_TYPE == "CID (BRAS)" || p.JOB_TYPE == "CID (Global)" || p.JOB_TYPE == "CID (IPCORE)" || p.JOB_TYPE == "CID (MetroE)") && p.G3E_OWNER == user
                            orderby p.SCH_NO ascending
                            select new { Text = p.SCH_NO, Value = p.SCH_NO };

                List<SelectListItem> list = new List<SelectListItem>();
                if (searchNo == "Select" || searchNo == null)
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                    foreach (var a in query2.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }
                else
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    list.Add(new SelectListItem() { Text = searchNo.ToString(), Value = searchNo.ToString(), Selected = true });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null && a.Value != searchNo.ToString())
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                    foreach (var a in query2.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null && a.Value != searchNo.ToString())
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }

                ViewBag.x_SCH_NO = list;
            }

            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.G3E_JOB
                            where (p.JOB_TYPE == "Copper (Equip)" || p.JOB_TYPE == "Fiber (Equip)" || p.JOB_TYPE == "HSBB (Equip)" || p.JOB_TYPE == "ISP") && p.G3E_OWNER == user
                            orderby p.JOB_TYPE ascending
                            select new { Text = p.JOB_TYPE, Value = p.JOB_TYPE };

                var query2 = from p in ctxData.WV_ISP_JOB
                             where (p.JOB_TYPE == "ISP" || p.JOB_TYPE == "METROE" || p.JOB_TYPE == "IPCORE" || p.JOB_TYPE == "NGND" || p.JOB_TYPE == "NGND(AG)" || p.JOB_TYPE == "NGND(MG)" ||
                             p.JOB_TYPE == "CID (BRAS)" || p.JOB_TYPE == "CID (Global)" || p.JOB_TYPE == "CID (IPCORE)" || p.JOB_TYPE == "CID (MetroE)") && p.G3E_OWNER == user 
                             orderby p.JOB_TYPE ascending
                            select new { Text = p.JOB_TYPE, Value = p.JOB_TYPE };

                List<SelectListItem> list = new List<SelectListItem>();
                if (searchType == "Select" || searchType == null)
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                    foreach (var a in query2.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }
                else
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    list.Add(new SelectListItem() { Text = searchType.ToString(), Value = searchType.ToString(), Selected = true });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null && a.Value != searchType.ToString())
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                    foreach (var a in query2.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null && a.Value != searchType.ToString())
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }

                ViewBag.x_SCH_TYPE = list;
            }


            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.G3E_JOB
                            where (p.JOB_TYPE == "Copper (Equip)" || p.JOB_TYPE == "Fiber (Equip)" || p.JOB_TYPE == "HSBB (Equip)" || p.JOB_TYPE == "ISP") && p.G3E_OWNER == user
                            orderby p.EXC_ABB ascending
                            select new { Text = p.EXC_ABB, Value = p.EXC_ABB };
                var query2 = from p in ctxData.WV_ISP_JOB
                             where (p.JOB_TYPE == "ISP" || p.JOB_TYPE == "METROE" || p.JOB_TYPE == "IPCORE" || p.JOB_TYPE == "NGND" || p.JOB_TYPE == "NGND(AG)" || p.JOB_TYPE == "NGND(MG)" ||
                             p.JOB_TYPE == "CID (BRAS)" || p.JOB_TYPE == "CID (Global)" || p.JOB_TYPE == "CID (IPCORE)" || p.JOB_TYPE == "CID (MetroE)") && p.G3E_OWNER == user
                            orderby p.EXC_ABB ascending
                            select new { Text = p.EXC_ABB, Value = p.EXC_ABB };

                List<SelectListItem> list = new List<SelectListItem>();
                if (searchExc == "Select" || searchExc == null)
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                    foreach (var a in query2.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }
                else
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    list.Add(new SelectListItem() { Text = searchExc.ToString(), Value = searchExc.ToString(), Selected = true });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null && a.Value != searchExc.ToString())
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                    foreach (var a in query2.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null && a.Value != searchExc.ToString())
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }

                ViewBag.x_EXC_ABB = list;
            }


            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.G3E_JOB
                            where (p.JOB_TYPE == "Copper (Equip)" || p.JOB_TYPE == "Fiber (Equip)" || p.JOB_TYPE == "HSBB (Equip)" || p.JOB_TYPE == "ISP") && p.G3E_OWNER == user
                            orderby p.YEAR_INSTALL ascending
                            select new { Text = p.YEAR_INSTALL, Value = p.YEAR_INSTALL };

                var query2 = from p in ctxData.WV_ISP_JOB
                             where (p.JOB_TYPE == "ISP" || p.JOB_TYPE == "METROE" || p.JOB_TYPE == "IPCORE" || p.JOB_TYPE == "NGND" || p.JOB_TYPE == "NGND(AG)" || p.JOB_TYPE == "NGND(MG)" ||
                             p.JOB_TYPE == "CID (BRAS)" || p.JOB_TYPE == "CID (Global)" || p.JOB_TYPE == "CID (IPCORE)" || p.JOB_TYPE == "CID (MetroE)") && p.G3E_OWNER == user
                            orderby p.YEAR_INSTALL ascending
                            select new { Text = p.YEAR_INSTALL, Value = p.YEAR_INSTALL };

                List<SelectListItem> list = new List<SelectListItem>();
                if (searchYear == "Select" || searchYear == null)
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                    foreach (var a in query2.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null)
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }
                else
                {
                    list.Add(new SelectListItem() { Text = "", Value = "Select" });
                    list.Add(new SelectListItem() { Text = searchYear.ToString(), Value = searchYear.ToString(), Selected = true });
                    foreach (var a in query.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null && a.Value.ToString() != searchYear.ToString())
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                    foreach (var a in query2.Distinct().OrderBy(it => it.Value))
                    {
                        if (a.Value != null && a.Value.ToString() != searchYear.ToString())
                            list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });

                    }
                }

                ViewBag.x_YEAR_INSTALL = list;
            }

            //return View();
            int pageSize = 5;
            int pageNumber = (page ?? 1);
            //BOQ_MAIN.BOQ_MAIN_Excel.ToPagedList(pageNumber, pageSize);
            return View(BOQ_GEN.BOQ_GENList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult UpdateData(string txtJOBID, string txtSCHEMENO, string txtSCHEMENAME, string txtSCHEMETYPE, string txtEXCABB, string txtYEARINSTALL, string txtOWNER)
        {
            bool success = true;
            bool selected = false;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.BOQ_GEN newBOQ_GEN = new WebService._base.BOQ_GEN();
            newBOQ_GEN.x_G3E_ID = txtJOBID;
            newBOQ_GEN.x_Sch_No = txtSCHEMENO;
            newBOQ_GEN.x_Scheme_Name = txtSCHEMENAME;
            newBOQ_GEN.x_Sch_Type = txtSCHEMETYPE;
            newBOQ_GEN.x_Exc_Abb = txtEXCABB;
            newBOQ_GEN.x_Year_Install = txtYEARINSTALL;
            newBOQ_GEN.x_G3E_OWNER = txtOWNER;

            success = myWebService.AddBOQ_GEN(newBOQ_GEN);
            //System.Diagnostics.Debug.WriteLine("update:" + success);

            selected = true;

            //if (ModelState.IsValid && selected)
            //{
            //    if (success == true)
            //        return RedirectToAction("BOQ_MAIN/BOQ_MAIN_List");
            //    else
            //        return RedirectToAction("NewSaveFail"); // store to db failed.
            //}

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

         [HttpPost]
        public ActionResult BOQPuList(string schemeName)
        {
            string PuList = "";
            using (Entities ctxData = new Entities())
            {
               /* var query = from p in ctxData.WV_PU_MAST
                            where p.NETWORK_FLAG == "U" || p.NETWORK_FLAG == "G"
                            select new { p.PU_ID, p.PU_DESC, p.NETWORK_FLAG };*/

                var query = from p in ctxData.WV_PU_MAST
                            where p.NETWORK_FLAG == "U" 
                            select new { p.PU_ID, p.PU_DESC, p.NETWORK_FLAG };

                var queryPu = from p in ctxData.WV_BOQ_DATA
                              where p.SCHEME_NAME == schemeName
                              select new { p.PU_ID };

                foreach (var a in query.Distinct().OrderBy(it => it.PU_ID))
                {
                    int checkPu = 0;
                    foreach (var b in queryPu)
                    {
                        if (b.PU_ID.Count() == 4)
                            checkPu = 1;
                    }
                    if (checkPu == 0)
                        PuList = PuList + a.PU_ID + ": (" + a.NETWORK_FLAG + ") " + a.PU_DESC + "|";
                }
            }
            return Json(new
            {
                txtSchName = schemeName,
                PuList = PuList
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
         public ActionResult BOQList(string schemeName)
        {
            System.Diagnostics.Debug.WriteLine("BOQ LIST : 1");
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            

            string check = "";
            string Proj_no_isp = "";
            string Proj_no_osp = "";
            using (Entities ctxData = new Entities())
            {
                var query = from q in ctxData.G3E_JOB
                            where q.SCHEME_NAME == schemeName
                            select new { q.JOB_TYPE, q.PROJECT_NO };

                System.Diagnostics.Debug.WriteLine("BOQ LIST : 2");
                foreach (var a in query)
                {
                    if (a.JOB_TYPE == "Civil" || a.JOB_TYPE == "E/Side" || a.JOB_TYPE == "D/Side" || a.JOB_TYPE == "Fiber E/Side" || a.JOB_TYPE == "HSBB E/Side" || a.JOB_TYPE == "HSBB D/Side" || a.JOB_TYPE == "Fiber Trunk" || a.JOB_TYPE == "Fiber Junction" || a.JOB_TYPE == "Others")
                    {
                        check = "JKH";
                        Proj_no_isp = a.PROJECT_NO;
                    }
                    else
                    {
                        check = "CONTRACT";
                    }
                    //Proj_no = a.PROJECT_NO;
                }
                System.Diagnostics.Debug.WriteLine("BOQ LIST : 3");
                var queryISP = from q in ctxData.WV_ISP_JOB
                               where q.SCHEME_NAME == schemeName
                               select new { q.JOB_TYPE, q.PROJECT_NO };

                foreach (var a in queryISP)
                {
                    if (a.JOB_TYPE == "Civil" || a.JOB_TYPE == "E/Side" || a.JOB_TYPE == "D/Side" || a.JOB_TYPE == "Fiber E/Side" || a.JOB_TYPE == "HSBB E/Side" || a.JOB_TYPE == "HSBB D/Side" || a.JOB_TYPE == "Fiber Trunk" || a.JOB_TYPE == "Fiber Junction" || a.JOB_TYPE == "Others")
                    {
                        check = "JKH";
                    }
                    else
                    {
                        check = "CONTRACT";
                        Proj_no_osp = a.PROJECT_NO;
                    }
                }
            }
            System.Diagnostics.Debug.WriteLine("BOQ LIST : 4");
            //string res = "bodo";
            string res =  myWebService.GetBOQ(schemeName, check);
            string mat = "no";
            if (check == "JKH")
            {
                mat = myWebService.GetMaterial(schemeName);
            }
            System.Diagnostics.Debug.WriteLine("BOQ LIST : 5");
            string dataError = myWebService.GetBOQError(schemeName, check);
            System.Diagnostics.Debug.WriteLine("BOQ LIST : 6");
            string gems = myWebService.CheckGEMS(schemeName);
            //string MainList = myWebService.GetEstimate_MainList(schemeName);
            System.Diagnostics.Debug.WriteLine("BOQ LIST : 7");
            WebService._base.SUMM_ProjCost summ_projCost = new WebService._base.SUMM_ProjCost();
            summ_projCost = myWebService.Get_ProjCost(schemeName);
            string SumLbrOT = "";
            string SumLbrSalary = "";
            string SumMaterial = "";
            string SumJKH = "";
            string SumTNT_Mileage = "";
            string SumMilling = "";
            string SumMisc = "";
            string SumContract = "";
            for (int i = 0; i < summ_projCost.ProjCost.Count; i++)
            {
                SumLbrOT = summ_projCost.ProjCost[i].SUM_LBR_OT;
                SumLbrSalary = summ_projCost.ProjCost[i].SUM_LBR_SALARY;
                SumMaterial = summ_projCost.ProjCost[i].SUM_MATERIAL;
                SumJKH = summ_projCost.ProjCost[i].SUM_JKH;
                SumTNT_Mileage = summ_projCost.ProjCost[i].SUM_TNT_MILEAGE;
                SumMilling = summ_projCost.ProjCost[i].SUM_MILLING;
                SumMisc = summ_projCost.ProjCost[i].SUM_MISC;
                SumContract = summ_projCost.ProjCost[i].SUM_CONTRACT;
            }
            System.Diagnostics.Debug.WriteLine("BOQ LIST : 8");

            string PuList = "";
            if (check == "JKH")
            {
                using (Entities ctxData = new Entities())
                {
                    /*var query = from p in ctxData.WV_PU_MAST
                                where p.NETWORK_FLAG == "U" || p.NETWORK_FLAG == "G"
                                select new { p.PU_ID, p.NETWORK_FLAG ,p.PU_DESC };*/
                    var query = from p in ctxData.WV_PU_MAST
                                where p.NETWORK_FLAG == "U"
                                select new { p.PU_ID, p.NETWORK_FLAG, p.PU_DESC };

                    var queryPu = from p in ctxData.WV_BOQ_DATA
                                  where p.SCHEME_NAME == schemeName
                                  select new { p.PU_ID, p.PU_DESC };

                    foreach (var a in query.Distinct().OrderBy(it => it.PU_ID))
                    {
                        int checkPu = 0;
                        foreach (var b in queryPu)
                        {
                            if (b.PU_ID.Count() == 4)
                                checkPu = 1;
                        }
                        if (checkPu == 0)
                            PuList = PuList + a.PU_ID + ": (" + a.NETWORK_FLAG + ") " + a.PU_DESC + "|";
                    }
                }
            }

            string ContractList = "";
            string SystemPrice = "false";
            if (check == "CONTRACT")
            {
                using (Entities ctxData = new Entities())
                {
                    int count = (from a in ctxData.WV_BOQ_DATA
                                 where a.SCHEME_NAME == schemeName
                                 select a).Count();
                    if (count == 0)
                        SystemPrice = "false";
                    else
                    {
                        int countPackage = (from b in ctxData.WV_PACKAGE_MAST
                                            join c in ctxData.WV_BOQ_DATA on b.CONTRACT_NO equals c.CONTRACT_NO
                                            where c.SCHEME_NAME == schemeName
                                            select b).Count();
                        if (countPackage == 0)
                            SystemPrice = "false";
                        else
                            SystemPrice = "true";
                    }
                }

                using (Entities ctxData = new Entities())
                {
                    var query = from p in ctxData.WV_BOQ_DATA
                                where p.SCHEME_NAME == schemeName
                                select new { p.CONTRACT_NO };

                    foreach (var a in query.Distinct().OrderBy(it => it.CONTRACT_NO))
                    {
                        ContractList = ContractList + a.CONTRACT_NO + "|";
                    }
                }
            }//-----------------------------------------------------------------------------------------
            //System.Diagnostics.Debug.WriteLine("BOQ LIST : 2");
            //WebService._base.Estimate_Lab estimateLab = new WebService._base.Estimate_Lab();
            //estimateLab = myWebService.GetEstimate_Lab(schemeName);
            //string EstLabBuruh = "";
            //string EstLabBiasa = "";
            //string EstLabKanan = "";
            //string EstLabPembantu = "";
            //string EstLabTotal = "";
            //for (int i = 0; i < estimateLab.LabList.Count; i++)
            //{
            //    EstLabBuruh = estimateLab.LabList[i].LAB_BURUH;
            //    EstLabBiasa = estimateLab.LabList[i].LAB_BIASA;
            //    EstLabKanan = estimateLab.LabList[i].LAB_KANAN;
            //    EstLabPembantu = estimateLab.LabList[i].LAB_PEMBANTU;
            //    EstLabTotal = estimateLab.LabList[i].LAB_EST_TOTAL;
            //}

            //WebService._base.Estimate_Inc estimateInc = new WebService._base.Estimate_Inc();
            //estimateInc = myWebService.GetEstimate_Inc(schemeName);
            //string EstIncSupHours = "";
            //string EstIncOtSupHours = "";
            //string EstIncValue = "";
            //string EstIncOtValue = "";
            //string EstIncExecMileage = "";
            //string EstIncNonExecMileage = "";
            //string EstIncTotalHours = "";
            //string EstIncTotalValue = "";
            //string EstIncTotalMileage = "";
            //for (int i = 0; i < estimateInc.IncList.Count; i++)
            //{
            //    EstIncSupHours = estimateInc.IncList[i].INC_CTRT_SUP_HOURS;
            //    EstIncOtSupHours = estimateInc.IncList[i].INC_CTRT_OT_SUP_HOURS;
            //    EstIncValue = estimateInc.IncList[i].INC_CTRT_VALUE;
            //    EstIncOtValue = estimateInc.IncList[i].INC_CTRT_OT_VALUE;
            //    EstIncExecMileage = estimateInc.IncList[i].INC_EXEC_MILEAGE;
            //    EstIncNonExecMileage = estimateInc.IncList[i].INC_NON_EXEC_MILEAGE;
            //    EstIncTotalHours = estimateInc.IncList[i].INC_TOTAL_HOURS;
            //    EstIncTotalValue = estimateInc.IncList[i].INC_TOTAL_VALUE;
            //    EstIncTotalMileage = estimateInc.IncList[i].INC_TOTAL_MILEAGE;
            //}

            //WebService._base.Estimate_BPlanEstimate estimateBP = new WebService._base.Estimate_BPlanEstimate();
            //estimateBP = myWebService.GetBussinessPlan_Estimate(schemeName);
            //string EstBPJamBuruh = "";
            //string EstBPJamNilai = "";
            //string EstBPOTBuruh = "";
            //string EstBPOTNilai = "";
            //string EstBPJKHBahan = "";
            //string EstBPJKHPelbagai = "";
            //string EstBPTMBahan = "";
            //string EstBPTMPelbagai = "";
            //for (int i = 0; i < estimateBP.BPlanEstimate.Count; i++)
            //{
            //    EstBPJamBuruh = estimateBP.BPlanEstimate[i].BP_JAM_BURUH;
            //    EstBPJamNilai = estimateBP.BPlanEstimate[i].BP_JAM_NILAI;
            //    EstBPOTBuruh = estimateBP.BPlanEstimate[i].BP_OT_BURUH;
            //    EstBPOTNilai = estimateBP.BPlanEstimate[i].BP_OT_NILAI;
            //    EstBPJKHBahan = estimateBP.BPlanEstimate[i].BP_JKH_BAHAN;
            //    EstBPJKHPelbagai = estimateBP.BPlanEstimate[i].BP_JKH_PEL;
            //    EstBPTMBahan = estimateBP.BPlanEstimate[i].BP_TM_BAHAN;
            //    EstBPTMPelbagai = estimateBP.BPlanEstimate[i].BP_TM_PEL;
            //}
           // System.Diagnostics.Debug.WriteLine("RECORD : " + res);
            //ContractList = "test";
            string Proj_no = "";
            if (check == "JKH")
            {
                Proj_no = Proj_no_isp;
            }
            else
            {
                Proj_no = Proj_no_osp;
            }
            return Json(new
            {
                //txtEstBPJamBuruh = EstBPJamBuruh,
                //txtEstBPJamNilai = EstBPJamNilai,
                //txtEstBPOTBuruh = EstBPOTBuruh,
                //txtEstBPOTNilai = EstBPOTNilai,
                //txtEstBPJKHBahan = EstBPJKHBahan,
                //txtEstBPJKHPelbagai = EstBPJKHPelbagai,
                //txtEstBPTMBahan = EstBPTMBahan,
                //txtEstBPTMPelbagai = EstBPTMPelbagai,
                txtSchName = schemeName,
                //txtEstLabBuruh = EstLabBuruh,
                //txtEstLabBiasa = EstLabBiasa,
                //txtEstLabKanan = EstLabKanan,
                //txtEstLabPembantu = EstLabPembantu,
                //txtEstLabTotal = EstLabTotal,
                //txtEstIncSupHours = EstIncSupHours,
                //txtEstIncOtSupHours = EstIncOtSupHours,
                //txtEstIncValue = EstIncValue,
                //txtEstIncOtValue = EstIncOtValue,
                //txtEstIncExecMileage = EstIncExecMileage,
                //txtEstIncNonExecMileage = EstIncNonExecMileage,
                //txtEstIncTotalHours = EstIncTotalHours,
                //txtEstIncTotalValue = EstIncTotalValue,
                //txtEstIncTotalMileage = EstIncTotalMileage,
                //mainList = MainList,
                mat = mat,
                SystemPrice = SystemPrice,
                ContractList = ContractList,
                PuList = PuList,
                checkGEMS = gems,
                txtSumLbrOT = SumLbrOT,
                txtSumLbrSalary = SumLbrSalary,
                txtSumMaterial = SumMaterial,
                txtSumJKH = SumJKH,
                txtSumTNT_Mileage = SumTNT_Mileage,
                txtSumMilling = SumMilling,
                txtSumMisc = SumMisc,
                txtSumContract = SumContract,
                errorRecord = dataError,
                record = res,
                proj_no = Proj_no
            }, JsonRequestBehavior.AllowGet);
        }

        #region Fatihin 24042019 
        public ActionResult CheckProjNoBOQ(string schemeName)
        {
            int checkProjNoOSP = 0;
            int checkProjNoISP = 0;

            using (Entities ctxData = new Entities())
            {
                
                checkProjNoOSP = (from a in ctxData.G3E_JOB
                              where a.SCHEME_NAME == schemeName && a.PROJECT_NO == null
                              select a).Count();

                checkProjNoISP = (from a in ctxData.WV_ISP_JOB
                                  where a.SCHEME_NAME == schemeName && a.PROJECT_NO == null
                                  select a).Count();
            }
            return Json(new
            {
                checkProjNoOSP = checkProjNoOSP,
                checkProjNoISP = checkProjNoISP
                
            }, JsonRequestBehavior.AllowGet);
        }
        #endregion

        public ActionResult BOQ_Generate(string schemeName, string conndb)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            string isposp = "";
            string output = "A";
            using (Entities ctxdata = new Entities())
            {
                isposp = (from d in ctxdata.REF_JOB_TYPES
                          join a in ctxdata.G3E_JOB on d.JOB_TYPE equals a.JOB_TYPE
                          where a.SCHEME_NAME == schemeName
                          select d.ISPOSP).Single();
            }
            bool populateDataISP = true;
            bool checkISP = false;
            //Fatihin 08052019 
            //if (isposp == "ISP" || isposp == "ISPOSP")
            if(isposp == "ISP")
            {
                //update by atiqah 22/10/2015 - Disable BOQ ISP
                populateDataISP = myWebService.BOQ_POPULATE_ISP(schemeName, conndb);
                checkISP = true;
            }
            //// output = myWebService.BOQ_POPULATE_ISP(schemeName);
            bool success = true;
            if (populateDataISP)
            {
                if (checkISP)
                {
                    success = myWebService.BOQ_GENERATE_CONTRACT_ISP(schemeName);
                }
                //success = myWebService.BOQ_GENERATE(schemeName, checkISP.ToString());

            }
            success = myWebService.BOQ_GENERATE(schemeName, checkISP.ToString());
            //WebView.Controllers.BOQEngine engine = new BOQEngine();
            //using (OleDbConnection conn = UtilityDb.GetConnection("nepstrn", "neps", "nepstrn"))
            //{
            //    engine.StartProcess(schemeName, conn);
            //}

            //WebService._base.OSP_NETELEM BOQ_GEN = new WebService._base.OSP_NETELEM();
            //BOQ_GEN = myWebService.GetOSP_BOQ(schemeName);

            //WebService._base.ISP_NETELEM ISP_BOQ_GEN = new WebService._base.ISP_NETELEM();
            //ISP_BOQ_GEN = myWebService.GetISP_NETELEM(schemeName);

            //bool success = true;
            //string output = "";
            //success = myWebService.DeleteBOQ(schemeName);
            //System.Diagnostics.Debug.WriteLine("boq count : " + BOQ_GEN.NETELEMList.Count);
            //for (int i = 0; i < BOQ_GEN.NETELEMList.Count; i++)
            //{

            //    System.Diagnostics.Debug.WriteLine("checkStat " + i+" :" + (BOQ_GEN.NETELEMList[i].CHECK_PU_CONTRACT == "PU"));
            //    if (BOQ_GEN.NETELEMList[i].CHECK_PU_CONTRACT == "PU")
            //    {
            //        System.Diagnostics.Debug.WriteLine("CONTROLLER PUID : " + BOQ_GEN.NETELEMList[i].NET_PU_ID);
            //        string checkPuId = "NO PU ID";
            //        if (BOQ_GEN.NETELEMList[i].NET_PU_ID != checkPuId)
            //        {
            //            success = myWebService.AddBOQ_PU(BOQ_GEN.NETELEMList[i]);
            //        }
            //        else if (BOQ_GEN.NETELEMList[i].NET_PU_ID == checkPuId)
            //        {
            //            output += BOQ_GEN.NETELEMList[i].NET_MIN_MATERIAL;

            //            output += "!";
            //            System.Diagnostics.Debug.WriteLine("OUTPUT " + output);
            //        }
            //    }
            //    else if (success && BOQ_GEN.NETELEMList[i].CHECK_PU_CONTRACT == "CONTRACT")
            //    {
            //        success = myWebService.AddBOQ_Contract(BOQ_GEN.NETELEMList[i]);
            //    }
            //}

            //for (int i = 0; i < ISP_BOQ_GEN.NETELEMList.Count; i++)
            //{
            //    string minMaterial = ISP_BOQ_GEN.NETELEMList[i].NET_MIN_MATERIAL;
            //    string ispOsp = ISP_BOQ_GEN.NETELEMList[i].ISP_OSP;
            //    string puid = ISP_BOQ_GEN.NETELEMList[i].NET_PU_ID;
            //    success = myWebService.DeleteBOQ(schemeName);
            //    if (success && ISP_BOQ_GEN.NETELEMList[i].NET_MIN_MATERIAL.Count() == 8)
            //    {
            //        success = myWebService.AddBOQ_PU(ISP_BOQ_GEN.NETELEMList[i]);
            //    }
            //    else if (success && ISP_BOQ_GEN.NETELEMList[i].NET_MIN_MATERIAL.Count() > 8)
            //    {
            //        success = myWebService.AddBOQ_Contract(ISP_BOQ_GEN.NETELEMList[i]);
            //    }
            //}

            return Json(new
            {
                schemename = schemeName,
                record = checkISP.ToString(),
                result = success
            }, JsonRequestBehavior.AllowGet);
        }


         public ActionResult Delete_PU_BOQ(string schemeName, string puid, string billrate)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool success = true;
            success = myWebService.Delete_BOQ_PUID(schemeName, puid, billrate);
            return Json(new
            {
                schemename = schemeName,
                result = success
            }, JsonRequestBehavior.AllowGet);
        }

         public ActionResult Delete_CONTRACT_BOQ(string schemeName, string contractNo, string itemNo)
         {
             WebView.WebService._base myWebService;
             myWebService = new WebService._base();

             bool success = true;
             success = myWebService.Delete_BOQ_Contract(schemeName, contractNo, itemNo);
             return Json(new
             {
                 schemename = schemeName,
                 result = success
             }, JsonRequestBehavior.AllowGet);
         }

         public ActionResult BOQ_Generate_Contract(string schemeName, string SystemPrice, string ConnDB)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            string isposp = "";
            string output = "A";
            using (Entities ctxdata = new Entities())
            {
                int count = (from d in ctxdata.REF_JOB_TYPES
                             join a in ctxdata.G3E_JOB on d.JOB_TYPE equals a.JOB_TYPE
                             where a.SCHEME_NAME == schemeName
                             select d.ISPOSP).Count();
                if (count != 0)
                {
                    output += "B";
                    isposp = (from d in ctxdata.REF_JOB_TYPES
                              join a in ctxdata.G3E_JOB on d.JOB_TYPE equals a.JOB_TYPE
                              where a.SCHEME_NAME == schemeName
                              select d.ISPOSP).Single();
                }
                else
                {
                    output += "C";
                    isposp = (from d in ctxdata.REF_JOB_TYPES
                              join a in ctxdata.WV_ISP_JOB on d.JOB_TYPE equals a.JOB_TYPE
                              where a.SCHEME_NAME == schemeName
                              select d.ISPOSP).Single();
                }
            }
            bool populateDataISP = true;
            bool checkISP = false;
            if (isposp == "ISP" || isposp == "ISPOSP")
            {
                output += "D";
                populateDataISP = myWebService.BOQ_POPULATE_ISP(schemeName, ConnDB);
                output += populateDataISP.ToString();
                checkISP = true;
            }
           
            bool success = true;
            populateDataISP = true;
            if (populateDataISP)
            {
                output += "E";
                if (checkISP)
                {
                    output += "F";
                    success = myWebService.BOQ_GENERATE_CONTRACT_ISP(schemeName);
                }
                success = myWebService.BOQ_GENERATE_CONTRACT(schemeName, checkISP.ToString());
            }

            if (SystemPrice == "true")
            {
                NEPS.BOQ.Classes.ContractBOQ.Engine engine = new NEPS.BOQ.Classes.ContractBOQ.Engine();
                string[] arr = ConnDB.Split('|');
                string IP = arr[0];
                string SID = arr[1];
                string Password = arr[2];
                using (OracleConnection conn = UtilityDb2.GetConnection(arr[0], arr[1], arr[2]))
                {
                    List<Item> source = new List<Item>();
                    List<Package> packages = null;
                    List<Item> individualItems = null;
                    //string schemeName = "BJL31-ISP-449-2012";
                    string packagesToCreate = "1,1,1,2,2,2,3,3,3"; // this is hardcoded for now, but Atiqah to discuss with Kamal on what's the best configuration
                    Engine.Run
                        (false, // true = write the output file to help debugging. Should specify False in real implementation
                        conn, schemeName, packagesToCreate,
                        out packages,
                        out individualItems);

                }


                //Engine engine = new Engine();

                //List<Item> source = new List<Item>();
                //List<Package> packages = null;
                //List<Item> individualItems = null;
                ////string schemeName = "TEST";
                //output = engine.Execute(true, // true = write the output file to help debugging. Should specify False in real implementation
                //         schemeName,
                //         out packages,
                //         out individualItems);

                   //NEPS.BOQ.Classes.ContractBOQ.Engine engine = new NEPS.BOQ.Classes.ContractBOQ.Engine();
                   //using (OracleConnection conn = UtilityDb2.GetConnection("10.14.61.177/NEPSTRN", "NEPS", "nepstrn"))
                   //{
                   //    List<NEPS.BOQ.Classes.ContractBOQ.Item> source = new List<NEPS.BOQ.Classes.ContractBOQ.Item>();
                   //    List<NEPS.BOQ.Classes.ContractBOQ.Package> packages = null;
                   //    List<NEPS.BOQ.Classes.ContractBOQ.Item> individualItems = null;
                       
                   //    string packagesToCreate = "1,1,1,2,2,2,3,3,3";
                   //    NEPS.BOQ.Classes.ContractBOQ.Engine.Run
                   //        (false, // true = write the output file to help debugging. Should specify False in real implementation
                   //        conn, schemeName, packagesToCreate,
                   //        out packages,
                   //        out individualItems);

                   //}
                   
                    //Engine.Run
                    //    (true, // true = write the output file to help debugging. Should specify False in real implementation
                    //    schemeName,
                    //    out packages,
                    //    out individualItems);
                   
               
            }

           // string output = "A";

            //WebService._base.OSP_NETELEM BOQ_GEN = new WebService._base.OSP_NETELEM();
            //BOQ_GEN = myWebService.GetOSP_BOQ(schemeName);

            //WebService._base.ISP_NETELEM ISP_BOQ_GEN = new WebService._base.ISP_NETELEM();
            //ISP_BOQ_GEN = myWebService.GetISP_NETELEM(schemeName);

            //bool success = true;
            //string output = "";
            //success = myWebService.DeleteBOQ(schemeName);
            //System.Diagnostics.Debug.WriteLine("boq count : " + BOQ_GEN.NETELEMList.Count);
            //for (int i = 0; i < BOQ_GEN.NETELEMList.Count; i++)
            //{

            //    System.Diagnostics.Debug.WriteLine("checkStat " + i+" :" + (BOQ_GEN.NETELEMList[i].CHECK_PU_CONTRACT == "PU"));
            //    if (BOQ_GEN.NETELEMList[i].CHECK_PU_CONTRACT == "PU")
            //    {
            //        System.Diagnostics.Debug.WriteLine("CONTROLLER PUID : " + BOQ_GEN.NETELEMList[i].NET_PU_ID);
            //        string checkPuId = "NO PU ID";
            //        if (BOQ_GEN.NETELEMList[i].NET_PU_ID != checkPuId)
            //        {
            //            success = myWebService.AddBOQ_PU(BOQ_GEN.NETELEMList[i]);
            //        }
            //        else if (BOQ_GEN.NETELEMList[i].NET_PU_ID == checkPuId)
            //        {
            //            output += BOQ_GEN.NETELEMList[i].NET_MIN_MATERIAL;

            //            output += "!";
            //            System.Diagnostics.Debug.WriteLine("OUTPUT " + output);
            //        }
            //    }
            //    else if (success && BOQ_GEN.NETELEMList[i].CHECK_PU_CONTRACT == "CONTRACT")
            //    {
            //        success = myWebService.AddBOQ_Contract(BOQ_GEN.NETELEMList[i]);
            //    }
            //}

            //for (int i = 0; i < ISP_BOQ_GEN.NETELEMList.Count; i++)
            //{
            //    string minMaterial = ISP_BOQ_GEN.NETELEMList[i].NET_MIN_MATERIAL;
            //    string ispOsp = ISP_BOQ_GEN.NETELEMList[i].ISP_OSP;
            //    string puid = ISP_BOQ_GEN.NETELEMList[i].NET_PU_ID;
            //    success = myWebService.DeleteBOQ(schemeName);
            //    if (success && ISP_BOQ_GEN.NETELEMList[i].NET_MIN_MATERIAL.Count() == 8)
            //    {
            //        success = myWebService.AddBOQ_PU(ISP_BOQ_GEN.NETELEMList[i]);
            //    }
            //    else if (success && ISP_BOQ_GEN.NETELEMList[i].NET_MIN_MATERIAL.Count() > 8)
            //    {
            //        success = myWebService.AddBOQ_Contract(ISP_BOQ_GEN.NETELEMList[i]);
            //    }
            //}
            return Json(new
            {
                schemename = schemeName,
                record = output,
                result = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult BOQ_GEMS(string schemeName)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            string check = "";
            bool success = true;
            using (Entities ctxData = new Entities())
            {
                int queryCount = (from q in ctxData.G3E_JOB
                            where q.SCHEME_NAME == schemeName
                            select q).Count();
                string checkProjNoISP = "";
                string checkProjNoOSP = "";
                if (queryCount != 0)
                {
                    checkProjNoOSP = (from q in ctxData.G3E_JOB
                                             where q.SCHEME_NAME == schemeName
                                             select q.PROJECT_NO).Single();
                }
                else
                {
                    checkProjNoISP = (from q in ctxData.WV_ISP_JOB
                                             where q.SCHEME_NAME == schemeName
                                             select q.PROJECT_NO).Single();
                }
                System.Diagnostics.Debug.WriteLine(checkProjNoOSP);
                System.Diagnostics.Debug.WriteLine(checkProjNoISP);

                var findusername = (from q in ctxData.WV_USER
                                    where q.USERNAME.ToUpper() == User.Identity.Name.ToUpper() || q.USERNAME.ToLower() == User.Identity.Name.ToLower()
                                    select q.USERNAME).Single();

                if ((checkProjNoOSP != "") || (checkProjNoISP != ""))
                {
                    if (queryCount == 0)
                    {
                        check = "CONTRACT";
                        System.Diagnostics.Debug.WriteLine(check);
                        success = myWebService.BOQ_GEMS(schemeName, check, findusername);
                    }
                    else
                    {
                        var query = from q in ctxData.G3E_JOB
                                    where q.SCHEME_NAME == schemeName
                                    select new { q.JOB_TYPE, q.G3E_OWNER };
                        foreach (var a in query)
                        {
                            if (a.JOB_TYPE == "Civil" || a.JOB_TYPE == "E/Side" || a.JOB_TYPE == "D/Side" || a.JOB_TYPE == "Fiber E/Side" || a.JOB_TYPE == "HSBB E/Side" || a.JOB_TYPE == "HSBB D/Side" || a.JOB_TYPE == "Fiber Trunk" || a.JOB_TYPE == "Fiber Junction" || a.JOB_TYPE == "Others")
                            {
                                check = "JKH";
                                success = myWebService.BOQ_GEMS(schemeName, check, a.G3E_OWNER);
                            }
                            else
                            {
                                check = "CONTRACT";
                                success = myWebService.BOQ_GEMS(schemeName, check, a.G3E_OWNER);
                            }
                        }
                    }
                }
                else
                {
                    success = false;
                }
               
            }
            System.Diagnostics.Debug.WriteLine(success);
           
            return Json(new
            {
                schemename = schemeName,
                result = success
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult GetBillRate(string PU_ID, string schemeName)
        {
            string billrate = "";
            string pu_desc = "";
            using (Entities ctxdata = new Entities())
            {
                var queryPu = from p in ctxdata.WV_BOQ_DATA
                              where p.SCHEME_NAME == schemeName && p.PU_ID == PU_ID
                              select new { p.RATE_INDICATOR };
                
                var query = from a in ctxdata.WV_PU_MAST
                            where a.PU_ID == PU_ID 
                            orderby a.BILL_RATE
                            select new { a.BILL_RATE };

                var queryDESC = (from d in ctxdata.WV_PU_MAST
                                 where d.PU_ID == PU_ID
                                 select new { d.PU_DESC });

                
                foreach (var a in queryDESC)
                {
                    pu_desc = a.PU_DESC;
                }

                foreach (var a in query)
                {
                    int checkPu = 0;
                    foreach (var b in queryPu)
                    {
                        if (b.RATE_INDICATOR == a.BILL_RATE )
                            checkPu = 1;
                    }
                    if (checkPu == 0)
                    {
                        if (a.BILL_RATE == "D")
                            billrate = billrate + "!" + "DAY";
                        else if (a.BILL_RATE == "N")
                            billrate = billrate + "!" + "NIGHT";
                        else if (a.BILL_RATE == "W")
                            billrate = billrate + "!" + "WEEKEND";
                        else if (a.BILL_RATE == "P")
                            billrate = billrate + "!" + "HOLIDAY";

                    }
                }
                System.Diagnostics.Debug.WriteLine(billrate);
            }

         
            return Json(new
            {
                pu_desc = pu_desc,
                billrate = billrate
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult GetItemNo(string contractNo, string schemeName)
        {
            string itemNo = "";
            using (Entities ctxdata = new Entities())
            {
                var queryCon = from p in ctxdata.WV_BOQ_DATA
                              where p.SCHEME_NAME == schemeName && p.CONTRACT_NO == contractNo
                              select new { p.ITEM_NO, p.CONTRACT_DESC };

                var query = from a in ctxdata.WV_CONTRACT_MAST
                            where a.CONTRACT_NO == contractNo  && a.NETWORK_FLAG == "U"
                            orderby a.ITEM_NO
                            select new { a.ITEM_NO, a.CONTRACT_DESC };
               
                foreach (var a in query)
                {
                    int checkPu = 0;
                    foreach (var b in queryCon)
                    {
                        if (Convert.ToInt32(b.ITEM_NO) == Convert.ToInt32(a.ITEM_NO))
                            checkPu = 1;
                    }
                    if (checkPu == 0)
                    {
                        itemNo = itemNo + "!" + a.ITEM_NO + "- " + a.CONTRACT_DESC;
                    }
                }
                System.Diagnostics.Debug.WriteLine(itemNo);
            }


            return Json(new
            {
                itemNo = itemNo
            }, JsonRequestBehavior.AllowGet);
        }

        
        [HttpPost]
        public ActionResult MaintenanceDetailsContractBOQ(string ContractNo_ItemNo, string schemeName)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.Estimate_Main estimateMain = new WebService._base.Estimate_Main();
            estimateMain = myWebService.GetEstimate_Main_Contract(ContractNo_ItemNo, schemeName);


            
            string EstMainConDesc = "";
            string EstMainNetPrice = "";
            string contractNo = "";
            string itemNo = "";
            string qty = "";
            for (int i = 0; i < estimateMain.MainList.Count; i++)
            {
                //System.Diagnostics.Debug.WriteLine(estimateMain.MainList[i].MAIN_PU_ID + " : " + estimateMain.MainList[i].MAIN_PU_DESC + " : " + estimateMain.MainList[i].MAIN_PU_MAT_PR);
                EstMainConDesc = estimateMain.MainList[i].MAIN_CONTRACT_DESC;
                EstMainNetPrice = estimateMain.MainList[i].MAIN_NET_PRICE;
                contractNo = estimateMain.MainList[i].MAIN_CONTRACT_NO;
                itemNo = estimateMain.MainList[i].MAIN_ITEM_NO;
                qty = estimateMain.MainList[i].MAIN_PU_QTY;
            }

            return Json(new
            {
                txtEstMainConDesc = EstMainConDesc,
                contractNo = contractNo,
                itemNo = itemNo,
                qty = qty,
                txtEstMainNetPrice = EstMainNetPrice
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult MaintenanceDetailsContract(string ContractNo_ItemNo)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            
            WebService._base.Estimate_Main estimateMain = new WebService._base.Estimate_Main();
            estimateMain = myWebService.GetEstimate_Main_Contract(ContractNo_ItemNo, "");



            string EstMainConDesc = "";
            string EstMainNetPrice = "";
            string contractNo = "";
            string itemNo = "";
            for (int i = 0; i < estimateMain.MainList.Count; i++)
            {
                //System.Diagnostics.Debug.WriteLine(estimateMain.MainList[i].MAIN_PU_ID + " : " + estimateMain.MainList[i].MAIN_PU_DESC + " : " + estimateMain.MainList[i].MAIN_PU_MAT_PR);
                EstMainConDesc = estimateMain.MainList[i].MAIN_CONTRACT_DESC;
                EstMainNetPrice = estimateMain.MainList[i].MAIN_NET_PRICE;
                contractNo = estimateMain.MainList[i].MAIN_CONTRACT_NO;
                itemNo = estimateMain.MainList[i].MAIN_ITEM_NO;
            }

            return Json(new
            {
                txtEstMainConDesc = EstMainConDesc,
                contractNo = contractNo,
                itemNo = itemNo,
                txtEstMainNetPrice = EstMainNetPrice
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult MaintenanceDetails(string PU_BILLRATE)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.Estimate_Main estimateMain = new WebService._base.Estimate_Main();
            estimateMain = myWebService.GetEstimate_Main(PU_BILLRATE);
            string EstMainPUId = "";
            string EstMainPUDesc = "";
            string EstMainMatPrice = "";
            string EstMainInstPrice = "";
            string EstMainBillRate = "";
            for (int i = 0; i < estimateMain.MainList.Count; i++)
            {
                System.Diagnostics.Debug.WriteLine(estimateMain.MainList[i].MAIN_PU_ID + " : " + estimateMain.MainList[i].MAIN_PU_DESC + " : " + estimateMain.MainList[i].MAIN_PU_MAT_PR);
                EstMainPUId = estimateMain.MainList[i].MAIN_PU_ID;
                EstMainPUDesc = estimateMain.MainList[i].MAIN_PU_DESC;
                EstMainMatPrice = estimateMain.MainList[i].MAIN_PU_MAT_PR;
                EstMainInstPrice = estimateMain.MainList[i].MAIN_PU_INST_PR;
                EstMainBillRate = estimateMain.MainList[i].MAIN_BILL_RATE;
            }

            return Json(new
            {
                txtEstMainPUId = EstMainPUId,
                txtEstMainPUDesc = EstMainPUDesc,
                txtEstMainMatPrice = EstMainMatPrice,
                txtEstMainInstPrice = EstMainInstPrice,
                txtEstMainBillRate = EstMainBillRate
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult MaintenanceDetailsBOQ(string PU_BILLRATE, string schemeName)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.Estimate_Main estimateMain = new WebService._base.Estimate_Main();
            estimateMain = myWebService.GetEstimate_Main_BOQ(PU_BILLRATE, schemeName);
            string EstMainPUId = "";
            string EstMainPUDesc = "";
            string EstMainMatPrice = "";
            string EstMainInstPrice = "";
            string EstMainBillRate = "";
            string EstMainPUQty = "";
            string EstMainOldMatPr = ""; 
           
            for (int i = 0; i < estimateMain.MainList.Count; i++)
            {
                System.Diagnostics.Debug.WriteLine(estimateMain.MainList[i].MAIN_PU_ID + " : " + estimateMain.MainList[i].MAIN_PU_DESC + " : " + estimateMain.MainList[i].MAIN_PU_MAT_PR);
                EstMainPUId = estimateMain.MainList[i].MAIN_PU_ID;
                EstMainPUDesc = estimateMain.MainList[i].MAIN_PU_DESC;
                EstMainMatPrice = estimateMain.MainList[i].MAIN_PU_MAT_PR;
                EstMainInstPrice = estimateMain.MainList[i].MAIN_PU_INST_PR;
                EstMainBillRate = estimateMain.MainList[i].MAIN_BILL_RATE;
                EstMainPUQty = estimateMain.MainList[i].MAIN_PU_QTY;
                EstMainOldMatPr = estimateMain.MainList[i].MAIN_OLD_MAT_PR;
            
            }

            return Json(new
            {
                txtEstMainPUId = EstMainPUId,
                txtEstMainPUDesc = EstMainPUDesc,
                txtEstMainMatPrice = EstMainMatPrice,
                txtEstMainInstPrice = EstMainInstPrice,
                txtEstMainBillRate = EstMainBillRate,
                txtEstMainPUQty = EstMainPUQty,
                txtEstMainOldMatPr = EstMainOldMatPr
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult ReportLabourWorkingHours(string id)
        {
            string[] arr = id.Split('/');
            string schemeName = arr[0];
            string excel_html = arr[1];

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            ViewBag.date = DateTime.Now.ToString("d/M/yyyy");
            ViewBag.schemeName = schemeName;
            ViewBag.projNo = "";//myWebService.BOQ_ProjNo(id);

            WebService._base.Estimate_Lab estimateLab = new WebService._base.Estimate_Lab();
            estimateLab = myWebService.GetEstimate_Lab(schemeName);

            for (int i = 0; i < estimateLab.LabList.Count; i++)
            {
                ViewBag.EstLabBuruh = estimateLab.LabList[i].LAB_BURUH;
                ViewBag.EstLabBiasa = estimateLab.LabList[i].LAB_BIASA;
                ViewBag.EstLabKanan = estimateLab.LabList[i].LAB_KANAN;
                ViewBag.EstLabPembantu = estimateLab.LabList[i].LAB_PEMBANTU;
                ViewBag.EstLabTotal = estimateLab.LabList[i].LAB_EST_TOTAL;
            }
            if (excel_html == "EXCEL")
            {
                this.Response.AddHeader("Content-Disposition", "attachment; filename=report.xls");
                this.Response.ContentType = "application/vnd.ms-excel";
            }
            return View();
        }

        public ActionResult ReportIncidentialCost(string id)
        {
            string[] arr = id.Split('/');
            string schemeName = arr[0];
            string excel_html = arr[1];

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            ViewBag.date = DateTime.Now.ToString("d/M/yyyy");
            ViewBag.schemeName = schemeName;
            ViewBag.projNo = "";//myWebService.BOQ_ProjNo(id);

            WebService._base.Estimate_Inc estimateInc = new WebService._base.Estimate_Inc();
            estimateInc = myWebService.GetEstimate_Inc(schemeName);

            for (int i = 0; i < estimateInc.IncList.Count; i++)
            {
                ViewBag.EstIncSupHours = estimateInc.IncList[i].INC_CTRT_SUP_HOURS;
                ViewBag.EstIncOtSupHours = estimateInc.IncList[i].INC_CTRT_OT_SUP_HOURS;
                ViewBag.EstIncValue = estimateInc.IncList[i].INC_CTRT_VALUE;
                ViewBag.EstIncOtValue = estimateInc.IncList[i].INC_CTRT_OT_VALUE;
                ViewBag.EstIncExecMileage = estimateInc.IncList[i].INC_EXEC_MILEAGE;
                ViewBag.EstIncNonExecMileage = estimateInc.IncList[i].INC_NON_EXEC_MILEAGE;
                ViewBag.EstIncTotalHours = estimateInc.IncList[i].INC_TOTAL_HOURS;
                ViewBag.EstIncTotalValue = estimateInc.IncList[i].INC_TOTAL_VALUE;
                ViewBag.EstIncTotalMileage = estimateInc.IncList[i].INC_TOTAL_MILEAGE;

            }
            if (excel_html == "EXCEL")
            {
                this.Response.AddHeader("Content-Disposition", "attachment; filename=report.xls");
                this.Response.ContentType = "application/vnd.ms-excel";
            }
            return View();
        }

        public ActionResult ReportBOQ(string id)
        {
            string[] arr = id.Split('|');
            string schemeName = arr[0];
            string excel_html = arr[1];

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            
            ViewBag.date = DateTime.Now.ToString("d/M/yyyy");
            ViewBag.schemeName = schemeName;
            //ViewBag.projNo = myWebService.BOQ_ProjNo(schemeName);
            using (Entities ctxdata = new Entities())
            {
                ViewBag.Exc = (from a in ctxdata.WV_EXC_MAST
                              join b in ctxdata.G3E_JOB on a.EXC_ABB.Trim() equals b.EXC_ABB.Trim()
                              where b.SCHEME_NAME == schemeName
                              select a.EXC_NAME).Single();

                ViewBag.projNo = (from  b in ctxdata.G3E_JOB 
                                   where b.SCHEME_NAME == schemeName
                                   select b.PROJECT_NO).Single();
                string projNo1 = ViewBag.projNo;

                ViewBag.projDesc = (from a in ctxdata.WV_GEM_PROJNO
                                    where a.PROJECT_NO.Trim() == projNo1
                                    select a.PROJ_DESC).FirstOrDefault();
            }

            WebService._base.OSPBOQ_MAIN_EXCEL BOQ_REPORT = new WebService._base.OSPBOQ_MAIN_EXCEL();
            string checkJKH_Contract = "JKH";
            BOQ_REPORT = myWebService.GetOSPBOQ_MAIN_Excel(0, 100, schemeName);
            int count = BOQ_REPORT.BOQ_MAIN_Excel.Count;
            ViewBag.Page = count/16;

            WebService._base.OSPBOQ_MAIN_EXCEL[] BOQ_REPORT2 = new WebService._base.OSPBOQ_MAIN_EXCEL[count / 16];
            for (int i = 0; i < (count/16); i++)
            {
                BOQ_REPORT2[i] = new WebService._base.OSPBOQ_MAIN_EXCEL();
                for (int j = (16*i); j < 16*(i+1); j++)
                {
                    WebView.WebService._base.BOQ_MAIN_EXCEL BOQ_MAIN_EXCEL = new WebView.WebService._base.BOQ_MAIN_EXCEL();
                    BOQ_MAIN_EXCEL.x_TOTAL_PRICE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_TOTAL_PRICE;
                    BOQ_MAIN_EXCEL.x_INSTALL_VALUE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_INSTALL_VALUE;
                    BOQ_MAIN_EXCEL.x_BQ_INSTALL_PRICE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_BQ_INSTALL_PRICE;
                    BOQ_MAIN_EXCEL.x_MATERIAL_VALUE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_MATERIAL_VALUE;
                    BOQ_MAIN_EXCEL.x_BQ_MAT_PRICE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_BQ_MAT_PRICE;
                    BOQ_MAIN_EXCEL.x_PU_UOM = BOQ_REPORT.BOQ_MAIN_Excel[j].x_PU_UOM;
                    BOQ_MAIN_EXCEL.x_BQ_CONT_PRICE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_BQ_CONT_PRICE;
                    BOQ_MAIN_EXCEL.x_CONT_VALUE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_CONT_VALUE;
                    BOQ_MAIN_EXCEL.x_TOTAL_VALUE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_TOTAL_VALUE;
                    BOQ_MAIN_EXCEL.x_PU_DESC = BOQ_REPORT.BOQ_MAIN_Excel[j].x_PU_DESC; // PLANT UNIT
                    BOQ_MAIN_EXCEL.x_CONTRACT_NO = BOQ_REPORT.BOQ_MAIN_Excel[j].x_CONTRACT_NO;
                    BOQ_MAIN_EXCEL.x_CONTRACT_DESC = BOQ_REPORT.BOQ_MAIN_Excel[j].x_CONTRACT_DESC;
                    BOQ_MAIN_EXCEL.x_PU_ID = BOQ_REPORT.BOQ_MAIN_Excel[j].x_PU_ID;
                    BOQ_MAIN_EXCEL.x_ISP_OSP = BOQ_REPORT.BOQ_MAIN_Excel[j].x_ISP_OSP;
                    BOQ_MAIN_EXCEL.x_PU_QTY = BOQ_REPORT.BOQ_MAIN_Excel[j].x_PU_QTY;
                    BOQ_MAIN_EXCEL.x_RATE_INDICATOR = BOQ_REPORT.BOQ_MAIN_Excel[j].x_RATE_INDICATOR;
                    BOQ_MAIN_EXCEL.x_ITEM_NO = BOQ_REPORT.BOQ_MAIN_Excel[j].x_ITEM_NO;
                    BOQ_MAIN_EXCEL.x_NET_PRICE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_NET_PRICE;

                    BOQ_REPORT2[i].BOQ_MAIN_Excel.Add(BOQ_MAIN_EXCEL);
                }
            }

            ViewBag.TotalInstallValue =  BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_INSTALL_VALUE_TOTAL;
            ViewBag.TotalMaterialValue = BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_MATERIAL_VALUE_TOTAL;
            ViewBag.TotalContractValue =   BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_CONT_VALUE_TOTAL;
            ViewBag.TotalValue =   BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_TOTAL_VALUE_TOTAL;

            if (myWebService.CheckBOQErrorPUID(schemeName) && myWebService.CheckBOQErrorMinMaterial(schemeName))
            {
                ViewBag.ErrorMsgInList = "Unmatched Plan Unit and Min Material";
                ViewBag.ErrorMsgInTotal = "WARNING! COST FIGURES ARE INACCURATE";
            }
            else if (myWebService.CheckBOQErrorPUID(schemeName) && myWebService.CheckBOQErrorContractNo(schemeName))
            {
                ViewBag.ErrorMsgInList = "Unmatched Contract No and Min Material";
                ViewBag.ErrorMsgInTotal = "WARNING! COST FIGURES ARE INACCURATE";
            }
            else if (myWebService.CheckBOQErrorPUID(schemeName))
            {
                ViewBag.ErrorMsgInList = "Unmatched Plan Unit";
                ViewBag.ErrorMsgInTotal = "WARNING! COST FIGURES ARE INACCURATE";
            }
            else if (myWebService.CheckBOQErrorContractNo(schemeName))
            {
                ViewBag.ErrorMsgInList = "Unmatched Contract No";
                ViewBag.ErrorMsgInTotal = "WARNING! COST FIGURES ARE INACCURATE";
            }
            else if (myWebService.CheckBOQErrorMinMaterial(schemeName))
            {
                ViewBag.ErrorMsgInList = "Unmatched Min Material";
                ViewBag.ErrorMsgInTotal = "WARNING! COST FIGURES ARE INACCURATE";
            }
            else
            {
                ViewBag.ErrorMsgInList = "";
                ViewBag.ErrorMsgInTotal = "";
            }
            for (int pageNum = 1; pageNum <= (count/16); pageNum++)
            {
                ViewData["BOQReportData" + pageNum] = BOQ_REPORT2[pageNum-1].BOQ_MAIN_Excel;
            }
            // this.Response.AddHeader("Content-Disposition", "Report.xls");
            if (excel_html == "EXCEL")
            {
                this.Response.AddHeader("Content-Disposition", "attachment; filename=" +schemeName+".xls");
                this.Response.ContentType = "application/vnd.ms-excel";
            }
            return View();
        }

        #region Mubin CR40-05062018
        public ActionResult ReportProjectCostBreakdown(string id)
        {
            string[] arr = id.Split('|');
            string schemeName = arr[0];
            string excel_html = arr[1];
            string jkhContract = arr[2];

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            int queryftth = 0;
            using (Entities1 ctxdata2 = new Entities1())
            {
                queryftth = (from a in ctxdata2.GC_NETELEM2
                             join b in ctxdata2.GC_FDC2 on a.G3E_FID equals b.G3E_FID
                             where a.SCHEME_NAME == schemeName && b.FDC_TYPE == "FTTH"
                             select a).Count();
            }

            using (Entities ctxdata = new Entities())
            {
                var queryFnoVdsl = (from a in ctxdata.GC_NETELEM
                                    where a.SCHEME_NAME == schemeName && (a.G3E_FNO == 9800 || a.G3E_FNO == 9200)
                                    select a).Count();

                var queryFnoMsan = (from a in ctxdata.GC_NETELEM
                                    where a.SCHEME_NAME == schemeName && a.G3E_FNO == 9100
                                    select a).Count();

                var queryProj = (from a in ctxdata.G3E_JOB
                                where a.SCHEME_NAME == schemeName
                                select new {a.PROJECT_NO, a.SCHEME_TYPE}).Single();

                var projDesc = "";
                if (queryProj.PROJECT_NO != "" && queryProj != null)
                {
                    projDesc = (from a in ctxdata.WV_GEM_PROJNO
                                where a.PROJECT_NO.Trim() == queryProj.PROJECT_NO.Trim()
                                select a.PROJ_DESC).FirstOrDefault();
                }

                var queryRegion = (from a in ctxdata.G3E_JOB
                                  join b in ctxdata.WV_PRT_MAST on a.SEGMENT.Trim() equals b.PRT_ID.Trim()
                                  join c in ctxdata.WV_REGION_MAST on b.REGION_ID equals c.REGION_ID
                                  where a.SCHEME_NAME == schemeName
                                  select new { c.REGION_NAME_GRN, c.NATIONWIDE_GRP }).Single();

                using (var xlsFile = new ExcelPackage())
                {
                    if (queryProj.SCHEME_TYPE == "D/Side" || queryProj.SCHEME_TYPE == "E/Side" || queryProj.SCHEME_TYPE == "Copper (Equip)" || queryProj.SCHEME_TYPE == "Copper Transfer")
                    {
                        #region copper

                        var tabDSC = xlsFile.Workbook.Worksheets.Add("StdCosting");
                        tabDSC.Cells.Style.Font.Name = "Arial";
                        tabDSC.Cells.Style.Font.Size = 8;
                        tabDSC.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        tabDSC.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                        tabDSC.Cells["A1:W228"].Style.Fill.BackgroundColor.SetColor(Color.White);
                        tabDSC.Cells["A1:A225,W1:W225"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        tabDSC.DefaultRowHeight = 13.2;
                        tabDSC.Column(1).Width = 0.75;
                        tabDSC.Column(2).Width = 0.63;
                        tabDSC.Column(3).Width = 3.11;
                        tabDSC.Column(4).Width = 3.11;
                        tabDSC.Column(5).Width = 12.33;
                        for (int i = 6; i < 9; i++)
                            tabDSC.Column(i).Width = 9.00;
                        tabDSC.Column(9).Width = 8.56;
                        tabDSC.Column(10).Width = 8.56;
                        for (int i = 11; i < 15; i++)
                            tabDSC.Column(i).Width = 6.78;
                        tabDSC.Column(15).Width = 3.67;
                        tabDSC.Column(16).Width = 3.67;
                        tabDSC.Column(17).Width = 1.33;
                        tabDSC.Column(18).Width = 7.56;
                        for (int i = 19; i < 22; i++)
                            tabDSC.Column(i).Width = 7.89;
                        tabDSC.Column(22).Width = 19.00;
                        tabDSC.Column(23).Width = 0.81;

                        tabDSC.Row(1).Height = 4.2;
                        tabDSC.Row(2).Height = 30;
                        tabDSC.Row(6).Height = 4.2;
                        tabDSC.Cells["C2:U10,C11:C13,L11:U18,C20:V20,C22:D22"].Style.Font.Bold = true;
                        tabDSC.Cells["C3,F3"].Style.Font.UnderLine = true;
                        tabDSC.Cells["L3"].Style.Font.Color.SetColor(Color.Red);
                        tabDSC.Cells["C2"].Style.Font.Size = 24;
                        tabDSC.Cells["C3:F3"].Style.Font.Size = 12;
                        tabDSC.Cells["L3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        tabDSC.Cells["L3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        tabDSC.Cells["M7:M11,I5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        tabDSC.Cells["N9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        tabDSC.Cells["F4:G4,J4:O4,F5:G5,J4:O5,F6:F7,N6:N11,F10:F14,I11:K18,C19:V19,C21:V21"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabDSC.Cells["E5,G5,I5,O5,E7:F7,M7:M11,N7:N11,E11:E14,F11:F14,H12:H18,I12:I18,J12:J18,K12:K18"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabDSC.Cells["F5:G5, J5:O5,F7,N7:N11,K12:K18"].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                        tabDSC.Cells["C20:Q21,V20:V21"].Style.Fill.BackgroundColor.SetColor(Color.LightPink);
                        tabDSC.Cells["R20:R21,U20:U21"].Style.Fill.BackgroundColor.SetColor(Color.SandyBrown);
                        tabDSC.Cells["S20:T21"].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);

                        tabDSC.Cells["C2"].Value = "STANDARD COSTING FOR COPPER";
                        tabDSC.Cells["C3"].Value = "PROJECT COST BREAKDOWN";
                        tabDSC.Cells["C5"].Value = "PROJECT NO :";
                        tabDSC.Cells["F5"].Value = queryProj.PROJECT_NO;
                        tabDSC.Cells["I5"].Value = "PROJECT TITLE:";
                        tabDSC.Cells["J5"].Value = projDesc;
                        tabDSC.Cells["C7"].Value = "REGION :";
                        tabDSC.Cells["F7"].Value = queryRegion.REGION_NAME_GRN;
                        tabDSC.Cells["M7"].Value = "SUPPLIER";
                        tabDSC.Cells["M8"].Value = "CABLE FOC";
                        tabDSC.Cells["M9"].Value = "LOCATION FACTOR";
                        if (queryRegion.NATIONWIDE_GRP == "WEST")
                            tabDSC.Cells["N9"].Value = 1.00;
                        else if (queryRegion.NATIONWIDE_GRP == "EAST")
                            tabDSC.Cells["N9"].Value = 1.15;
                        tabDSC.Cells["M11"].Value = "CLOSURE FOC";
                        if (queryRegion.NATIONWIDE_GRP == "WEST")
                            tabDSC.Cells["N10"].Value = "SEM";
                        else if (queryRegion.NATIONWIDE_GRP == "EAST")
                            tabDSC.Cells["N10"].Value = "SBH/SWK";
                        tabDSC.Cells["F10"].Value = "TOTAL";
                        tabDSC.Cells["I10"].Value = "PRKM";
                        tabDSC.Cells["C11"].Value = "MATERIALS COST";
                        tabDSC.Cells["C12"].Value = "INCIDENTAL COST";
                        tabDSC.Cells["C13"].Value = "LABOUR COST";

                        tabDSC.Cells["F14"].Formula = "=SUM(F11:F13)";

                        tabDSC.Cells["I12"].Value = "COPPER";
                        tabDSC.Cells["I13"].Value = "10pr";
                        tabDSC.Cells["I14"].Value = "20pr";
                        tabDSC.Cells["I15"].Value = "30pr";
                        tabDSC.Cells["I16"].Value = "50pr";
                        tabDSC.Cells["I17"].Value = "100pr";
                        tabDSC.Cells["I18"].Value = "200pr";

                        tabDSC.Cells["J12"].Value = "Dist(m)";
                      
                        using (Entities1 ctxdata2 = new Entities1())
                        {
                            //10pr
                            int total10pr = 0 ;
                            var check10pr = from a in ctxdata2.REF_COP_PR
                                             join b in ctxdata2.WV_MAT_DATA2 on a.PU_ID equals b.MAT_ID.Trim()
                                             where b.SCHEME_NAME == schemeName && a.PAIR== "10"
                                             select new
                                             {
                                                 b.MAT_QTY,
                                                 a.PU_ID,
                                                 a.PAIR
                                             };
                            if (check10pr.Count() > 0)
                            {
                                foreach( var a in check10pr)
                                    total10pr += Convert.ToInt32(a.MAT_QTY);
                           
                            }
                            var check10prInc = from a in ctxdata2.REF_COP_PR
                                               join b in ctxdata2.WV_BOQ_DATA2 on a.PU_ID equals b.PU_ID.Trim()
                                               where b.SCHEME_NAME == schemeName && a.PAIR == "10"
                                               select new 
                                               {
                                                   a.PAIR,
                                                   b.PU_QTY,
                                                   a.PU_ID

                                               };
                            total10pr = 0;
                            if (check10prInc.Count() > 0)
                            {
                                foreach( var a in check10prInc)
                                    total10pr += Convert.ToInt32(a.PU_QTY);
                           
                            }
                            tabDSC.Cells["J13"].Value = total10pr;
                            
                            //20PR
                            int total20pr = 0;
                            var check20pr = from a in ctxdata2.REF_COP_PR
                                            join b in ctxdata2.WV_MAT_DATA2 on a.PU_ID equals b.MAT_ID.Trim()
                                            where b.SCHEME_NAME == schemeName && a.PAIR == "20"
                                            select new
                                            {
                                                b.MAT_QTY,
                                                a.PU_ID,
                                                a.PAIR
                                            };
                            if (check20pr.Count() > 0)
                            {
                                foreach (var a in check20pr)
                                    total20pr += Convert.ToInt32(a.MAT_QTY);

                            }

                            var check20prInc = from a in ctxdata2.REF_COP_PR
                                            join b in ctxdata2.WV_BOQ_DATA2 on a.PU_ID equals b.PU_ID.Trim()
                                            where b.SCHEME_NAME == schemeName && a.PAIR == "20"
                                            select new
                                            {
                                                b.PU_QTY,
                                                a.PU_ID,
                                                a.PAIR
                                            };
                            total20pr = 0;
                            if (check20prInc.Count() > 0)
                            {
                                foreach (var a in check20prInc)
                                    total20pr += Convert.ToInt32(a.PU_QTY);

                            }
                            tabDSC.Cells["J14"].Value = total20pr;
                            //System.Diagnostics.Debug.WriteLine("--TotalQty20---" + total20pr);

                            //30pr
                            int total30pr = 0;
                            var check30pr = from a in ctxdata2.REF_COP_PR
                                            join b in ctxdata2.WV_MAT_DATA2 on a.PU_ID equals b.MAT_ID.Trim()
                                            where b.SCHEME_NAME == schemeName && a.PAIR == "30"
                                            select new
                                            {
                                                b.MAT_QTY,
                                                a.PU_ID,
                                                a.PAIR
                                            };
                            if (check30pr.Count() > 0)
                            {
                                foreach (var a in check30pr)
                                    total30pr += Convert.ToInt32(a.MAT_QTY);

                            }
                            var check30prInc = from a in ctxdata2.REF_COP_PR
                                            join b in ctxdata2.WV_BOQ_DATA2 on a.PU_ID equals b.PU_ID.Trim()
                                            where b.SCHEME_NAME == schemeName && a.PAIR == "30"
                                            select new
                                            {
                                                b.PU_QTY,
                                                a.PU_ID,
                                                a.PAIR
                                            };
                            total30pr = 0;
                            if (check30prInc.Count() > 0)
                            {
                                foreach (var a in check30prInc)
                                    total30pr += Convert.ToInt32(a.PU_QTY);

                            }
                            tabDSC.Cells["J15"].Value = total30pr;
                            
                            //50pr
                            int total50pr = 0;
                            var check50pr = from a in ctxdata2.REF_COP_PR
                                            join b in ctxdata2.WV_MAT_DATA2 on a.PU_ID equals b.MAT_ID.Trim()
                                            where b.SCHEME_NAME == schemeName && a.PAIR == "50"
                                            select new
                                            {
                                                b.MAT_QTY,
                                                a.PU_ID,
                                                a.PAIR
                                            };
                            if (check50pr.Count() > 0)
                            {
                                foreach (var a in check50pr)
                                    total50pr += Convert.ToInt32(a.MAT_QTY);

                            }

                             var check50prInc = from a in ctxdata2.REF_COP_PR
                                            join b in ctxdata2.WV_BOQ_DATA2 on a.PU_ID equals b.PU_ID.Trim()
                                            where b.SCHEME_NAME == schemeName && a.PAIR == "50"
                                            select new
                                            {
                                                b.PU_QTY,
                                                a.PU_ID,
                                                a.PAIR
                                            };
                             total50pr = 0;
                            if (check50prInc.Count() > 0)
                            {
                                foreach (var a in check50prInc)
                                    total50pr += Convert.ToInt32(a.PU_QTY);

                            }
                            tabDSC.Cells["J16"].Value = total50pr;

                            //100pr
                            int total100pr = 0;
                            var check100pr = from a in ctxdata2.REF_COP_PR
                                            join b in ctxdata2.WV_MAT_DATA2 on a.PU_ID equals b.MAT_ID.Trim()
                                            where b.SCHEME_NAME == schemeName && a.PAIR == "100"
                                            select new
                                            {
                                                b.MAT_QTY,
                                                a.PU_ID,
                                                a.PAIR
                                            };
                            if (check100pr.Count() > 0)
                            {
                                foreach (var a in check100pr)
                                    total100pr += Convert.ToInt32(a.MAT_QTY);

                            }

                             var check100prInc = from a in ctxdata2.REF_COP_PR
                                            join b in ctxdata2.WV_BOQ_DATA2 on a.PU_ID equals b.PU_ID.Trim()
                                            where b.SCHEME_NAME == schemeName && a.PAIR == "100"
                                            select new
                                            {
                                                b.PU_QTY,
                                                a.PU_ID,
                                                a.PAIR
                                            };
                             total100pr = 0;
                            if (check100prInc.Count() > 0)
                            {
                                foreach (var a in check100prInc)
                                    total100pr += Convert.ToInt32(a.PU_QTY);

                            }
                            tabDSC.Cells["J17"].Value = total100pr;

                            //200pr
                            int total200pr = 0;
                            var check200pr = from a in ctxdata2.REF_COP_PR
                                             join b in ctxdata2.WV_MAT_DATA2 on a.PU_ID equals b.MAT_ID.Trim()
                                             where b.SCHEME_NAME == schemeName && a.PAIR == "200"
                                             select new
                                             {
                                                 b.MAT_QTY,
                                                 a.PU_ID,
                                                 a.PAIR
                                             };
                            if (check200pr.Count() > 0)
                            {
                                foreach (var a in check200pr)
                                    total200pr += Convert.ToInt32(a.MAT_QTY);

                            }

                             var check200prInc = from a in ctxdata2.REF_COP_PR
                                             join b in ctxdata2.WV_BOQ_DATA2 on a.PU_ID equals b.PU_ID.Trim()
                                             where b.SCHEME_NAME == schemeName && a.PAIR == "200"
                                             select new
                                             {
                                                 b.PU_QTY,
                                                 a.PU_ID,
                                                 a.PAIR
                                             };
                             total200pr = 0;
                            if (check200prInc.Count() > 0)
                            {
                                foreach (var a in check200prInc)
                                    total200pr += Convert.ToInt32(a.PU_QTY);

                            }
                            tabDSC.Cells["J18"].Value = total200pr;
                        }

                        tabDSC.Cells["K12"].Value = "PRKM";
                        tabDSC.Cells["K13"].Formula = "=(J13*10)/1000";
                        tabDSC.Cells["K14"].Formula = "=(J14*20)/1000";
                        tabDSC.Cells["K15"].Formula = "=(J15*30)/1000";
                        tabDSC.Cells["K16"].Formula = "=(J16*50)/1000";
                        tabDSC.Cells["K17"].Formula = "=(J17*100)/1000";
                        tabDSC.Cells["K18"].Formula = "=(J18*200)/1000";

                        tabDSC.Cells["C20"].Value = "EXPENDITURE TYPE";
                        tabDSC.Cells["R20"].Value = "QTY";
                        tabDSC.Cells["S20"].Value = "UNIT PRICE";
                        tabDSC.Cells["U20"].Value = "COST";
                        tabDSC.Cells["V20"].Value = "REMARKS";
                        tabDSC.Cells["C22"].Value = "1";
                        tabDSC.Cells["C22"].Style.Font.Size = 11;
                        tabDSC.Cells["D22"].Value = "MATERIAL";
                        tabDSC.Cells["D22"].Style.Font.Size = 11;
                        WebService._base.OSPBOQ_MAIN_EXCEL BOQ_REPORT = new WebService._base.OSPBOQ_MAIN_EXCEL();
                        BOQ_REPORT = myWebService.GetOSPBOQ_MAT_Excel2(0, 100, schemeName);
                        int count = BOQ_REPORT.BOQ_MAIN_Excel.Count;
                        tabDSC.Cells["N24"].Value = "UNIT";
                        tabDSC.Cells["N24"].Style.Font.UnderLine = true;
                        tabDSC.Cells["R24"].Value = "QTY";
                        tabDSC.Cells["R24"].Style.Font.UnderLine = true;
                        int j = 25;
                        for (int i = 0; i < count; i++)
                        {
                            tabDSC.Cells["E" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID;
                            tabDSC.Cells["F" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_DESC;
                            tabDSC.Cells["N" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_UOM;
                            tabDSC.Cells["R" + j].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                            tabDSC.Cells["R" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            tabDSC.Cells["Q" + j + ":R" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            tabDSC.Cells["R" + j + ":U" + j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            tabDSC.Cells["R" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_MAT_PRICE;
                            tabDSC.Cells["S" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_MATERIAL_VALUE;
                            tabDSC.Cells["U" + j].Style.Font.Bold = true;
                            tabDSC.Cells["U" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_TOTAL_VALUE;
                            j += 1;
                        }
                        int k = j + 1;
                        k++;
                        
                        j = k + 1;
                        tabDSC.Cells["U" + k + ":U" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabDSC.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabDSC.Cells["U" + j].Style.Font.Bold = true;
                        tabDSC.Cells["U" + j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        double totalMat = 0;
                        if (count > 0)
                            totalMat = Convert.ToDouble(BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_TOTAL_VALUE_TOTAL);
                        else
                            totalMat = Convert.ToDouble(String.Format("{0:0.00}", (0.00)));
                        tabDSC.Cells["U" + j].Value = totalMat;
                        tabDSC.Cells["F11"].Formula = "=U" + j;
                        tabDSC.Cells["V" + j].Style.Font.Bold = true;
                        tabDSC.Cells["V" + j].Value = "Subtotal Materials Cost";

                        j += 2;
                        tabDSC.Cells["B" + j + ":W" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        k = j + 1;
                        tabDSC.Cells["C" + j + ":D" + j + ",Q" + k].Style.Font.Bold = true;
                        tabDSC.Cells["C" + j + ":D" + j].Style.Font.Size = 11;
                        tabDSC.Cells["C" + j].Value = 2;
                        tabDSC.Cells["D" + j].Value = "INCIDENTALS";
                        j = k + 1;
                        tabDSC.Cells["E" + k].Value = "PU NO.";
                        tabDSC.Cells["F" + k].Value = "PU DESCRIPTION";
                        tabDSC.Cells["N" + k + ":P" + k].Merge = true;
                        tabDSC.Cells["R24"].Value = "QTY";
                        BOQ_REPORT = myWebService.GetOSPBOQ_MAIN_Excel2(0, 100, schemeName);
                        count = BOQ_REPORT.BOQ_MAIN_Excel.Count;
                        int count2 = count;

                        for (int i = 0; i < count; i++)
                        {
                            tabDSC.Cells["E" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID;
                            tabDSC.Cells["F" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_DESC;
                            tabDSC.Cells["N" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_UOM;
                            tabDSC.Cells["P" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_RATE_INDICATOR;
                            tabDSC.Cells["R" + j].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                            tabDSC.Cells["R" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            tabDSC.Cells["Q" + j + ":R" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            tabDSC.Cells["R" + j + ":U" + j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            tabDSC.Cells["R" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_QTY;
                            tabDSC.Cells["S" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_INSTALL_PRICE;
                            tabDSC.Cells["U" + j].Style.Font.Bold = true;
                            tabDSC.Cells["U" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_INSTALL_VALUE;
                            
                            j += 1;
                            
                        }
                        k = j + 1;
                        k++;
                        j = k + 2;
                        k++;
                        tabDSC.Cells["D" + j].Style.Font.Bold = true;
                        tabDSC.Cells["D" + j].Value = "TNT : MILEAGE + MEAL + TOLL etc";
                        tabDSC.Cells["I" + j].Value = "AND TnT";
                        tabDSC.Cells["U" + k + ":U" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabDSC.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                        WebService._base.SUMM_ProjCost estimateBP = new WebService._base.SUMM_ProjCost();
                        estimateBP = myWebService.Get_ProjCost(schemeName);
                        count = estimateBP.ProjCost.Count;

                        double SumLbrOT = 0;
                        double SumLbrSalary = 0;
                        double SumTNT_Mileage = 0;
                        double SumTotal = 0;
                        for (int i = 0; i < count; i++)
                        {
                            SumLbrOT = Convert.ToDouble(estimateBP.ProjCost[i].SUM_LBR_OT);
                            SumLbrSalary = Convert.ToDouble(estimateBP.ProjCost[i].SUM_LBR_SALARY);
                            SumTNT_Mileage = Convert.ToDouble(estimateBP.ProjCost[i].SUM_TNT_MILEAGE);
                            SumTotal = Convert.ToDouble(estimateBP.ProjCost[i].SUM_TOTAL);
                        }
                        tabDSC.Cells["U" + j].Style.Font.Bold = true;
                        tabDSC.Cells["U" + j].Value = SumTNT_Mileage;
                        j++;
                        k++;
                        tabDSC.Cells["I" + j].Value = "AD TnT";
                        tabDSC.Cells["U" + k + ":U" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabDSC.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabDSC.Cells["U" + j].Style.Font.Bold = true;

                        double planNImple = 0;
                        if (queryRegion.REGION_NAME_GRN == "CENTRAL")
                            planNImple = 130;
                        else if (queryRegion.REGION_NAME_GRN == "NORTHERN")
                            planNImple = 560;
                        else if (queryRegion.REGION_NAME_GRN == "SOUTHERN")
                            planNImple = 500;
                        else if (queryRegion.REGION_NAME_GRN == "EASTERN")
                            planNImple = 540;
                        else if (queryRegion.REGION_NAME_GRN == "SABAH")
                            planNImple = 910;
                        else if (queryRegion.REGION_NAME_GRN == "SARAWAK")
                            planNImple = 760;

                        double totalInci = 0;
                        if (count2 > 0)
                            totalInci = Convert.ToDouble(BOQ_REPORT.BOQ_MAIN_Excel[count2 - 1].x_INSTALL_VALUE_TOTAL);
                        else
                            totalInci = Convert.ToDouble(String.Format("{0:0.00}", (0.00)));
                        double matPlusInci = totalMat + totalInci;

                        double adMileage = 0;
                        if (matPlusInci <= 25000)
                            adMileage = 100;
                        else
                            adMileage = planNImple;

                        tabDSC.Cells["U" + j].Value = adMileage;
                        j += 2;
                        k = j + 1;
                        tabDSC.Cells["U" + j + ":U" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabDSC.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabDSC.Cells["U" + j].Style.Font.Bold = true;
                        double inciPlusMileage = totalInci + SumTNT_Mileage + adMileage;
                        tabDSC.Cells["U" + j].Value = Convert.ToDouble(String.Format("{0:0.00}", inciPlusMileage));
                        tabDSC.Cells["V" + j].Style.Font.Bold = true;
                        tabDSC.Cells["F12"].Formula = "=U" + j;
                        tabDSC.Cells["V" + j].Value = "Subtotal Incidentals Cost";

                        j += 2;
                        tabDSC.Cells["B" + j + ":W" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabDSC.Cells["C" + j + ":D" + j].Style.Font.Bold = true;
                        tabDSC.Cells["C" + j + ":D" + j].Style.Font.Size = 11;
                        tabDSC.Cells["C" + j].Value = 3;
                        tabDSC.Cells["D" + j].Value = "LABOUR";
                        j++;
                        k = j + 1;
                        tabDSC.Cells["E" + j].Value = "OVERTIME";
                        tabDSC.Cells["G" + j].Value = "AND Staff OT";
                        tabDSC.Cells["U" + j + ":U" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabDSC.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabDSC.Cells["U" + j].Style.Font.Bold = true;
                        tabDSC.Cells["U" + j].Value = SumLbrOT;
                        j++;
                        k = j + 1;
                        tabDSC.Cells["E" + j].Value = "SALARY";
                        tabDSC.Cells["G" + j].Value = "AND Staff Salary";
                        tabDSC.Cells["U" + j + ":U" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabDSC.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabDSC.Cells["U" + j].Style.Font.Bold = true;
                        tabDSC.Cells["U" + j].Value = SumLbrSalary;
                        j++;
                        k = j + 1;
                        tabDSC.Cells["E" + j].Value = "SALARY";
                        tabDSC.Cells["G" + j].Value = "Access Development Staff Salary";
                        tabDSC.Cells["U" + j + ":U" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabDSC.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabDSC.Cells["U" + j].Style.Font.Bold = true;
                        tabDSC.Cells["U" + j].Value = 10;
                        j++;
                        k = j + 1;
                        tabDSC.Cells["U" + j + ":U" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabDSC.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabDSC.Cells["U" + j].Style.Font.Bold = true;
                        tabDSC.Cells["U" + j].Value = SumLbrOT + SumLbrSalary + 10;
                        double totalLabour = SumLbrOT + SumLbrSalary + 10;
                        tabDSC.Cells["V" + j].Style.Font.Bold = true;
                        tabDSC.Cells["V" + j].Value = "Subtotal Labour Cost";
                        tabDSC.Cells["F13"].Formula = "=U" + j;
                        j += 2;
                        k = j + 1;
                        tabDSC.Cells["U" + j + ":U" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabDSC.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabDSC.Cells["U" + j].Style.Font.Bold = true;
                        double grandTot = matPlusInci + totalLabour + SumTNT_Mileage + adMileage;
                        tabDSC.Cells["U" + j].Value = Convert.ToDouble(String.Format("{0:0.00}", grandTot));
                        tabDSC.Cells["V" + j].Style.Font.Bold = true;
                        tabDSC.Cells["V" + j].Value = "GRAND TOTAL COST";
                        j += 2;
                        tabDSC.Cells["B" + j + ":W" + j].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        j++;
                        tabDSC.Cells["D" + j + ":R" + j].Style.Font.Bold = true;
                        tabDSC.Cells["D" + j + ":R" + j].Style.Font.Size = 10;
                        tabDSC.Cells["D" + j].Value = "PREPARED BY :";
                        tabDSC.Cells["I" + j].Value = "CHECKED BY:";
                        tabDSC.Cells["R" + j].Value = "ENDORSED BY:";
                        j += 2;
                        tabDSC.Cells["D" + j + ":F" + j + ",I" + j + ":L" + j + ",R" + j + ":U" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        j += 2;
                        tabDSC.Cells["B" + j + ":W" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

                        var stream = new MemoryStream();
                        xlsFile.SaveAs(stream);

                        string fileName = "BOQReport_" + schemeName + ".xlsx";
                        string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                        stream.Position = 0;
                        return File(stream, contentType, fileName);
                        #endregion
                    }
                    else if (queryProj.SCHEME_TYPE == "Civil")
                    {
                        #region cws
                        var tabCWS = xlsFile.Workbook.Worksheets.Add("CWS");
                        tabCWS.Cells.Style.Font.Name = "Arial";
                        tabCWS.Cells.Style.Font.Size = 8;
                        tabCWS.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        tabCWS.Cells.Style.Fill.BackgroundColor.SetColor(Color.DarkGray);
                        tabCWS.Cells["A1:AD134"].Style.Fill.BackgroundColor.SetColor(Color.White);
                        tabCWS.DefaultRowHeight = 13.5;
                        tabCWS.Column(1).Width = 0.85;
                        tabCWS.Column(2).Width = 1;
                        tabCWS.Column(3).Width = 2.75;
                        tabCWS.Column(4).Width = 2.25;
                        tabCWS.Column(5).Width = 9.25;
                        tabCWS.Column(6).Width = 3.63;
                        for (int i = 7; i < 12; i++)
                            tabCWS.Column(i).Width = 7.75;
                        tabCWS.Column(12).Width = 9.38;
                        tabCWS.Column(13).Width = 5.63;
                        tabCWS.Column(14).Width = 5.75;
                        tabCWS.Column(15).Width = 4.63;
                        tabCWS.Column(16).Width = 3.38;
                        tabCWS.Column(17).Width = 2.25;
                        tabCWS.Column(18).Width = 6;
                        tabCWS.Column(19).Width = 6;
                        tabCWS.Column(20).Width = 1.38;
                        tabCWS.Column(21).Width = 7.5;
                        tabCWS.Column(22).Width = 7.5;
                        tabCWS.Column(23).Width = 8.13;
                        tabCWS.Column(24).Width = 9.75;
                        tabCWS.Column(25).Width = 9.13;
                        tabCWS.Column(26).Width = 0.77;
                        tabCWS.Column(27).Width = 0.85;
                        tabCWS.Column(28).Width = 19.38;
                        tabCWS.Column(29).Width = 1.25;
                        tabCWS.Column(30).Width = 0.69;
                        tabCWS.Row(1).Height = 10.2;
                        tabCWS.Row(2).Height = 15.6;
                        tabCWS.Cells["C2:L7,N4:S7,C8:C10,L11,X7:Y15,C17:Z17,C19:D20"].Style.Font.Bold = true;
                        tabCWS.Cells["C2,AA7"].Style.Font.UnderLine = true;
                        tabCWS.Cells["O2"].Style.Font.Color.SetColor(Color.Red);
                        tabCWS.Cells["C2"].Style.Font.Size = 12;
                        tabCWS.Cells["C19:D19"].Style.Font.Size = 11;
                        tabCWS.Cells["C4:L5,O4,S6"].Style.Font.Size = 10;
                        tabCWS.Cells["R8,Z19"].Style.WrapText = true;
                        tabCWS.Cells["N2,R10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        tabCWS.Cells["R8"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                        tabCWS.Cells["N2,F4:G5,G7:L7,R6,R7:X7,F8:F10,R8:R10,U8:V9,W8:Y15,P14:V15,U17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        tabCWS.Cells["L4,N6:O7,M8:M11,X17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        tabCWS.Cells["G3:H5,O3:AB4,R5:S7,G6:J11,C16:AB16,C18:AB18"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["F4:H5,N4:AB4,F7:J11,Q6:S7"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["G4:H5,O4,R6:S7"].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                        tabCWS.Cells["C17:T18,Z17:AB18"].Style.Fill.BackgroundColor.SetColor(Color.LightPink);
                        tabCWS.Cells["U17:U18,X17:Y18"].Style.Fill.BackgroundColor.SetColor(Color.SandyBrown);
                        tabCWS.Cells["V17:W18"].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                        //tabCWS.Cells["X8:X12"].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                        //tabCWS.Cells["W15"].Style.Fill.BackgroundColor.SetColor(Color.Gray);

                        tabCWS.Cells["C2"].Value = "PROJECT COST BREAKDOWN";
                        tabCWS.Cells["C4"].Value = "PROJECT NO :";
                        tabCWS.Cells["F4:F5,F8:F10"].Value = "=";
                        tabCWS.Cells["G4"].Value = queryProj.PROJECT_NO;
                        tabCWS.Cells["L4"].Value = "PROJECT TITLE:";
                        tabCWS.Cells["O4"].Value = projDesc;
                        tabCWS.Cells["C5"].Value = "REGION :";
                        tabCWS.Cells["G5"].Value = queryRegion.REGION_NAME_GRN;
                        tabCWS.Cells["O6"].Value = "Area:";
                        if (queryRegion.NATIONWIDE_GRP == "WEST")
                            tabCWS.Cells["R6"].Value = "SEM";
                        else if (queryRegion.NATIONWIDE_GRP == "EAST")
                            tabCWS.Cells["R6"].Value = "SBH/SWK";
                        tabCWS.Cells["G7"].Value = "OPEN CUT";
                        tabCWS.Cells["H7"].Value = "HDD";
                        tabCWS.Cells["I7"].Value = "MILL/REIN";
                        //tabCWS.Cells["J7"].Value = "M'TRENC";
                        //tabCWS.Cells["K7"].Value = "MOBIL";
                        tabCWS.Cells["J7"].Value = "TOTAL";
                        tabCWS.Cells["G8"].Value = 0;
                        tabCWS.Cells["H8"].Value = 0;
                        tabCWS.Cells["I8"].Value = 0;
                        tabCWS.Cells["G10"].Value = 0;
                        tabCWS.Cells["H10"].Value = 0;
                        tabCWS.Cells["G11"].Formula = "=SUM(G8:G10)";
                        tabCWS.Cells["H11"].Formula = "=SUM(H8:H10)";
                        tabCWS.Cells["I11"].Formula = "=SUM(I8:I10)";
                        tabCWS.Cells["J11"].Formula = "=SUM(J8:J10)";
                        //tabCWS.Cells["K11"].Formula = "=SUM(K8:K10)";
                        //tabCWS.Cells["L11"].Formula = "=SUM(L8:L10)";

                        tabCWS.Cells["N7"].Value = "Location Factor:";
                        if (queryRegion.NATIONWIDE_GRP == "WEST")
                            tabCWS.Cells["R7"].Value = 1.00;
                        else if (queryRegion.NATIONWIDE_GRP == "EAST")
                            tabCWS.Cells["R7"].Value = 1.15;
                        //tabCWS.Cells["U7,U14"].Value = "Km";
                        //tabCWS.Cells["V7"].Value = "Duct Way";
                        //tabCWS.Cells["W7"].Value = "Total Cost";
                        //tabCWS.Cells["X7"].Value = "Dt_Km";
                        //tabCWS.Cells["Y7"].Value = "Cost Dt_Km";
                        tabCWS.Cells["C8"].Value = "MATERIALS COST";
                        tabCWS.Cells["C9"].Value = "INCIDENTAL COST";
                        tabCWS.Cells["C10"].Value = "LABOUR COST";
                        //tabCWS.Cells["R8"].Value = "OPEN CUT";
                        //tabCWS.Cells["R10"].Value = "HDD";
                        //tabCWS.Cells["S10"].Value = "2W";
                        //tabCWS.Cells["S11"].Value = "4W";
                        //tabCWS.Cells["S12"].Value = "6W";
                        //tabCWS.Cells["T8,T10,C19"].Value = 1;
                        //tabCWS.Cells["T9,T11,V10"].Value = 2;
                        //tabCWS.Cells["T12"].Value = 3;
                        //tabCWS.Cells["V11"].Value = 4;
                        //tabCWS.Cells["V12"].Value = 6;
                        //tabCWS.Cells["P15"].Value = "ROAD REINSTATE";
                        //tabCWS.Cells["V14"].Value = "Lane";
                        tabCWS.Cells["C17"].Value = "EXPENDITURE TYPE";
                        tabCWS.Cells["U17"].Value = "QTY";
                        tabCWS.Cells["V17"].Value = "UNIT PRICE";
                        tabCWS.Cells["X17"].Value = "COST";
                        tabCWS.Cells["Z17"].Value = "REMARKS";
                        tabCWS.Cells["C19"].Value = "1";
                        tabCWS.Cells["D19"].Value = "MATERIAL";

                        WebService._base.OSPBOQ_MAIN_EXCEL BOQ_REPORT = new WebService._base.OSPBOQ_MAIN_EXCEL();
                        BOQ_REPORT = myWebService.GetOSPBOQ_MAT_Excel2(0, 100, schemeName);


                        int count = BOQ_REPORT.BOQ_MAIN_Excel.Count;
                        int j = 21;
                        for (int i = 0; i < count; i++)
                        {
                            //tabCWS.Cells["A" + j].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                            tabCWS.Cells["E" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID;
                            tabCWS.Cells["F" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_DESC;
                            tabCWS.Cells["P" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_UOM;
                            tabCWS.Cells["U" + j].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                            tabCWS.Cells["U" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            tabCWS.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            tabCWS.Cells["U" + j + ":X" + j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            tabCWS.Cells["U" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_MAT_PRICE;
                            tabCWS.Cells["V" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_MATERIAL_VALUE;
                            tabCWS.Cells["X" + j].Style.Font.Bold = true;
                            tabCWS.Cells["X" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_TOTAL_VALUE;
                            j += 1;
                        }
                        int k = j + 1;
                        tabCWS.Cells["E" + j + ":E" + k + ",U" + j + ":U" + k].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                        tabCWS.Cells["D" + j + ":E" + k + ",T" + j + ":U" + k].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["X" + j + ":X" + k].Style.Font.Bold = true;
                        tabCWS.Cells["X" + j + ":X" + k].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        tabCWS.Cells["X" + j].Value = "0.00";
                        tabCWS.Cells["X" + k].Value = "0.00";
                        k++;
                        tabCWS.Cells["E" + j + ":E" + k + ",U" + j + ":U" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        j = k + 1;
                        tabCWS.Cells["X" + k + ":X" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["W" + j + ":X" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["X" + j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        tabCWS.Cells["X" + j].Style.Font.Bold = true;
                        double totalMat = 0;
                        if (count > 0)
                            totalMat = Convert.ToDouble(BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_TOTAL_VALUE_TOTAL);
                        else
                            totalMat = Convert.ToDouble(String.Format("{0:0.00}", (0.00)));
                        tabCWS.Cells["X" + j].Value = totalMat;

                        tabCWS.Cells["J8"].Formula = "=X" + j;

                        tabCWS.Cells["Z" + j].Style.Font.Bold = true;
                        tabCWS.Cells["Z" + j].Value = "Subtotal Materials Cost";

                        j += 2;
                        tabCWS.Cells["B" + j + ":AD" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        k = j + 1;
                        tabCWS.Cells["C" + j + ":D" + j + ",Q" + k].Style.Font.Bold = true;
                        tabCWS.Cells["C" + j + ":D" + j].Style.Font.Size = 11;
                        tabCWS.Cells["C" + j].Value = 2;
                        tabCWS.Cells["D" + j].Value = "INCIDENTALS";
                        j = k + 1;

                        BOQ_REPORT = myWebService.GetOSPBOQ_MAIN_Excel2(0, 100, schemeName);
                        count = BOQ_REPORT.BOQ_MAIN_Excel.Count;
                        int count2 = count;
                        System.Diagnostics.Debug.WriteLine("====OPEN CUT====");
                        j++;
                        tabCWS.Cells["E" + j].Value = "OPEN CUT";
                        tabCWS.Cells["E" + j].Style.Font.Bold = true;
                        double totoc1 = 0;
                        for (int p = 0; p < count; p++)
                        {
                            int puid = Convert.ToInt32(BOQ_REPORT.BOQ_MAIN_Excel[p].x_PU_ID);
                            var queryOPENCUT = (from a in ctxdata.WV_FEAT_MAST
                                                join b in ctxdata.REF_CIV_DUCTPATH on a.MIN_MATERIAL equals b.MIN_MATERIAL
                                                where (a.DAY == puid || a.HOLIDAY == puid || a.NIGHT == puid || a.WEEKEND == puid)
                                                select b.DT_S_PLACMNT).FirstOrDefault();
                            if (queryOPENCUT != "HDD")
                            {
                                j++;
                                tabCWS.Cells["E" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[p].x_PU_ID;
                                tabCWS.Cells["F" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[p].x_PU_DESC;
                                tabCWS.Cells["P" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[p].x_PU_UOM;
                                tabCWS.Cells["S" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[p].x_RATE_INDICATOR;
                                tabCWS.Cells["U" + j].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                                tabCWS.Cells["U" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                tabCWS.Cells["U" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                tabCWS.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                tabCWS.Cells["U" + j + ":X" + j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                tabCWS.Cells["U" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[p].x_PU_QTY;
                                tabCWS.Cells["V" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[p].x_BQ_INSTALL_PRICE;
                                tabCWS.Cells["X" + j].Style.Font.Bold = true;
                                tabCWS.Cells["X" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[p].x_INSTALL_VALUE;

                                double totoc = Convert.ToDouble(tabCWS.Cells["X" + j].Value);
                                totoc1 = totoc1 + totoc;
                                tabCWS.Cells["G9"].Value = totoc1;
                                System.Diagnostics.Debug.WriteLine(totoc1);
                            }

                        }
                        j++;
                        tabCWS.Cells["E" + j].Value = "HDD";
                        tabCWS.Cells["E" + j].Style.Font.Bold = true;
                        System.Diagnostics.Debug.WriteLine("===HDD===");
                        double tothdd1 = 0;
                        for (int q = 0; q < count; q++)
                        {
                            int puid = Convert.ToInt32(BOQ_REPORT.BOQ_MAIN_Excel[q].x_PU_ID);
                            var queryhdd = (from a in ctxdata.WV_FEAT_MAST
                                            join b in ctxdata.REF_CIV_DUCTPATH on a.MIN_MATERIAL equals b.MIN_MATERIAL
                                            where (a.DAY == puid || a.HOLIDAY == puid || a.NIGHT == puid || a.WEEKEND == puid)
                                            select b.DT_S_PLACMNT).FirstOrDefault();

                            if (queryhdd == "HDD")
                            {
                                j++;
                                tabCWS.Cells["E" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[q].x_PU_ID;
                                tabCWS.Cells["F" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[q].x_PU_DESC;
                                tabCWS.Cells["P" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[q].x_PU_UOM;
                                tabCWS.Cells["S" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[q].x_RATE_INDICATOR;
                                tabCWS.Cells["U" + j].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                                tabCWS.Cells["U" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                tabCWS.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                tabCWS.Cells["U" + j + ":X" + j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                tabCWS.Cells["U" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[q].x_PU_QTY;
                                tabCWS.Cells["V" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[q].x_BQ_INSTALL_PRICE;
                                tabCWS.Cells["X" + j].Style.Font.Bold = true;
                                tabCWS.Cells["X" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[q].x_INSTALL_VALUE;

                                double tothdd = Convert.ToDouble(tabCWS.Cells["X" + j].Value);
                                tothdd1 = tothdd1 + tothdd;
                                tabCWS.Cells["H9"].Value = tothdd1;
                            }
                        }
                        j++;

                        k = j + 1;
                        tabCWS.Cells["E" + j + ":E" + k + ",S" + j + ":S" + k + ",U" + j + ":U" + k].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                        tabCWS.Cells["D" + j + ":E" + k + ",R" + j + ":S" + k + ",T" + j + ":U" + k].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["X" + j + ":X" + k].Style.Font.Bold = true;
                        tabCWS.Cells["X" + j + ":X" + k].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        tabCWS.Cells["X" + j].Value = "0.00";
                        tabCWS.Cells["X" + k].Value = "0.00";
                        k++;
                        tabCWS.Cells["E" + j + ":E" + k + ",S" + j + ":S" + k + ",U" + j + ":U" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        j = k + 2;
                        k++;
                        tabCWS.Cells["D" + j].Style.Font.Bold = true;
                        tabCWS.Cells["D" + j].Value = "TNT : MILEAGE + MEAL + TOLL etc";
                        tabCWS.Cells["I" + j].Value = "AND TnT";
                        tabCWS.Cells["X" + k + ":X" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["W" + j + ":X" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                        WebService._base.SUMM_ProjCost estimateBP = new WebService._base.SUMM_ProjCost();
                        estimateBP = myWebService.Get_ProjCost(schemeName);
                        count = estimateBP.ProjCost.Count;

                        double SumLbrOT = 0;
                        double SumLbrSalary = 0;
                        double SumTNT_Mileage = 0;
                        double SumTotal = 0;
                        for (int i = 0; i < count; i++)
                        {
                            SumLbrOT = Convert.ToDouble(estimateBP.ProjCost[i].SUM_LBR_OT);
                            SumLbrSalary = Convert.ToDouble(estimateBP.ProjCost[i].SUM_LBR_SALARY);
                            SumTNT_Mileage = Convert.ToDouble(estimateBP.ProjCost[i].SUM_TNT_MILEAGE);
                            SumTotal = Convert.ToDouble(estimateBP.ProjCost[i].SUM_TOTAL);
                        }
                        tabCWS.Cells["X" + j].Style.Font.Bold = true;
                        tabCWS.Cells["X" + j].Value = SumTNT_Mileage;
                        j++;
                        k++;
                        tabCWS.Cells["I" + j].Value = "AD TnT";
                        tabCWS.Cells["X" + k + ":X" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["W" + j + ":X" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["X" + j].Style.Font.Bold = true;

                        double planNImple = 0;
                        if (queryRegion.REGION_NAME_GRN == "CENTRAL")
                            planNImple = 130;
                        else if (queryRegion.REGION_NAME_GRN == "NORTHERN")
                            planNImple = 560;
                        else if (queryRegion.REGION_NAME_GRN == "SOUTHERN")
                            planNImple = 500;
                        else if (queryRegion.REGION_NAME_GRN == "EASTERN")
                            planNImple = 540;
                        else if (queryRegion.REGION_NAME_GRN == "SABAH")
                            planNImple = 910;
                        else if (queryRegion.REGION_NAME_GRN == "SARAWAK")
                            planNImple = 760;

                        double totalInci = 0;
                        if (count2 > 0)
                            totalInci = Convert.ToDouble(BOQ_REPORT.BOQ_MAIN_Excel[count2 - 1].x_INSTALL_VALUE_TOTAL);
                        else
                            totalInci = Convert.ToDouble(String.Format("{0:0.00}", (0.00)));
                        double matPlusInci = totalMat + totalInci;

                        double adMileage = 0;
                        if (matPlusInci <= 25000)
                            adMileage = 100;
                        else
                            adMileage = planNImple;

                        tabCWS.Cells["X" + j].Value = adMileage;
                        double tottnt = adMileage + SumTNT_Mileage;
                        tabCWS.Cells["I9"].Value = tottnt;
                        j += 2;
                        k = j + 1;
                        tabCWS.Cells["X" + j + ":X" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["W" + j + ":X" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["X" + j].Style.Font.Bold = true;
                        double inciPlusMileage = totalInci + SumTNT_Mileage + adMileage;
                        tabCWS.Cells["X" + j].Value = Convert.ToDouble(String.Format("{0:0.00}", inciPlusMileage));

                        tabCWS.Cells["J9"].Formula = "=X" + j;

                        tabCWS.Cells["Z" + j].Style.Font.Bold = true;
                        tabCWS.Cells["Z" + j].Value = "Subtotal Incidentals Cost";

                        j += 2;
                        tabCWS.Cells["B" + j + ":AD" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["C" + j + ":D" + j].Style.Font.Bold = true;
                        tabCWS.Cells["C" + j + ":D" + j].Style.Font.Size = 11;
                        tabCWS.Cells["C" + j].Value = 3;
                        tabCWS.Cells["D" + j].Value = "LABOUR";
                        j++;
                        k = j + 1;
                        tabCWS.Cells["E" + j].Value = "OVERTIME";
                        tabCWS.Cells["G" + j].Value = "AND Staff OT";
                        tabCWS.Cells["X" + j + ":X" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["W" + j + ":X" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["X" + j].Style.Font.Bold = true;
                        tabCWS.Cells["X" + j].Value = SumLbrOT;
                        j++;
                        k = j + 1;
                        tabCWS.Cells["E" + j].Value = "SALARY";
                        tabCWS.Cells["G" + j].Value = "AND Staff Salary";
                        tabCWS.Cells["X" + j + ":X" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["W" + j + ":X" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["X" + j].Style.Font.Bold = true;
                        tabCWS.Cells["X" + j].Value = SumLbrSalary;
                        j++;
                        k = j + 1;
                        tabCWS.Cells["E" + j].Value = "SALARY";
                        tabCWS.Cells["G" + j].Value = "Access Development Staff Salary";
                        tabCWS.Cells["X" + j + ":X" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["W" + j + ":X" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["X" + j].Style.Font.Bold = true;
                        tabCWS.Cells["X" + j].Value = 10;
                        j++;
                        k = j + 1;
                        tabCWS.Cells["X" + j + ":X" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["W" + j + ":X" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["X" + j].Style.Font.Bold = true;
                        tabCWS.Cells["X" + j].Value = SumLbrOT + SumLbrSalary + 10;
                        double totalLabour = SumLbrOT + SumLbrSalary + 10;
                        tabCWS.Cells["I10"].Value = totalLabour;
                        tabCWS.Cells["J10"].Formula = "=X" + j;

                        tabCWS.Cells["Z" + j].Style.Font.Bold = true;
                        tabCWS.Cells["Z" + j].Value = "Subtotal Labour Cost";
                        j += 2;
                        k = j + 1;
                        tabCWS.Cells["X" + j + ":X" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["W" + j + ":X" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabCWS.Cells["X" + j].Style.Font.Bold = true;
                        double grandTot = matPlusInci + totalLabour + SumTNT_Mileage + adMileage;
                        tabCWS.Cells["X" + j].Value = Convert.ToDouble(String.Format("{0:0.00}", grandTot));
                        tabCWS.Cells["Z" + j].Style.Font.Bold = true;
                        tabCWS.Cells["Z" + j].Value = "GRAND TOTAL COST";
                        j += 2;
                        tabCWS.Cells["B" + j + ":AD" + j].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        j++;
                        tabCWS.Cells["D" + j + ":W" + j].Style.Font.Bold = true;
                        tabCWS.Cells["D" + j + ":W" + j].Style.Font.Size = 10;
                        tabCWS.Cells["D" + j].Value = "PREPARED BY :";
                        tabCWS.Cells["L" + j].Value = "CHECKED BY:";
                        tabCWS.Cells["W" + j].Value = "ENDORSED BY:";
                        j += 2;
                        tabCWS.Cells["D" + j + ":F" + j + ",L" + j + ":O" + j + ",W" + j + ":Z" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        j += 2;
                        tabCWS.Cells["B" + j + ":AD" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

                        tabCWS.Cells["G4:H4,L4:N4,O4:AB4,G5:H5,O6:P6,R6:S6,N7:P7,R7:S7"].Merge = true;
                        tabCWS.Cells["R8:R9,R10:R12,AA7:AC7,AA8:AC8,AA9:AC11,P15:T15,Z19:AB20"].Merge = true;

                        var stream = new MemoryStream();
                        xlsFile.SaveAs(stream);

                        string fileName = "BOQReport_" + schemeName + ".xlsx";
                        string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                        stream.Position = 0;
                        return File(stream, contentType, fileName);
                        #endregion
                    }

                    else if (queryFnoVdsl > 0)
                    {
                        #region vdsl
                        var tabVDSL = xlsFile.Workbook.Worksheets.Add("StdCosting");
                        tabVDSL.Cells.Style.Font.Name = "Arial";
                        tabVDSL.Cells.Style.Font.Size = 8;
                        tabVDSL.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        tabVDSL.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                        tabVDSL.Cells["A1:W228"].Style.Fill.BackgroundColor.SetColor(Color.White);
                        tabVDSL.Cells["A1:A228,W1:W228"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        tabVDSL.DefaultRowHeight = 13.2;
                        tabVDSL.Column(1).Width = 0.75; //A
                        tabVDSL.Column(2).Width = 0.63; //B
                        tabVDSL.Column(3).Width = 3.11; //C
                        tabVDSL.Column(4).Width = 3.11; //D
                        for (int i = 5; i < 9; i++)//E F G H
                            tabVDSL.Column(i).Width = 12.33;
                        tabVDSL.Column(9).Width = 3.89; //I
                        tabVDSL.Column(10).Width = 8.56; //J
                        for (int i = 11; i < 17; i++)//K L M N O P
                            tabVDSL.Column(i).Width = 6.78;
                        for (int i = 17; i < 21; i++)//Q R S T
                            tabVDSL.Column(i).Width = 8.89;
                        tabVDSL.Column(21).Width = 11.11; //U
                        tabVDSL.Column(22).Width = 27.78; //V
                        tabVDSL.Column(23).Width = 0.81; //W

                        tabVDSL.Row(1).Height = 4.2;
                        tabVDSL.Row(2).Height = 30;
                        tabVDSL.Row(7).Height = 4.4;
                        tabVDSL.Cells["A1:W11,C12:C14,F15:I15,L12:U14,C25:D25"].Style.Font.Bold = true;
                        tabVDSL.Cells["C3:F3,U8"].Style.Font.UnderLine = true;
                        tabVDSL.Cells["M3"].Style.Font.Color.SetColor(Color.Red);
                        tabVDSL.Cells["C2"].Style.Font.Size = 24;
                        tabVDSL.Cells["C3,C25:D25"].Style.Font.Size = 10;
                        tabVDSL.Cells["M8:M14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        tabVDSL.Cells["F4:O4,F5:O5,F6:F9,N7:N14,L2:L3,B22:W22,B24:W24"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin; //KOTAK PROJECT + SUPPLIER
                        tabVDSL.Cells["E5:E6,G5,I5,O5,F6,E8:E9,F8:F9,M8:M14,N8:N14,K3:L3"].Style.Border.Right.Style = ExcelBorderStyle.Thin; //KOTAK PROJECT + SUPPLIER

                        tabVDSL.Cells["F11:J11,F12:J12,F13:J13,F14:J14,F15:J15,U7:V7,U13:V13"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin; //KOTAK MATERIAL, BENCHMARKING
                        tabVDSL.Cells["E12:E15,F12:F15,G12:G15,H12:H15,J12:J15,T8:T13,V8:V13"].Style.Border.Right.Style = ExcelBorderStyle.Thin; //KOTAK MATERIAL, BENCHMARKING

                        tabVDSL.Cells["M15:N15,M16:N16,M17:N17,M18:N18,M19:N19,M20:N20,R14:T14,R15:T15,R16:T16,R17:T17,R18:T18,R19:T19,R20:T20,R21:T21"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin; //KOTAK FIBER, COPPER
                        tabVDSL.Cells["L16:L20,M16:M20,N16:N20,Q15:Q21,R15:R21,S15:S21,T15:T21"].Style.Border.Right.Style = ExcelBorderStyle.Thin; //KOTAK FIBER, COPPER

                        tabVDSL.Cells["L3,F5:O5,F6,F8,F9,N8:N14"].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                        tabVDSL.Cells["C23:Q24,V23:V24"].Style.Fill.BackgroundColor.SetColor(Color.LightPink);
                        tabVDSL.Cells["R23:R24,U23:U24"].Style.Fill.BackgroundColor.SetColor(Color.SandyBrown);
                        tabVDSL.Cells["S23:T24"].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);

                        tabVDSL.Cells["C2"].Value = "STANDARD COSTING FOR MSAN VDSL2";
                        tabVDSL.Cells["C3"].Value = "PROJECT COST BREAKDOWN";
                        tabVDSL.Cells["L3"].Value = "X";
                        tabVDSL.Cells["M3"].Value = "Input in this blue box only!!";
                        tabVDSL.Cells["C5"].Value = "PROJECT NO :";
                        tabVDSL.Cells["F5"].Value = queryProj.PROJECT_NO;
                        tabVDSL.Cells["H5"].Value = "PROJECT TITLE:";
                        tabVDSL.Cells["J5"].Value = projDesc;
                        tabVDSL.Cells["C6"].Value = "REGION :";
                        tabVDSL.Cells["F6"].Value = queryRegion.REGION_NAME_GRN;
                        tabVDSL.Cells["C8"].Value = "HOMEPASS";
                        tabVDSL.Cells["U8"].Value = "COSTING BENCHMARKING";
                        tabVDSL.Cells["C9"].Value = "COST/HOMEPASS";

                        tabVDSL.Cells["F9"].Formula = "=J15/F8";

                        tabVDSL.Cells["F11"].Value = "MSAN VDSL2";
                        tabVDSL.Cells["G11"].Value = "DSS";
                        tabVDSL.Cells["H11"].Value = "ESIDE";
                        tabVDSL.Cells["I11"].Value = "TOTAL";

                        tabVDSL.Cells["F12"].Formula = "=0";
                        tabVDSL.Cells["G14"].Formula = "=0";
                        tabVDSL.Cells["H14"].Formula = "=0";
                        tabVDSL.Cells["J12"].Formula = "=SUM(F12:H12)";
                        tabVDSL.Cells["J13"].Formula = "=SUM(F13:H13)";
                        tabVDSL.Cells["J14"].Formula = "=SUM(F14:H14)";
                        tabVDSL.Cells["F15"].Formula = "=SUM(F12:F14)";
                        tabVDSL.Cells["G15"].Formula = "=SUM(G12:G14)";
                        tabVDSL.Cells["H15"].Formula = "=SUM(H12:H14)";
                        tabVDSL.Cells["J15"].Formula = "=SUM(J12:J14)";

                        tabVDSL.Cells["C12"].Value = "MATERIAL COST";
                        tabVDSL.Cells["C13"].Value = "INCIDENTAL COST";
                        tabVDSL.Cells["C14"].Value = "LABOUR COST";
                        tabVDSL.Cells["M8"].Value = "SUPPLIER";

                        using (Entities1 ctxdata2 = new Entities1())
                        {
                            var supp = (from a in ctxdata2.GC_VDSL2_2
                                        join b in ctxdata2.GC_NETELEM2 on a.G3E_FID equals b.G3E_FID
                                        where b.SCHEME_NAME == schemeName
                                        select a.MANUFACTURER).FirstOrDefault();
                            tabVDSL.Cells["N8"].Value = supp;

                            var fiberSpliceManufac = (from a in ctxdata2.GC_NETELEM2
                                                      join b in ctxdata2.GC_FSPLICE on a.G3E_FID equals b.G3E_FID
                                                      where a.SCHEME_NAME == schemeName
                                                      select b.MANUFACTURER).FirstOrDefault();
                            tabVDSL.Cells["N12"].Value = fiberSpliceManufac;
                        }

                        tabVDSL.Cells["M9"].Value = "CABLE FOC";

                        var checkFiberEside = (from a in ctxdata.GC_NETELEM
                                               where a.SCHEME_NAME == schemeName && a.G3E_FNO == 7200
                                               select a).Count();
                        if (checkFiberEside > 0)
                        {
                            var fiberEsideManufac = (from a in ctxdata.GC_NETELEM
                                                     join b in ctxdata.GC_FCBL on a.G3E_FID equals b.G3E_FID
                                                     where a.SCHEME_NAME == schemeName && a.G3E_FNO == 7200
                                                     select b.MANUFACTURER).FirstOrDefault();
                            tabVDSL.Cells["N9"].Value = fiberEsideManufac;
                        }
                        tabVDSL.Cells["M10"].Value = "LOCATION FACTOR";
                        if (queryRegion.NATIONWIDE_GRP == "WEST")
                            tabVDSL.Cells["N10"].Value = 1.00;
                        else if (queryRegion.NATIONWIDE_GRP == "EAST")
                            tabVDSL.Cells["N10"].Value = 1.15;
                        tabVDSL.Cells["M11"].Value = "SEM / SBH&SWK";
                        if (queryRegion.NATIONWIDE_GRP == "WEST")
                            tabVDSL.Cells["N11"].Value = "SEM";
                        else if (queryRegion.NATIONWIDE_GRP == "EAST")
                            tabVDSL.Cells["N11"].Value = "SBH/SWK";
                        tabVDSL.Cells["M12"].Value = "CLOSURE FOC";

                        tabVDSL.Cells["M13"].Value = "INDOOR/OUTDOOR";
                        tabVDSL.Cells["M14"].Value = "SYSTEM/TOP UP";
                        tabVDSL.Cells["M16"].Value = "FIBER";
                        tabVDSL.Cells["N16"].Value = "FKM";

                        tabVDSL.Cells["M17"].Value = "12";
                        tabVDSL.Cells["M18"].Value = "24";
                        tabVDSL.Cells["M19"].Value = "36";
                        tabVDSL.Cells["M20"].Value = "48";
                        tabVDSL.Cells["N17"].Value = 0;
                        tabVDSL.Cells["N18"].Value = 0;
                        tabVDSL.Cells["N19"].Value = 0;
                        tabVDSL.Cells["N20"].Value = 0;

                        tabVDSL.Cells["R15"].Value = "COPPER";
                        tabVDSL.Cells["S15"].Value = "Dist (m)";
                        tabVDSL.Cells["T15"].Value = "PRKM";

                        tabVDSL.Cells["R16"].Value = "10pr";
                        tabVDSL.Cells["R17"].Value = "20pr";
                        tabVDSL.Cells["R18"].Value = "30pr";
                        tabVDSL.Cells["R19"].Value = "50pr";
                        tabVDSL.Cells["R20"].Value = "100pr";
                        tabVDSL.Cells["R21"].Value = "200pr";
                        tabVDSL.Cells["S16"].Value = 0;
                        tabVDSL.Cells["S17"].Value = 0;
                        tabVDSL.Cells["S18"].Value = 0;
                        tabVDSL.Cells["S19"].Value = 0;
                        tabVDSL.Cells["S20"].Value = 0;
                        tabVDSL.Cells["S21"].Value = 0;
                        tabVDSL.Cells["T16"].Formula = "=(S16*10)/1000";
                        tabVDSL.Cells["T17"].Formula = "=(S17*10)/1000";
                        tabVDSL.Cells["T18"].Formula = "=(S18*10)/1000";
                        tabVDSL.Cells["T19"].Formula = "=(S19*10)/1000";
                        tabVDSL.Cells["T20"].Formula = "=(S20*10)/1000";
                        tabVDSL.Cells["T21"].Formula = "=(S21*10)/1000";

                        tabVDSL.Cells["C23"].Value = "EXPENDITURE TYPE";
                        tabVDSL.Cells["R23"].Value = "QTY";
                        tabVDSL.Cells["S23"].Value = "UNIT PRICE";
                        tabVDSL.Cells["U23"].Value = "COST";
                        tabVDSL.Cells["V23"].Value = "REMARKS";
                        tabVDSL.Cells["C25"].Value = "1";
                        tabVDSL.Cells["C25"].Style.Font.Size = 11;
                        tabVDSL.Cells["D25"].Value = "MATERIAL";
                        tabVDSL.Cells["D25"].Style.Font.Size = 11;

                        WebService._base.OSPBOQ_MAIN_EXCEL BOQ_REPORT = new WebService._base.OSPBOQ_MAIN_EXCEL();
                        BOQ_REPORT = myWebService.GetOSPBOQ_MAT_Excel2(0, 100, schemeName);
                        int count = BOQ_REPORT.BOQ_MAIN_Excel.Count;
                        int j = 27;
                        for (int i = 0; i < count; i++)
                        {
                            tabVDSL.Cells["E" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID;
                            tabVDSL.Cells["F" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_DESC;
                            tabVDSL.Cells["N" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_UOM;
                            tabVDSL.Cells["R" + j].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                            tabVDSL.Cells["R" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            tabVDSL.Cells["Q" + j + ":R" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            tabVDSL.Cells["R" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_MAT_PRICE;
                            tabVDSL.Cells["S" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_MATERIAL_VALUE;
                            tabVDSL.Cells["U" + j].Style.Font.Bold = true;
                            tabVDSL.Cells["U" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_TOTAL_VALUE;
                            if (BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID == "1000000190")
                                tabVDSL.Cells["N17"].Formula = "=" + BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_MAT_PRICE + "*(1/1000)";
                            else if (BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID == "1000000708")
                                tabVDSL.Cells["N18"].Formula = "=" + BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_MAT_PRICE + "*(1/1000)";
                            else if (BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID == "1000000709")
                                tabVDSL.Cells["N19"].Formula = "=" + BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_MAT_PRICE + "*(1/1000)";
                            else if (BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID == "1000000710")
                                tabVDSL.Cells["N20"].Formula = "=" + BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_MAT_PRICE + "*(1/1000)";
                            else if (BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID == "1000000139")
                                tabVDSL.Cells["S16"].Formula = "=" + BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_MAT_PRICE;
                            else if (BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID == "1000000140")
                                tabVDSL.Cells["S17"].Formula = "=" + BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_MAT_PRICE;
                            else if (BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID == "1000000141")
                                tabVDSL.Cells["S18"].Formula = "=" + BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_MAT_PRICE;
                            else if (BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID == "1000000142")
                                tabVDSL.Cells["S19"].Formula = "=" + BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_MAT_PRICE;
                            else if (BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID == "1000000143")
                                tabVDSL.Cells["S20"].Formula = "=" + BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_MAT_PRICE;
                            else if (BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID == "1000000144")
                                tabVDSL.Cells["S21"].Formula = "=" + BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_MAT_PRICE;
                            j += 1;
                        }
                        int k = j + 1;
                        tabVDSL.Cells["E" + j + ":E" + k + ",R" + j + ":R" + k].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                        tabVDSL.Cells["D" + j + ":E" + k + ",Q" + j + ":R" + k].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabVDSL.Cells["U" + j].Value = "0.00";
                        tabVDSL.Cells["U" + k].Value = "0.00";
                        k++;
                        tabVDSL.Cells["E" + j + ":E" + k + ",R" + j + ":R" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        j = k + 1;
                        tabVDSL.Cells["U" + k + ":U" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabVDSL.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabVDSL.Cells["U" + j].Style.Font.Bold = true;
                        double totalMat = 0;
                        if (count > 0)
                            totalMat = Convert.ToDouble(BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_TOTAL_VALUE_TOTAL);
                        else
                            totalMat = Convert.ToDouble(String.Format("{0:0.00}", (0.00)));
                        tabVDSL.Cells["U" + j].Value = totalMat;

                        tabVDSL.Cells["V" + j].Style.Font.Bold = true;
                        tabVDSL.Cells["V" + j].Value = "Subtotal Materials Cost";

                        j += 2;
                        tabVDSL.Cells["B" + j + ":W" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        k = j + 1;
                        tabVDSL.Cells["C" + j + ":D" + j].Style.Font.Bold = true;
                        tabVDSL.Cells["C" + j + ":D" + j].Style.Font.Size = 11;
                        tabVDSL.Cells["C" + j].Value = 2;
                        tabVDSL.Cells["D" + j].Value = "INCIDENTALS";
                        j = k + 1;

                        BOQ_REPORT = myWebService.GetOSPBOQ_MAIN_Excel2(0, 100, schemeName);
                        count = BOQ_REPORT.BOQ_MAIN_Excel.Count;
                        int count2 = count;

                        double totalContract = 0;
                        double totalDside = 0;
                        double totalEside = 0;
                        for (int i = 0; i < count; i++)
                        {
                            if (BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONTRACT_NO == null || BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONTRACT_NO == "")
                            {
                                tabVDSL.Cells["E" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID;
                                tabVDSL.Cells["F" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_DESC;
                                tabVDSL.Cells["N" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_UOM;
                                tabVDSL.Cells["Q" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_RATE_INDICATOR;
                                tabVDSL.Cells["R" + j].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                                tabVDSL.Cells["R" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                tabVDSL.Cells["Q" + j + ":R" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                tabVDSL.Cells["R" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_QTY;
                                tabVDSL.Cells["S" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_INSTALL_PRICE;
                                tabVDSL.Cells["U" + j].Style.Font.Bold = true;
                                tabVDSL.Cells["U" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_INSTALL_VALUE;
                                decimal numPuId = Convert.ToDecimal(BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID);
                                string checkLookupTable = "";
                                if (BOQ_REPORT.BOQ_MAIN_Excel[i].x_RATE_INDICATOR == "D")
                                {
                                    var checkFsplice = (from a in ctxdata.WV_FEAT_MAST
                                                        where a.DAY == numPuId
                                                        select a).Count();
                                    if (checkFsplice > 0)
                                        checkLookupTable = (from a in ctxdata.WV_FEAT_MAST
                                                            where a.DAY == numPuId
                                                            select a.LOOKUP_TABLE).FirstOrDefault();
                                }
                                else if (BOQ_REPORT.BOQ_MAIN_Excel[i].x_RATE_INDICATOR == "N")
                                {
                                    var checkFsplice = (from a in ctxdata.WV_FEAT_MAST
                                                        where a.NIGHT == numPuId
                                                        select a).Count();
                                    if (checkFsplice > 0)
                                        checkLookupTable = (from a in ctxdata.WV_FEAT_MAST
                                                            where a.NIGHT == numPuId
                                                            select a.LOOKUP_TABLE).FirstOrDefault();
                                }
                                else if (BOQ_REPORT.BOQ_MAIN_Excel[i].x_RATE_INDICATOR == "W")
                                {
                                    var checkFsplice = (from a in ctxdata.WV_FEAT_MAST
                                                        where a.WEEKEND == numPuId
                                                        select a).Count();
                                    if (checkFsplice > 0)
                                        checkLookupTable = (from a in ctxdata.WV_FEAT_MAST
                                                            where a.WEEKEND == numPuId
                                                            select a.LOOKUP_TABLE).FirstOrDefault();
                                }
                                else
                                {
                                    var checkFsplice = (from a in ctxdata.WV_FEAT_MAST
                                                        where a.HOLIDAY == numPuId
                                                        select a).Count();
                                    if (checkFsplice > 0)
                                        checkLookupTable = (from a in ctxdata.WV_FEAT_MAST
                                                            where a.HOLIDAY == numPuId
                                                            select a.LOOKUP_TABLE).FirstOrDefault();
                                }
                                if (checkLookupTable == "REF_FDC" && BOQ_REPORT.BOQ_MAIN_Excel[i].x_INSTALL_VALUE != "-")
                                    totalEside += Convert.ToDouble(BOQ_REPORT.BOQ_MAIN_Excel[i].x_INSTALL_VALUE);
                                else
                                    totalDside += Convert.ToDouble(BOQ_REPORT.BOQ_MAIN_Excel[i].x_INSTALL_VALUE);
                            }
                            else
                            {
                                tabVDSL.Cells["E" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONTRACT_NO;
                                tabVDSL.Cells["F" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_ITEM_NO;
                                tabVDSL.Cells["G" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONTRACT_DESC;
                                tabVDSL.Cells["N" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_UOM;
                                tabVDSL.Cells["R" + j].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                                tabVDSL.Cells["R" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                tabVDSL.Cells["Q" + j + ":R" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                tabVDSL.Cells["R" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_QTY;
                                tabVDSL.Cells["S" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_CONT_PRICE;
                                tabVDSL.Cells["U" + j].Style.Font.Bold = true;
                                tabVDSL.Cells["U" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONT_VALUE;

                                if (BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONT_VALUE != "-")
                                    totalContract += Convert.ToDouble(BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONT_VALUE);
                            }
                            j += 1;
                        }
                        tabVDSL.Cells["G13"].Value = totalDside;
                        tabVDSL.Cells["H13"].Value = totalEside;
                        double dsideMat = 0;
                        double esideMat = 0;
                        if (totalDside == 0 && totalEside == 0)
                        {
                            dsideMat = totalMat * 0.5;
                            esideMat = totalMat * 0.5;
                        }
                        else
                        {
                            dsideMat = totalDside * totalMat / (totalDside + totalEside);
                            esideMat = totalEside * totalMat / (totalDside + totalEside);
                        }
                        tabVDSL.Cells["G12"].Value = Convert.ToDouble(String.Format("{0:0.00}", dsideMat));
                        tabVDSL.Cells["H12"].Value = Convert.ToDouble(String.Format("{0:0.00}", esideMat));
                        k = j + 1;
                        tabVDSL.Cells["E" + j + ":E" + k + ",R" + j + ":R" + k].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                        tabVDSL.Cells["D" + j + ":E" + k + ",Q" + j + ":R" + k].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabVDSL.Cells["U" + j].Value = "0.00";
                        tabVDSL.Cells["U" + k].Value = "0.00";
                        k++;
                        tabVDSL.Cells["E" + j + ":E" + k + ",R" + j + ":R" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        j = k + 2;
                        k++;
                        tabVDSL.Cells["D" + j].Style.Font.Bold = true;
                        tabVDSL.Cells["D" + j].Value = "TNT : MILEAGE + MEAL + TOLL etc";
                        tabVDSL.Cells["H" + j].Value = "AND TnT";
                        tabVDSL.Cells["U" + k + ":U" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabVDSL.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                        WebService._base.SUMM_ProjCost estimateBP = new WebService._base.SUMM_ProjCost();
                        estimateBP = myWebService.Get_ProjCost(schemeName);
                        count = estimateBP.ProjCost.Count;

                        double SumLbrOT = 0;
                        double SumLbrSalary = 0;
                        double SumTNT_Mileage = 0;
                        double SumTotal = 0;
                        for (int i = 0; i < count; i++)
                        {
                            SumLbrOT = Convert.ToDouble(estimateBP.ProjCost[i].SUM_LBR_OT);
                            SumLbrSalary = Convert.ToDouble(estimateBP.ProjCost[i].SUM_LBR_SALARY);
                            SumTNT_Mileage = Convert.ToDouble(estimateBP.ProjCost[i].SUM_TNT_MILEAGE);
                            SumTotal = Convert.ToDouble(estimateBP.ProjCost[i].SUM_TOTAL);
                        }
                        tabVDSL.Cells["U" + j].Style.Font.Bold = true;
                        tabVDSL.Cells["U" + j].Value = SumTNT_Mileage;
                        j++;
                        k++;
                        tabVDSL.Cells["H" + j].Value = "AD TnT";
                        tabVDSL.Cells["U" + k + ":U" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabVDSL.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabVDSL.Cells["U" + j].Style.Font.Bold = true;

                        double planNImple = 0;
                        if (queryRegion.REGION_NAME_GRN == "CENTRAL")
                            planNImple = 130;
                        else if (queryRegion.REGION_NAME_GRN == "NORTHERN")
                            planNImple = 560;
                        else if (queryRegion.REGION_NAME_GRN == "SOUTHERN")
                            planNImple = 500;
                        else if (queryRegion.REGION_NAME_GRN == "EASTERN")
                            planNImple = 540;
                        else if (queryRegion.REGION_NAME_GRN == "SABAH")
                            planNImple = 910;
                        else if (queryRegion.REGION_NAME_GRN == "SARAWAK")
                            planNImple = 760;

                        double totalInci = 0;
                        if (count2 > 0)
                            totalInci = Convert.ToDouble(BOQ_REPORT.BOQ_MAIN_Excel[count2 - 1].x_INSTALL_VALUE_TOTAL);
                        else
                            totalInci = Convert.ToDouble(String.Format("{0:0.00}", (0.00)));
                        double matPlusInci = totalMat + totalInci;

                        double adMileage = 0;
                        if (matPlusInci <= 25000)
                            adMileage = 100;
                        else
                            adMileage = planNImple;

                        tabVDSL.Cells["U" + j].Value = adMileage;
                        j += 2;
                        k = j + 1;
                        tabVDSL.Cells["U" + j + ":U" + k].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabVDSL.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabVDSL.Cells["U" + j].Style.Font.Bold = true;
                        double inciPlusMileage = totalInci + SumTNT_Mileage + adMileage;
                        tabVDSL.Cells["U" + j].Value = Convert.ToDouble(String.Format("{0:0.00}", inciPlusMileage));
                        tabVDSL.Cells["F13"].Value = totalContract + SumTNT_Mileage + adMileage;

                        tabVDSL.Cells["V" + j].Style.Font.Bold = true;
                        tabVDSL.Cells["V" + j].Value = "Subtotal Incidentals Cost";

                        j += 2;
                        tabVDSL.Cells["B" + j + ":W" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabVDSL.Cells["C" + j + ":D" + j].Style.Font.Bold = true;
                        tabVDSL.Cells["C" + j + ":D" + j].Style.Font.Size = 11;
                        tabVDSL.Cells["C" + j].Value = 3;
                        tabVDSL.Cells["D" + j].Value = "LABOUR";
                        j++;
                        k = j + 1;
                        tabVDSL.Cells["E" + j].Value = "OVERTIME";
                        tabVDSL.Cells["G" + j].Value = "AND Staff OT";
                        tabVDSL.Cells["U" + j + ":U" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabVDSL.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabVDSL.Cells["U" + j].Style.Font.Bold = true;
                        tabVDSL.Cells["U" + j].Value = SumLbrOT;
                        j++;
                        k = j + 1;
                        tabVDSL.Cells["E" + j].Value = "SALARY";
                        tabVDSL.Cells["G" + j].Value = "AND Staff Salary";
                        tabVDSL.Cells["U" + j + ":U" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabVDSL.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabVDSL.Cells["U" + j].Style.Font.Bold = true;
                        tabVDSL.Cells["U" + j].Value = SumLbrSalary;
                        j++;
                        k = j + 1;
                        tabVDSL.Cells["E" + j].Value = "SALARY";
                        tabVDSL.Cells["G" + j].Value = "Access Development Staff Salary";
                        tabVDSL.Cells["U" + j + ":U" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabVDSL.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabVDSL.Cells["U" + j].Style.Font.Bold = true;
                        tabVDSL.Cells["U" + j].Value = 10;
                        j++;
                        k = j + 1;
                        tabVDSL.Cells["U" + j + ":U" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabVDSL.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabVDSL.Cells["U" + j].Style.Font.Bold = true;
                        tabVDSL.Cells["U" + j].Value = SumLbrOT + SumLbrSalary + 10;
                        double totalLabour = SumLbrOT + SumLbrSalary + 10;

                        tabVDSL.Cells["F14"].Formula = "=U" + j;

                        tabVDSL.Cells["V" + j].Style.Font.Bold = true;
                        tabVDSL.Cells["V" + j].Value = "Subtotal Labour Cost";
                        j += 2;
                        k = j + 1;
                        tabVDSL.Cells["U" + j + ":U" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabVDSL.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabVDSL.Cells["U" + j].Style.Font.Bold = true;
                        double grandTot = matPlusInci + totalLabour + SumTNT_Mileage + adMileage;
                        tabVDSL.Cells["U" + j].Value = Convert.ToDouble(String.Format("{0:0.00}", grandTot));
                        tabVDSL.Cells["V" + j].Style.Font.Bold = true;
                        tabVDSL.Cells["V" + j].Value = "GRAND TOTAL COST";
                        j += 2;
                        tabVDSL.Cells["B" + j + ":W" + j].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        j++;
                        tabVDSL.Cells["D" + j + ":S" + j].Style.Font.Bold = true;
                        tabVDSL.Cells["D" + j + ":S" + j].Style.Font.Size = 10;
                        tabVDSL.Cells["D" + j].Value = "PREPARED BY :";
                        tabVDSL.Cells["I" + j].Value = "CHECKED BY:";
                        tabVDSL.Cells["R" + j].Value = "ENDORSED BY:";
                        j += 2;
                        tabVDSL.Cells["D" + j + ":F" + j + ",I" + j + ":L" + j + ",R" + j + ":U" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        j += 2;
                        tabVDSL.Cells["B" + j + ":W" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

                        var stream = new MemoryStream();
                        xlsFile.SaveAs(stream);

                        string fileName = "BOQReport_" + schemeName + ".xlsx";
                        string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                        stream.Position = 0;
                        return File(stream, contentType, fileName);
                        #endregion
                    }
                    else if (queryFnoMsan > 0)
                    {
                        #region msan
                        var tabMSAN = xlsFile.Workbook.Worksheets.Add("MSAN");
                        tabMSAN.Cells.Style.Font.Name = "Arial";
                        tabMSAN.Cells.Style.Font.Size = 10;
                        tabMSAN.Cells.Style.Numberformat.Format = "0.00";
                        tabMSAN.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        tabMSAN.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                        tabMSAN.Cells["A1:W228"].Style.Fill.BackgroundColor.SetColor(Color.White);
                        tabMSAN.Cells["A1:A228,W1:W228"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        tabMSAN.DefaultRowHeight = 13.2;
                        tabMSAN.Column(1).Width = 0.75; //A
                        tabMSAN.Column(2).Width = 0.63; //B
                        tabMSAN.Column(3).Width = 3.11; //C
                        tabMSAN.Column(4).Width = 3.11; //D
                        for (int i = 5; i < 9; i++)//E F G H
                            tabMSAN.Column(i).Width = 13.11;
                        tabMSAN.Column(9).Width = 6.78; //I
                        tabMSAN.Column(10).Width = 10.67; //J
                        for (int i = 11; i < 17; i++)//K L M N O P
                            tabMSAN.Column(i).Width = 6.78;
                        for (int i = 17; i < 21; i++)//Q R S T
                            tabMSAN.Column(i).Width = 8.89;
                        tabMSAN.Column(21).Width = 11.11; //U
                        tabMSAN.Column(22).Width = 30.44; //V
                        tabMSAN.Column(23).Width = 0.81; //W

                        tabMSAN.Row(1).Height = 4.2;
                        tabMSAN.Row(2).Height = 30;
                        tabMSAN.Row(1).Height = 7;
                        tabMSAN.Cells["A1:W11,C12:C14,F15:H15,J12:W16,A18:W19,C20:D20"].Style.Font.Bold = true;
                        tabMSAN.Cells["C3:F3,U8"].Style.Font.UnderLine = true;
                        tabMSAN.Cells["L3:M3"].Style.Font.Color.SetColor(Color.Red);
                        tabMSAN.Cells["C2"].Style.Font.Size = 24;
                        tabMSAN.Cells["C3"].Style.Font.Size = 12;
                        tabMSAN.Cells["C3,C20:D20"].Style.Font.Size = 10;
                        tabMSAN.Cells["L3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        tabMSAN.Cells["L3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        tabMSAN.Cells["N10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        tabMSAN.Cells["M8:M14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        tabMSAN.Cells["F4:O4,F5:O5,F6:F9,N7:N14,L2:L3,C17:V17,C19:V19"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin; //KOTAK PROJECT + SUPPLIER
                        tabMSAN.Cells["E5:E6,G5,I5,O5,F6,E8:E9,F8:F9,M8:M14,N8:N14,K3:L3"].Style.Border.Right.Style = ExcelBorderStyle.Thin; //KOTAK PROJECT + SUPPLIER

                        tabMSAN.Cells["F11:H11,F12:H12,F13:H13,F14:H14,F15:H15,U7:V7,U13:V13"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin; //KOTAK MATERIAL, BENCHMARKING
                        tabMSAN.Cells["E12:E15,F12:F15,G12:G15,H12:H15,T8:T13,V8:V13"].Style.Border.Right.Style = ExcelBorderStyle.Thin; //KOTAK MATERIAL, BENCHMARKING

                        tabMSAN.Cells["L3,F5:O5,F6,F8,F9,N8:N14"].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                        tabMSAN.Cells["C18:Q19,V18:V19"].Style.Fill.BackgroundColor.SetColor(Color.LightPink);
                        tabMSAN.Cells["R18:R19,U18:U19"].Style.Fill.BackgroundColor.SetColor(Color.SandyBrown);
                        tabMSAN.Cells["S18:T19"].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);

                        tabMSAN.Cells["C2"].Value = "STANDARD COSTING FOR MSAN";
                        tabMSAN.Cells["C3"].Value = "PROJECT COST BREAKDOWN";
                        tabMSAN.Cells["L3"].Value = "X";
                        tabMSAN.Cells["M3"].Value = "Input in this blue box only!!";
                        tabMSAN.Cells["C5"].Value = "PROJECT NO :";
                        tabMSAN.Cells["F5"].Value = queryProj.PROJECT_NO;
                        tabMSAN.Cells["H5"].Value = "PROJECT TITLE:";
                        tabMSAN.Cells["J5"].Value = projDesc;
                        tabMSAN.Cells["C6"].Value = "REGION :";
                        tabMSAN.Cells["F6"].Value = queryRegion.REGION_NAME_GRN;
                        tabMSAN.Cells["C8"].Value = "VFC";
                        tabMSAN.Cells["U8"].Value = "COSTING BENCHMARKING";
                        tabMSAN.Cells["C9"].Value = "PORTSA";
                        tabMSAN.Cells["F11"].Value = "IPMSAN";
                        tabMSAN.Cells["G11"].Value = "POWER/OTH";
                        tabMSAN.Cells["H11"].Value = "TOTAL";
                        tabMSAN.Cells["C12"].Value = "MATERIAL COST";
                        tabMSAN.Cells["C13"].Value = "INCIDENTAL COST";
                        tabMSAN.Cells["C14"].Value = "LABOUR COST";

                        tabMSAN.Cells["F12"].Formula = "=0";
                        tabMSAN.Cells["G12"].Formula = "=0";
                        tabMSAN.Cells["G14"].Formula = "=0";
                        tabMSAN.Cells["H12"].Formula = "=SUM(F12:G12)";
                        tabMSAN.Cells["H13"].Formula = "=SUM(F13:G13)";
                        tabMSAN.Cells["H14"].Formula = "=SUM(F14:G14)";
                        tabMSAN.Cells["F15"].Formula = "=SUM(F12:F14)";
                        tabMSAN.Cells["G15"].Formula = "=SUM(G12:G14)";
                        tabMSAN.Cells["H15"].Formula = "=SUM(H12:H14)";

                        tabMSAN.Cells["M8"].Value = "SUPPLIER";

                        var supp = (from a in ctxdata.GC_MSAN
                                    join b in ctxdata.GC_NETELEM on a.G3E_FID equals b.G3E_FID
                                    where b.SCHEME_NAME == schemeName && b.G3E_FNO == 9100
                                    select a.MANUFACTURER).FirstOrDefault();
                        tabMSAN.Cells["N8"].Value = supp;

                        tabMSAN.Cells["M9"].Value = "CABLE FOC";

                        var checkFiberEside = (from a in ctxdata.GC_NETELEM
                                               where a.SCHEME_NAME == schemeName && a.G3E_FNO == 7200
                                               select a).Count();
                        if (checkFiberEside > 0)
                        {
                            var fiberEsideManufac = (from a in ctxdata.GC_NETELEM
                                                     join b in ctxdata.GC_FCBL on a.G3E_FID equals b.G3E_FID
                                                     where a.SCHEME_NAME == schemeName && a.G3E_FNO == 7200
                                                     select b.MANUFACTURER).FirstOrDefault();
                            tabMSAN.Cells["N9"].Value = fiberEsideManufac;
                        }

                        tabMSAN.Cells["M10"].Value = "LOCATION FACTOR";
                        if (queryRegion.NATIONWIDE_GRP == "WEST")
                            tabMSAN.Cells["N10"].Value = 1.00;
                        else if (queryRegion.NATIONWIDE_GRP == "EAST")
                            tabMSAN.Cells["N10"].Value = 1.15;
                        tabMSAN.Cells["M11"].Value = "SEM / SBH&SWK";
                        if (queryRegion.NATIONWIDE_GRP == "WEST")
                            tabMSAN.Cells["N11"].Value = "SEM";
                        else if (queryRegion.NATIONWIDE_GRP == "EAST")
                            tabMSAN.Cells["N11"].Value = "SBH/SWK";
                        tabMSAN.Cells["M12"].Value = "CLOSURE FOC";

                        using (Entities1 ctxdata2 = new Entities1())
                        {
                            var fiberSpliceManufac = (from a in ctxdata2.GC_NETELEM2
                                                      join b in ctxdata2.GC_FSPLICE on a.G3E_FID equals b.G3E_FID
                                                      where a.SCHEME_NAME == schemeName
                                                      select b.MANUFACTURER).FirstOrDefault();
                            tabMSAN.Cells["N12"].Value = fiberSpliceManufac;
                        }

                        tabMSAN.Cells["M13"].Value = "INDOOR/OUTDOOR";
                        tabMSAN.Cells["M14"].Value = "SYSTEM/TOP UP";
                        tabMSAN.Cells["C18"].Value = "EXPENDITURE TYPE";
                        tabMSAN.Cells["R18"].Value = "QTY";
                        tabMSAN.Cells["S18"].Value = "UNIT PRICE";
                        tabMSAN.Cells["U18"].Value = "COST";
                        tabMSAN.Cells["V18"].Value = "REMARKS";
                        tabMSAN.Cells["C20"].Value = "1";
                        tabMSAN.Cells["C20"].Style.Font.Size = 11;
                        tabMSAN.Cells["D20"].Value = "MATERIAL";
                        tabMSAN.Cells["D20"].Style.Font.Size = 11;

                        WebService._base.OSPBOQ_MAIN_EXCEL BOQ_REPORT = new WebService._base.OSPBOQ_MAIN_EXCEL();
                        BOQ_REPORT = myWebService.GetOSPBOQ_MAT_Excel2(0, 100, schemeName);
                        int count = BOQ_REPORT.BOQ_MAIN_Excel.Count;
                        int j = 21;
                        for (int i = 0; i < count; i++)
                        {
                            tabMSAN.Cells["E" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID;
                            tabMSAN.Cells["F" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_DESC;
                            tabMSAN.Cells["N" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_UOM;
                            tabMSAN.Cells["R" + j].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                            tabMSAN.Cells["R" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            tabMSAN.Cells["Q" + j + ":R" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            tabMSAN.Cells["R" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_MAT_PRICE;
                            tabMSAN.Cells["S" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_MATERIAL_VALUE;
                            tabMSAN.Cells["U" + j].Style.Font.Bold = true;
                            tabMSAN.Cells["U" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_TOTAL_VALUE;
                            j += 1;
                        }
                        j++;
                        int k = j + 1;
                        tabMSAN.Cells["E" + j + ":E" + k + ",R" + j + ":R" + k].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                        tabMSAN.Cells["D" + j + ":E" + k + ",Q" + j + ":R" + k].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabMSAN.Cells["U" + j].Value = "0.00";
                        tabMSAN.Cells["U" + k].Value = "0.00";
                        k++;
                        tabMSAN.Cells["E" + j + ":E" + k + ",R" + j + ":R" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        j = k + 1;
                        tabMSAN.Cells["U" + k + ":U" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabMSAN.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabMSAN.Cells["U" + j].Style.Font.Bold = true;
                        double totalMat = 0;
                        if (count > 0)
                            totalMat = Convert.ToDouble(BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_TOTAL_VALUE_TOTAL);
                        else
                            totalMat = Convert.ToDouble(String.Format("{0:0.00}", (0.00)));
                        tabMSAN.Cells["U" + j].Value = totalMat;

                        tabMSAN.Cells["V" + j].Style.Font.Bold = true;
                        tabMSAN.Cells["V" + j].Value = "Subtotal Materials Cost";

                        j += 2;
                        tabMSAN.Cells["B" + j + ":W" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        k = j + 1;
                        tabMSAN.Cells["C" + j + ":D" + j].Style.Font.Bold = true;
                        tabMSAN.Cells["C" + j + ":D" + j].Style.Font.Size = 11;
                        tabMSAN.Cells["C" + j].Value = 2;
                        tabMSAN.Cells["D" + j].Value = "INCIDENTALS";
                        j = k + 1;

                        //var puNContract = from a in ctxdata.WV_BOQ_DATA
                        //                  where a.SCHEME_NAME == schemeName
                        //                  orderby a.ITEM_NO, a.PU_ID, a.CONTRACT_NO
                        //                  select new { a.ITEM_NO, a.PU_ID, a.CONTRACT_NO };
                        //string listpu = "";
                        //string listcontract = "";
                        //foreach (var a in puNContract)
                        //{
                        //    if (a.CONTRACT_NO == null || a.CONTRACT_NO == "")
                        //        listpu += a.ITEM_NO + "|";
                        //    else
                        //        listcontract += a.CONTRACT_NO + "," + a.ITEM_NO + "|";
                        //}

                        BOQ_REPORT = myWebService.GetOSPBOQ_MAIN_Excel2(0, 100, schemeName);
                        count = BOQ_REPORT.BOQ_MAIN_Excel.Count;
                        int count2 = count;

                        double msanContract = 0;
                        double powerContract = 0;
                        for (int i = 0; i < count; i++)
                        {
                            if (BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONTRACT_NO == null || BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONTRACT_NO == "")
                            {
                                tabMSAN.Cells["E" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID;
                                tabMSAN.Cells["F" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_DESC;
                                tabMSAN.Cells["N" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_UOM;
                                tabMSAN.Cells["P" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_RATE_INDICATOR;
                                tabMSAN.Cells["R" + j].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                                tabMSAN.Cells["R" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                tabMSAN.Cells["Q" + j + ":R" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                tabMSAN.Cells["R" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_QTY;
                                tabMSAN.Cells["S" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_INSTALL_PRICE;
                                tabMSAN.Cells["U" + j].Style.Font.Bold = true;
                                tabMSAN.Cells["U" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_INSTALL_VALUE;
                            }
                            else
                            {
                                tabMSAN.Cells["E" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONTRACT_NO;
                                tabMSAN.Cells["F" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_ITEM_NO;
                                tabMSAN.Cells["G" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONTRACT_DESC;
                                tabMSAN.Cells["N" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_UOM;
                                tabMSAN.Cells["R" + j].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                                tabMSAN.Cells["R" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                tabMSAN.Cells["Q" + j + ":R" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                tabMSAN.Cells["R" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_QTY;
                                tabMSAN.Cells["S" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_CONT_PRICE;
                                tabMSAN.Cells["U" + j].Style.Font.Bold = true;
                                tabMSAN.Cells["U" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONT_VALUE;
                                if ((BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONTRACT_NO == "3400002212" || BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONTRACT_NO == "3400002213") && BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONT_VALUE != "-")
                                    msanContract += Convert.ToDouble(BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONT_VALUE);
                                else if (BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONT_VALUE != "-")//if (BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONTRACT_NO == "3400001024" || BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONTRACT_NO == "3400001025" || BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONTRACT_NO == "3400001026" || BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONTRACT_NO == "3400001029")
                                    powerContract += Convert.ToDouble(BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONT_VALUE);
                            }
                            j += 1;
                        }
                        tabMSAN.Cells["G13"].Value = powerContract;
                        k = j + 1;
                        tabMSAN.Cells["E" + j + ":E" + k + ",R" + j + ":R" + k].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                        tabMSAN.Cells["D" + j + ":E" + k + ",Q" + j + ":R" + k].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabMSAN.Cells["U" + j].Value = "0.00";
                        tabMSAN.Cells["U" + k].Value = "0.00";
                        k++;
                        tabMSAN.Cells["E" + j + ":E" + k + ",R" + j + ":R" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        j = k + 2;
                        k++;
                        tabMSAN.Cells["D" + j].Style.Font.Bold = true;
                        tabMSAN.Cells["D" + j].Value = "TNT : MILEAGE + MEAL + TOLL etc";
                        tabMSAN.Cells["I" + j].Value = "AND TnT";
                        tabMSAN.Cells["U" + k + ":U" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabMSAN.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                        WebService._base.SUMM_ProjCost estimateBP = new WebService._base.SUMM_ProjCost();
                        estimateBP = myWebService.Get_ProjCost(schemeName);
                        count = estimateBP.ProjCost.Count;

                        double SumLbrOT = 0;
                        double SumLbrSalary = 0;
                        double SumTNT_Mileage = 0;
                        double SumTotal = 0;
                        for (int i = 0; i < count; i++)
                        {
                            SumLbrOT = Convert.ToDouble(estimateBP.ProjCost[i].SUM_LBR_OT);
                            SumLbrSalary = Convert.ToDouble(estimateBP.ProjCost[i].SUM_LBR_SALARY);
                            SumTNT_Mileage = Convert.ToDouble(estimateBP.ProjCost[i].SUM_TNT_MILEAGE);
                            SumTotal = Convert.ToDouble(estimateBP.ProjCost[i].SUM_TOTAL);
                        }
                        tabMSAN.Cells["U" + j].Style.Font.Bold = true;
                        tabMSAN.Cells["U" + j].Value = SumTNT_Mileage;
                        j++;
                        k++;
                        tabMSAN.Cells["I" + j].Value = "AD TnT";
                        tabMSAN.Cells["U" + k + ":U" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabMSAN.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabMSAN.Cells["U" + j].Style.Font.Bold = true;

                        double planNImple = 0;
                        if (queryRegion.REGION_NAME_GRN == "CENTRAL")
                            planNImple = 130;
                        else if (queryRegion.REGION_NAME_GRN == "NORTHERN")
                            planNImple = 560;
                        else if (queryRegion.REGION_NAME_GRN == "SOUTHERN")
                            planNImple = 500;
                        else if (queryRegion.REGION_NAME_GRN == "EASTERN")
                            planNImple = 540;
                        else if (queryRegion.REGION_NAME_GRN == "SABAH")
                            planNImple = 910;
                        else if (queryRegion.REGION_NAME_GRN == "SARAWAK")
                            planNImple = 760;

                        double totalInci = 0;
                        if (count2 > 0)
                            totalInci = Convert.ToDouble(BOQ_REPORT.BOQ_MAIN_Excel[count2 - 1].x_CONT_VALUE_TOTAL) + Convert.ToDouble(BOQ_REPORT.BOQ_MAIN_Excel[count2 - 1].x_INSTALL_VALUE_TOTAL);
                        else
                            totalInci = Convert.ToDouble(String.Format("{0:0.00}", (0.00)));
                        double matPlusInci = totalMat + totalInci;

                        double adMileage = 0;
                        if (matPlusInci <= 25000)
                            adMileage = 100;
                        else
                            adMileage = planNImple;

                        tabMSAN.Cells["U" + j].Value = adMileage;
                        j += 2;
                        k = j + 1;
                        tabMSAN.Cells["U" + j + ":U" + k].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabMSAN.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabMSAN.Cells["U" + j].Style.Font.Bold = true;
                        double inciPlusMileage = totalInci + SumTNT_Mileage + adMileage;
                        tabMSAN.Cells["U" + j].Value = Convert.ToDouble(String.Format("{0:0.00}", inciPlusMileage));
                        tabMSAN.Cells["F13"].Value = msanContract + SumTNT_Mileage + adMileage;

                        tabMSAN.Cells["V" + j].Style.Font.Bold = true;
                        tabMSAN.Cells["V" + j].Value = "Subtotal Incidentals Cost";

                        j += 2;
                        tabMSAN.Cells["B" + j + ":W" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabMSAN.Cells["C" + j + ":D" + j].Style.Font.Bold = true;
                        tabMSAN.Cells["C" + j + ":D" + j].Style.Font.Size = 11;
                        tabMSAN.Cells["C" + j].Value = 3;
                        tabMSAN.Cells["D" + j].Value = "LABOUR";
                        j++;
                        k = j + 1;
                        tabMSAN.Cells["E" + j].Value = "OVERTIME";
                        tabMSAN.Cells["G" + j].Value = "AND Staff OT";
                        tabMSAN.Cells["U" + j + ":U" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabMSAN.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabMSAN.Cells["U" + j].Style.Font.Bold = true;
                        tabMSAN.Cells["U" + j].Value = SumLbrOT;
                        j++;
                        k = j + 1;
                        tabMSAN.Cells["E" + j].Value = "SALARY";
                        tabMSAN.Cells["G" + j].Value = "AND Staff Salary";
                        tabMSAN.Cells["U" + j + ":U" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabMSAN.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabMSAN.Cells["U" + j].Style.Font.Bold = true;
                        tabMSAN.Cells["U" + j].Value = SumLbrSalary;
                        j++;
                        k = j + 1;
                        tabMSAN.Cells["E" + j].Value = "SALARY";
                        tabMSAN.Cells["G" + j].Value = "Access Development Staff Salary";
                        tabMSAN.Cells["U" + j + ":U" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabMSAN.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabMSAN.Cells["U" + j].Style.Font.Bold = true;
                        tabMSAN.Cells["U" + j].Value = 10;
                        j++;
                        k = j + 1;
                        tabMSAN.Cells["U" + j + ":U" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabMSAN.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabMSAN.Cells["U" + j].Style.Font.Bold = true;
                        tabMSAN.Cells["U" + j].Value = SumLbrOT + SumLbrSalary + 10;
                        double totalLabour = SumLbrOT + SumLbrSalary + 10;

                        tabMSAN.Cells["F14"].Formula = "=U" + j;

                        tabMSAN.Cells["V" + j].Style.Font.Bold = true;
                        tabMSAN.Cells["V" + j].Value = "Subtotal Labour Cost";
                        j += 2;
                        k = j + 1;
                        tabMSAN.Cells["U" + j + ":U" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabMSAN.Cells["T" + j + ":U" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabMSAN.Cells["U" + j].Style.Font.Bold = true;
                        double grandTot = matPlusInci + totalLabour + SumTNT_Mileage + adMileage;
                        tabMSAN.Cells["U" + j].Value = Convert.ToDouble(String.Format("{0:0.00}", grandTot));
                        tabMSAN.Cells["V" + j].Style.Font.Bold = true;
                        tabMSAN.Cells["V" + j].Value = "GRAND TOTAL COST";
                        j += 2;
                        tabMSAN.Cells["B" + j + ":W" + j].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        j++;
                        tabMSAN.Cells["D" + j + ":S" + j].Style.Font.Bold = true;
                        tabMSAN.Cells["D" + j + ":S" + j].Style.Font.Size = 10;
                        tabMSAN.Cells["D" + j].Value = "PREPARED BY :";
                        tabMSAN.Cells["I" + j].Value = "CHECKED BY:";
                        tabMSAN.Cells["R" + j].Value = "ENDORSED BY:";
                        j += 2;
                        tabMSAN.Cells["D" + j + ":F" + j + ",I" + j + ":L" + j + ",R" + j + ":U" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        j += 2;
                        tabMSAN.Cells["B" + j + ":W" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

                        var stream = new MemoryStream();
                        xlsFile.SaveAs(stream);

                        string fileName = "BOQReport_" + schemeName + ".xlsx";
                        string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                        stream.Position = 0;
                        return File(stream, contentType, fileName);
                        #endregion
                    }
                    else if (queryftth > 0 && jkhContract == "JKH")
                    {
                        #region FTTH
                        /*var vendkabel = (from a in ctxdata.GC_NETELEM
                                         join b in ctxdata.GC_FCBL on a.G3E_FID equals b.G3E_FID
                                         where a.SCHEME_NAME == schemeName
                                         select b.CONTRACTOR).FirstOrDefault();

                        var vendfdp = (from a in ctxdata.GC_NETELEM
                                       join b in ctxdata.GC_FDP on a.G3E_FID equals b.G3E_FID
                                       where a.SCHEME_NAME == schemeName
                                       select b.CONTRACTOR).FirstOrDefault();

                        var vendfdc = (from a in ctxdata.GC_NETELEM
                                       join b in ctxdata.GC_FDC on a.G3E_FID equals b.G3E_FID
                                       where a.SCHEME_NAME == schemeName
                                       select b.CONTRACTOR).FirstOrDefault();*/

                        var tabFTTH = xlsFile.Workbook.Worksheets.Add("StdCosting MC");
                        tabFTTH.Cells.Style.Font.Name = "Arial";
                        tabFTTH.Cells.Style.Font.Size = 8;
                        tabFTTH.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        tabFTTH.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                        tabFTTH.Cells["A1:Z328"].Style.Fill.BackgroundColor.SetColor(Color.White);
                        tabFTTH.Cells["A1:A328,Z1:Z328"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        tabFTTH.DefaultRowHeight = 13.2;
                        tabFTTH.Column(1).Width = 0.75; //A
                        tabFTTH.Column(2).Width = 0.63; //B
                        tabFTTH.Column(3).Width = 3.11; //C
                        tabFTTH.Column(4).Width = 3.11; //D
                        tabFTTH.Column(5).Width = 14.11; //E
                        for (int i = 6; i < 9; i++)
                            tabFTTH.Column(i).Width = 13.11; //F G H 
                        for (int i = 9; i < 13; i++)
                            tabFTTH.Column(i).Width = 6.78; //I J K L M N 
                        tabFTTH.Column(13).Width = 8.00; // M 
                        tabFTTH.Column(14).Width = 6.78; // N
                        tabFTTH.Column(15).Width = 3.67; //O
                        tabFTTH.Column(16).Width = 3.67; //P
                        tabFTTH.Column(17).Width = 1.33; //Q
                        tabFTTH.Column(18).Width = 6.78; //R
                        for (int i = 19; i < 23; i++)
                            tabFTTH.Column(i).Width = 8.89; //S T U V
                        tabFTTH.Column(23).Width = 11.11;//W
                        tabFTTH.Column(24).Width = 41.89; //X
                        tabFTTH.Column(25).Width = 0.81; //Y
                        tabFTTH.Column(26).Width = 0.81;//Z

                        tabFTTH.Row(1).Height = 4.2;
                        tabFTTH.Row(2).Height = 21;
                        tabFTTH.Row(7).Height = 4.2;
                        tabFTTH.Cells["C2:X10,C11:C13,F14:H14,L11:X16,C18:X18,C20:D20"].Style.Font.Bold = true;
                        tabFTTH.Cells["C3,F3"].Style.Font.UnderLine = true;
                        tabFTTH.Cells["L3:M3"].Style.Font.Color.SetColor(Color.Red);
                        tabFTTH.Cells["C2"].Style.Font.Size = 16;
                        tabFTTH.Cells["C3:F3"].Style.Font.Size = 12;
                        tabFTTH.Cells["L3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        tabFTTH.Cells["I5,L8:L16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        tabFTTH.Cells["L3,M8:M16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        tabFTTH.Cells["F4:G4,F5:G5,F6:F9,J4:R4,J5:R5,H7:H8,W7:X7,W16:X16,C17:X17,C19:X19"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin; //PROJECT HPASS COSTING
                        tabFTTH.Cells["E5:E6,G5,F6,E8:E9,F8:F9,G8:H8,V8:V16,X8:X16,I5,R5"].Style.Border.Right.Style = ExcelBorderStyle.Thin; //PROJECT HPASS COSTING

                        tabFTTH.Cells["F10:F14,M7:M16"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin; //MATERIAL VENDOR
                        tabFTTH.Cells["E11:E14,F11:F14,L8:L16,M8:M16"].Style.Border.Right.Style = ExcelBorderStyle.Thin; ////MATERIAL VENDOR

                        tabFTTH.Cells["F5:G5,F6,F8:F9,H8,M8:M16,J5:R5"].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                        tabFTTH.Cells["C18:Q19,V18:X19"].Style.Fill.BackgroundColor.SetColor(Color.LightPink);
                        tabFTTH.Cells["R18:R19,W18:W19"].Style.Fill.BackgroundColor.SetColor(Color.SandyBrown);
                        tabFTTH.Cells["S18:V19"].Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);

                        tabFTTH.Cells["C2"].Value = "HSBB FTTH";
                        tabFTTH.Cells["C3"].Value = "PROJECT COST BREAKDOWN";
                        
                        tabFTTH.Cells["C5"].Value = "PROJECT NO :";
                        tabFTTH.Cells["F5"].Value = queryProj.PROJECT_NO;
                        tabFTTH.Cells["I5"].Value = "PROJECT TITLE:";
                        tabFTTH.Cells["J5"].Value = projDesc;
                        tabFTTH.Cells["C6"].Value = "REGION :";
                        tabFTTH.Cells["F6"].Value = queryRegion.REGION_NAME_GRN;
                        tabFTTH.Cells["C8"].Value = "H/PASS DSIDE: PHYSICAL :";
                        tabFTTH.Cells["G8"].Value = "CONNECTED :";
                        tabFTTH.Cells["L8"].Value = "VENDOR KABEL :";
                        tabFTTH.Cells["C9"].Value = "H/PASS HIGH RISE:";
                        tabFTTH.Cells["L9"].Value = "VENDOR FDP :";
                        tabFTTH.Cells["F10"].Value = "TOTAL";
                        //tabFTTH.Cells["G10"].Value = "E/SIDE SPUR";

                        tabFTTH.Cells["F14"].Formula = "=SUM(F11:F13)";
                        //tabFTTH.Cells["G14"].Formula = "=SUM(G11:G13)";
                        //tabFTTH.Cells["H14"].Formula = "=SUM(H11:H13)";
                        //tabFTTH.Cells["M8"].Value = vendkabel;
                        //tabFTTH.Cells["M9"].Value = vendfdp;
                        //tabFTTH.Cells["M10"].Value = vendfdc;
                        /*using (Entities1 ctxdata1 = new Entities1())
                        {
                            var vendodf = (from a in ctxdata1.GC_NETELEM2
                                           join b in ctxdata1.GC_ODF on a.G3E_FID equals b.G3E_FID
                                           where a.SCHEME_NAME == schemeName
                                           select b.CONTRACTOR).FirstOrDefault();

                            var vendclosure = (from a in ctxdata1.GC_NETELEM2
                                               join b in ctxdata1.GC_FSPLICE on a.G3E_FID equals b.G3E_FID
                                               where a.SCHEME_NAME == schemeName
                                               select b.CONTRACTOR).FirstOrDefault();
                            tabFTTH.Cells["M11"].Value = vendclosure;
                            tabFTTH.Cells["M13"].Value = vendodf;
                        }*/
                        tabFTTH.Cells["L10"].Value = "VENDOR FDC :";
                        //tabFTTH.Cells["H10"].Value = "TOTAL";
                        tabFTTH.Cells["L11"].Value = "VENDOR CLOSURE :";
                        tabFTTH.Cells["L12"].Value = "SEM @ SBH/SWK : :";
                        if (queryRegion.NATIONWIDE_GRP == "WEST")
                            tabFTTH.Cells["M12"].Value = "SEM";
                        else if (queryRegion.NATIONWIDE_GRP == "EAST")
                            tabFTTH.Cells["M12"].Value = "SBH/SWK";
                        tabFTTH.Cells["L13"].Value = "VENDOR ODF";
                        tabFTTH.Cells["L14"].Value = "LOCATION FACTOR";
                        if (queryRegion.NATIONWIDE_GRP == "WEST")
                            tabFTTH.Cells["M14"].Value = 1.00;
                        else if (queryRegion.NATIONWIDE_GRP == "EAST")
                            tabFTTH.Cells["M14"].Value = 1.15;
                        tabFTTH.Cells["L15"].Value = "COST / HOMEPASS :";
                        tabFTTH.Cells["N15"].Value = "FTTH";
                        tabFTTH.Cells["L16"].Value = "COST / HOMEPASS :";
                        tabFTTH.Cells["N16"].Value = "FTTH HIGHRISE";
                        tabFTTH.Cells["W8"].Value = "Costing Benchmarking:";
                        tabFTTH.Cells["C11"].Value = "MATERIALS COST";
                        tabFTTH.Cells["C12"].Value = "INCIDENTAL COST";
                        tabFTTH.Cells["C13"].Value = "LABOUR COST";
                        tabFTTH.Cells["F11"].Value = 0;
                        tabFTTH.Cells["F13"].Value = 0;
                        //tabFTTH.Cells["G11"].Value = 0;
                        //tabFTTH.Cells["G13"].Value = 0;
                        tabFTTH.Cells["C18"].Value = "EXPENDITURE TYPE";
                        tabFTTH.Cells["R18"].Value = "QTY";
                        tabFTTH.Cells["S18"].Value = "UNIT PRICE";
                        tabFTTH.Cells["W18"].Value = "COST";
                        tabFTTH.Cells["X18"].Value = "REMARKS";
                        tabFTTH.Cells["C20"].Value = "1";
                        tabFTTH.Cells["C20"].Style.Font.Size = 11;
                        tabFTTH.Cells["D20"].Value = "MATERIAL";
                        tabFTTH.Cells["D20"].Style.Font.Size = 11;
                        WebService._base.OSPBOQ_MAIN_EXCEL BOQ_REPORT = new WebService._base.OSPBOQ_MAIN_EXCEL();
                        BOQ_REPORT = myWebService.GetOSPBOQ_MAT_Excel2(0, 100, schemeName);
                        int count = BOQ_REPORT.BOQ_MAIN_Excel.Count;

                        tabFTTH.Cells["N21"].Value = "UNIT";
                        tabFTTH.Cells["N21"].Style.Font.UnderLine = true;
                        tabFTTH.Cells["R21"].Value = "QTY";
                        tabFTTH.Cells["R21"].Style.Font.UnderLine = true;
                        int j = 23;

                        for (int i = 0; i < count; i++)
                        {
                            tabFTTH.Cells["E" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID;
                            tabFTTH.Cells["F" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_DESC;
                            tabFTTH.Cells["N" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_UOM;
                            tabFTTH.Cells["R" + j].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                            tabFTTH.Cells["R" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            tabFTTH.Cells["Q" + j + ":R" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            tabFTTH.Cells["R" + j + ":W" + j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            tabFTTH.Cells["R" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_MAT_PRICE;
                            tabFTTH.Cells["S" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_MATERIAL_VALUE;
                            tabFTTH.Cells["W" + j].Style.Font.Bold = true;
                            tabFTTH.Cells["W" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_TOTAL_VALUE;
                            j += 1;
                        }
                        int k = j + 1;
                        tabFTTH.Cells["E" + j + ":E" + k + ",R" + j + ":R" + k].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                        tabFTTH.Cells["D" + j + ":E" + k + ",Q" + j + ":R" + k].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabFTTH.Cells["W" + j + ":W" + k].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        tabFTTH.Cells["W" + j].Value = "0.00";
                        tabFTTH.Cells["W" + k].Value = "0.00";
                        k++;
                        tabFTTH.Cells["E" + j + ":E" + k + ",R" + j + ":R" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        j = k + 1;
                        tabFTTH.Cells["W" + k + ":W" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabFTTH.Cells["V" + j + ":W" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabFTTH.Cells["W" + j].Style.Font.Bold = true;
                        tabFTTH.Cells["W" + j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        double totalMat = 0;
                        if (count > 0)
                            totalMat = Convert.ToDouble(BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_TOTAL_VALUE_TOTAL);
                        else
                            totalMat = Convert.ToDouble(String.Format("{0:0.00}", (0.00)));
                        tabFTTH.Cells["W" + j].Value = totalMat;

                        tabFTTH.Cells["F11"].Formula = "=W" + j;

                        tabFTTH.Cells["X" + j].Style.Font.Bold = true;
                        tabFTTH.Cells["X" + j].Value = "Subtotal Materials Cost";

                        j += 2;
                        tabFTTH.Cells["B" + j + ":Z" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        k = j + 1;
                        tabFTTH.Cells["C" + j + ":D" + j + ",Q" + k].Style.Font.Bold = true;
                        tabFTTH.Cells["C" + j + ":D" + j].Style.Font.Size = 11;
                        tabFTTH.Cells["C" + j].Value = 2;
                        tabFTTH.Cells["D" + j].Value = "INCIDENTALS";
                        j = k + 1;
                        tabFTTH.Cells["E" + k].Value = "PU NO.";
                        tabFTTH.Cells["F" + k].Value = "PU DESCRIPTION";
                        tabFTTH.Cells["N" + k].Style.Font.Color.SetColor(Color.Red);
                        tabFTTH.Cells["N" + k].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        tabFTTH.Cells["N" + k].Value = "Rate";
                        tabFTTH.Cells["N" + k + ":P" + k].Merge = true;
                        BOQ_REPORT = myWebService.GetOSPBOQ_MAIN_Excel2(0, 100, schemeName);
                        count = BOQ_REPORT.BOQ_MAIN_Excel.Count;
                        int count2 = count;

                        j++;
                        //tabFTTH.Cells["E" + j].Value = "DSIDE";
                        //tabFTTH.Cells["E" + j].Style.Font.Bold = true;
                        double totds1 = 0;
                        //double totelse1 = 0;
                        for (int p = 0; p < count; p++)
                        {

                            /*int puid = Convert.ToInt32(BOQ_REPORT.BOQ_MAIN_Excel[p].x_PU_ID);
                            var queryDSIDE = (from a in ctxdata.WV_FEAT_MAST
                                              where (a.DAY == puid || a.HOLIDAY == puid || a.NIGHT == puid || a.WEEKEND == puid)
                                              select a.LOOKUP_TABLE).FirstOrDefault();
                            if (queryDSIDE != "REF_FDC")
                            {*/

                                j++;
                                tabFTTH.Cells["E" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[p].x_PU_ID;
                                tabFTTH.Cells["F" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[p].x_PU_DESC;
                                tabFTTH.Cells["N" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[p].x_PU_UOM;
                                tabFTTH.Cells["P" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[p].x_RATE_INDICATOR;
                                tabFTTH.Cells["R" + j].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                                tabFTTH.Cells["R" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                tabFTTH.Cells["Q" + j + ":R" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                tabFTTH.Cells["R" + j + ":W" + j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                tabFTTH.Cells["R" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[p].x_PU_QTY;
                                tabFTTH.Cells["S" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[p].x_BQ_INSTALL_PRICE;
                                tabFTTH.Cells["W" + j].Style.Font.Bold = true;
                                tabFTTH.Cells["W" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[p].x_INSTALL_VALUE;

                                //if (queryDSIDE != "REF_FDC")
                                //{
                                if (BOQ_REPORT.BOQ_MAIN_Excel[p].x_INSTALL_VALUE != "-")
                                {
                                    double totds = Convert.ToDouble(BOQ_REPORT.BOQ_MAIN_Excel[p].x_INSTALL_VALUE);
                                    totds1 = totds1 + totds;
                                }
                                //}
                                //tabFTTH.Cells["F12"].Value = totds1;
                            //}
                            //if (!(queryDSIDE == "REF_FDP" || queryDSIDE == "REF_FDC"))
                            //{
                            //    if (BOQ_REPORT.BOQ_MAIN_Excel[p].x_INSTALL_VALUE != "-")
                            //    {
                            //        double totelse = Convert.ToDouble(BOQ_REPORT.BOQ_MAIN_Excel[p].x_INSTALL_VALUE);
                            //        totelse1 = totelse1 + totelse;
                            //    }
                            //}
                        }
                        //totelse1 /= 2;
                        tabFTTH.Cells["F12"].Value = totds1; //+ totelse1
                        j++;
                        /*tabFTTH.Cells["E" + j].Value = "ESIDE SPUR";
                        tabFTTH.Cells["E" + j].Style.Font.Bold = true;

                        double toteside1 = 0;
                        for (int q = 0; q < count; q++)
                        {
                            int puid2 = Convert.ToInt32(BOQ_REPORT.BOQ_MAIN_Excel[q].x_PU_ID);

                            var queryESIDE = (from a in ctxdata.WV_FEAT_MAST
                                              where (a.DAY == puid2 || a.HOLIDAY == puid2 || a.NIGHT == puid2 || a.WEEKEND == puid2)
                                              select a.LOOKUP_TABLE).FirstOrDefault();
                            if (queryESIDE == "REF_FDC")
                            {
                                j++;
                                tabFTTH.Cells["E" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[q].x_PU_ID;
                                tabFTTH.Cells["F" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[q].x_PU_DESC;
                                tabFTTH.Cells["N" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[q].x_PU_UOM;
                                tabFTTH.Cells["P" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[q].x_RATE_INDICATOR;
                                tabFTTH.Cells["R" + j].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                                tabFTTH.Cells["R" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                tabFTTH.Cells["Q" + j + ":R" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                tabFTTH.Cells["R" + j + ":W" + j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                tabFTTH.Cells["R" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[q].x_PU_QTY;
                                tabFTTH.Cells["S" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[q].x_BQ_INSTALL_PRICE;
                                tabFTTH.Cells["W" + j].Style.Font.Bold = true;
                                tabFTTH.Cells["W" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[q].x_INSTALL_VALUE;

                                //if (queryESIDE == "REF_FDC")
                                //{
                                if (BOQ_REPORT.BOQ_MAIN_Excel[q].x_INSTALL_VALUE != "-")
                                {
                                    double toteside = Convert.ToDouble(BOQ_REPORT.BOQ_MAIN_Excel[q].x_INSTALL_VALUE);
                                    toteside1 = toteside1 + toteside;
                                }
                                //}
                                //tabFTTH.Cells["G12"].Value = toteside1;
                            }
                        }
                        tabFTTH.Cells["G12"].Value = toteside1; //+ totelse1
                         
                        j++;
                        */
                        k = j + 1;
                        tabFTTH.Cells["E" + j + ":E" + k + ",P" + j + ":P" + k + ",R" + j + ":R" + k].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                        tabFTTH.Cells["D" + j + ":E" + k + ",O" + j + ":P" + k + ",Q" + j + ":R" + k].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabFTTH.Cells["W" + j + ":W" + k].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        tabFTTH.Cells["W" + j].Value = "0.00";
                        tabFTTH.Cells["W" + k].Value = "0.00";
                        k++;
                        tabFTTH.Cells["E" + j + ":E" + k + ",P" + j + ":P" + k + ",R" + j + ":R" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        j = k + 2;
                        k++;
                        tabFTTH.Cells["D" + j].Style.Font.Bold = true;
                        tabFTTH.Cells["D" + j].Value = "TNT : MILEAGE + MEAL + TOLL etc";
                        tabFTTH.Cells["I" + j].Value = "AND TnT";
                        tabFTTH.Cells["W" + k + ":W" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabFTTH.Cells["V" + j + ":W" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                        WebService._base.SUMM_ProjCost estimateBP = new WebService._base.SUMM_ProjCost();
                        estimateBP = myWebService.Get_ProjCost(schemeName);
                        count = estimateBP.ProjCost.Count;

                        double SumLbrOT = 0;
                        double SumLbrSalary = 0;
                        double SumTNT_Mileage = 0;
                        double SumTotal = 0;
                        for (int i = 0; i < count; i++)
                        {
                            SumLbrOT = Convert.ToDouble(estimateBP.ProjCost[i].SUM_LBR_OT);
                            SumLbrSalary = Convert.ToDouble(estimateBP.ProjCost[i].SUM_LBR_SALARY);
                            SumTNT_Mileage = Convert.ToDouble(estimateBP.ProjCost[i].SUM_TNT_MILEAGE);
                            SumTotal = Convert.ToDouble(estimateBP.ProjCost[i].SUM_TOTAL);
                        }
                        tabFTTH.Cells["W" + j].Style.Font.Bold = true;
                        tabFTTH.Cells["W" + j].Value = SumTNT_Mileage;
                        j++;
                        k++;
                        tabFTTH.Cells["I" + j].Value = "AD TnT";
                        tabFTTH.Cells["W" + k + ":W" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabFTTH.Cells["V" + j + ":W" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabFTTH.Cells["W" + j].Style.Font.Bold = true;

                        double planNImple = 0;
                        if (queryRegion.REGION_NAME_GRN == "CENTRAL")
                            planNImple = 130;
                        else if (queryRegion.REGION_NAME_GRN == "NORTHERN")
                            planNImple = 560;
                        else if (queryRegion.REGION_NAME_GRN == "SOUTHERN")
                            planNImple = 500;
                        else if (queryRegion.REGION_NAME_GRN == "EASTERN")
                            planNImple = 540;
                        else if (queryRegion.REGION_NAME_GRN == "SABAH")
                            planNImple = 910;
                        else if (queryRegion.REGION_NAME_GRN == "SARAWAK")
                            planNImple = 760;

                        double totalInci = 0;
                        if (count2 > 0)
                            totalInci = Convert.ToDouble(BOQ_REPORT.BOQ_MAIN_Excel[count2 - 1].x_INSTALL_VALUE_TOTAL);
                        else
                            totalInci = Convert.ToDouble(String.Format("{0:0.00}", (0.00)));
                        double matPlusInci = totalMat + totalInci;

                        double adMileage = 0;
                        if (matPlusInci <= 25000)
                            adMileage = 100;
                        else
                            adMileage = planNImple;

                        tabFTTH.Cells["W" + j].Value = adMileage;
                        j += 2;
                        k = j + 1;
                        tabFTTH.Cells["W" + j + ":W" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabFTTH.Cells["V" + j + ":W" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabFTTH.Cells["W" + j].Style.Font.Bold = true;
                        double inciPlusMileage = totalInci + SumTNT_Mileage + adMileage;
                        tabFTTH.Cells["W" + j].Value = Convert.ToDouble(String.Format("{0:0.00}", inciPlusMileage));

                        tabFTTH.Cells["F12"].Formula = "=W" + j;

                        tabFTTH.Cells["X" + j].Style.Font.Bold = true;
                        tabFTTH.Cells["X" + j].Value = "Subtotal Incidentals Cost";

                        j += 2;
                        tabFTTH.Cells["B" + j + ":Z" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabFTTH.Cells["C" + j + ":D" + j].Style.Font.Bold = true;
                        tabFTTH.Cells["C" + j + ":D" + j].Style.Font.Size = 11;
                        tabFTTH.Cells["C" + j].Value = 3;
                        tabFTTH.Cells["D" + j].Value = "LABOUR";
                        j++;
                        k = j + 1;
                        tabFTTH.Cells["E" + j].Value = "OVERTIME";
                        tabFTTH.Cells["G" + j].Value = "AND Staff OT";
                        tabFTTH.Cells["W" + j + ":W" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabFTTH.Cells["V" + j + ":W" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabFTTH.Cells["W" + j].Style.Font.Bold = true;
                        tabFTTH.Cells["W" + j].Value = SumLbrOT;
                        j++;
                        k = j + 1;
                        tabFTTH.Cells["E" + j].Value = "SALARY";
                        tabFTTH.Cells["G" + j].Value = "AND Staff Salary";
                        tabFTTH.Cells["W" + j + ":W" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabFTTH.Cells["V" + j + ":W" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabFTTH.Cells["W" + j].Style.Font.Bold = true;
                        tabFTTH.Cells["W" + j].Value = SumLbrSalary;
                        j++;
                        k = j + 1;
                        tabFTTH.Cells["E" + j].Value = "SALARY";
                        tabFTTH.Cells["G" + j].Value = "Access Development Staff Salary";
                        tabFTTH.Cells["W" + j + ":W" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabFTTH.Cells["V" + j + ":W" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabFTTH.Cells["W" + j].Style.Font.Bold = true;
                        tabFTTH.Cells["W" + j].Value = 10;
                        j++;
                        k = j + 1;
                        tabFTTH.Cells["W" + j + ":W" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabFTTH.Cells["V" + j + ":W" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabFTTH.Cells["W" + j].Style.Font.Bold = true;
                        tabFTTH.Cells["W" + j].Value = SumLbrOT + SumLbrSalary + 10;
                        double totalLabour = SumLbrOT + SumLbrSalary + 10;

                        tabFTTH.Cells["F13"].Formula = "=W" + j;

                        tabFTTH.Cells["X" + j].Style.Font.Bold = true;
                        tabFTTH.Cells["X" + j].Value = "Subtotal Labour Cost";
                        j += 2;
                        k = j + 1;
                        tabFTTH.Cells["W" + j + ":W" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabFTTH.Cells["V" + j + ":W" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabFTTH.Cells["W" + j].Style.Font.Bold = true;
                        double grandTot = matPlusInci + totalLabour + SumTNT_Mileage + adMileage;
                        tabFTTH.Cells["W" + j].Value = Convert.ToDouble(String.Format("{0:0.00}", grandTot));
                        tabFTTH.Cells["X" + j].Style.Font.Bold = true;
                        tabFTTH.Cells["X" + j].Value = "GRAND TOTAL COST";
                        j += 2;
                        tabFTTH.Cells["B" + j + ":Z" + j].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        j++;
                        tabFTTH.Cells["D" + j + ":S" + j].Style.Font.Bold = true;
                        tabFTTH.Cells["D" + j + ":S" + j].Style.Font.Size = 10;
                        tabFTTH.Cells["D" + j].Value = "PREPARED BY :";
                        tabFTTH.Cells["J" + j].Value = "CHECKED BY:";
                        tabFTTH.Cells["S" + j].Value = "ENDORSED BY:";
                        j += 2;
                        tabFTTH.Cells["D" + j + ":F" + j + ",J" + j + ":M" + j + ",S" + j + ":W" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        j += 2;
                        tabFTTH.Cells["B" + j + ":Z" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

                        var stream = new MemoryStream();
                        xlsFile.SaveAs(stream);

                        string fileName = "BOQReport_" + schemeName + ".xlsx";
                        string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                        stream.Position = 0;
                        return File(stream, contentType, fileName);
                        #endregion
                    }
                    else if (queryProj.SCHEME_TYPE == "Fiber (Equip)" || queryProj.SCHEME_TYPE == "Fiber E/Side" || queryProj.SCHEME_TYPE == "HSBB (Equip)" || queryProj.SCHEME_TYPE == "HSBB E/Side" || queryProj.SCHEME_TYPE == "HSBB D/Side" || queryProj.SCHEME_TYPE == "Fiber Trunk" || queryProj.SCHEME_TYPE == "Fiber Junction" || queryProj.SCHEME_TYPE == "METROE")
                    {
                        #region foc
                        var tabFOC = xlsFile.Workbook.Worksheets.Add("FOC");
                        tabFOC.Cells.Style.Font.Name = "Arial";
                        tabFOC.Cells.Style.Font.Size = 8;
                        tabFOC.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        tabFOC.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                        tabFOC.Cells["A1:AD266"].Style.Fill.BackgroundColor.SetColor(Color.White);
                        tabFOC.Cells["A1:A265,AC1:AC265"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        tabFOC.DefaultRowHeight = 13.2;
                        tabFOC.Column(1).Width = 0.77; //A
                        tabFOC.Column(2).Width = 1; //B
                        tabFOC.Column(3).Width = 2.5; //C
                        tabFOC.Column(4).Width = 8; //D
                        tabFOC.Column(5).Width = 8; //E
                        tabFOC.Column(6).Width = 8; //F
                        tabFOC.Column(7).Width = 8; //G
                        tabFOC.Column(8).Width = 8; //H
                        tabFOC.Column(9).Width = 5; //I
                        tabFOC.Column(10).Width = 0.77; //J
                        tabFOC.Column(11).Width = 5.38; //K
                        tabFOC.Column(12).Width = 5.38; //L
                        tabFOC.Column(13).Width = 4.38; //M
                        tabFOC.Column(14).Width = 4.38; //N
                        for (int i = 15; i < 18; i++)
                            tabFOC.Column(i).Width = 7.13; //O P Q
                        tabFOC.Column(18).Width = 0.77; //R
                        tabFOC.Column(19).Width = 7.5; //S
                        for (int i = 20; i < 29; i++)
                            tabFOC.Column(i).Width = 7; // T U V W X Y Z AA AB 
                        tabFOC.Column(29).Width = 1.5; //AC
                        tabFOC.Column(30).Width = 0.54; //AC

                        tabFOC.Row(1).Height = 15.6;
                        tabFOC.Row(2).Height = 10.2;
                        tabFOC.Row(21).Height = 5.3;
                        tabFOC.Cells["C1:AB6,E8:H8,D9:D11,D22,C13:H15,U11:W15"].Style.Font.Bold = true;
                        tabFOC.Cells["C1,G1"].Style.Font.UnderLine = true;
                        tabFOC.Cells["L2"].Style.Font.Color.SetColor(Color.Red);
                        tabFOC.Cells["C1"].Style.Font.Size = 12;
                        tabFOC.Cells["C3,C4,C5"].Style.Font.Size = 10;
                        tabFOC.Cells["V11:V15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        tabFOC.Cells["K2,E8:H8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        tabFOC.Cells["F2:G2,F3:P3,F4:P4,F5:G5,K5:N5,K6:N6,K7:N7,K8:N8,K9:N9,K10:N10,K11:N11,K12:N12,K13:N13,K14:N14,K15:N15,K16:N16,K17:N17,K18:N18"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin; //KOTAK PROJECT + CORE USAGE
                        tabFOC.Cells["E3:E5,G3,P4,G5,J6:J18,K6:K18,L6:L18,M6:M18,N6:N18"].Style.Border.Right.Style = ExcelBorderStyle.Thin; //KOTAK PROJECT + CORE USAGE

                        tabFOC.Cells["F7:F13,C19:AB19,C21:AB21"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin; //KOTAK MATERIAL
                        tabFOC.Cells["E8:E13,F8:F13"].Style.Border.Right.Style = ExcelBorderStyle.Thin; //KOTAK MATERIAL

                        //tabFOC.Cells["X5:AB5,V6:AB6,V7:AB7,V8:AB8,W10:W15"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin; //KOTAK INNERDUCT LOC FACTOR
                        //tabFOC.Cells["U7:U8,W6:W8,X6:X8,Y6:Y8,Z6:Z8,AA6:AA8,AB6:AB8,V11:V15,W11:W15"].Style.Border.Right.Style = ExcelBorderStyle.Thin; //KOTAK MATERIAL

                        tabFOC.Cells["W10:W15"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabFOC.Cells["V11:W15"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                        tabFOC.Cells["F3:G3,F4:P4,F5:G5,W11:W15"].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                        tabFOC.Cells["C20:AB21"].Style.Fill.BackgroundColor.SetColor(Color.LightPink);
                        tabFOC.Cells["F8,K6:N6"].Style.Fill.BackgroundColor.SetColor(Color.SandyBrown);

                        tabFOC.Cells["C1"].Value = "PROJECT COST BREAKDOWN";
                        tabFOC.Cells["C3"].Value = "PROJECT NO :";
                        tabFOC.Cells["C4"].Value = "PROJECT TITLE:";
                        tabFOC.Cells["F4"].Value = projDesc;
                        tabFOC.Cells["C5"].Value = "REGION :";
                        tabFOC.Cells["F5"].Value = queryRegion.REGION_NAME_GRN;
                        tabFOC.Cells["K6"].Value = "%KM";
                        tabFOC.Cells["L6"].Value = "CORE";
                        tabFOC.Cells["M6"].Value = "KM";
                        tabFOC.Cells["N6"].Value = "FKM";
                        //tabFOC.Cells["O6"].Value = "MAT RM";
                        //tabFOC.Cells["P6"].Value = "INC RM";
                        //tabFOC.Cells["Q6"].Value = "LAB RM";
                        //tabFOC.Cells["S6"].Value = "TOT RM";
                        //tabFOC.Cells["T6"].Value = "COST/KM";
                        //tabFOC.Cells["X6"].Value = "%KM";
                        //tabFOC.Cells["Y6"].Value = "KM";
                        //tabFOC.Cells["Z6"].Value = "MAT RM";
                        //tabFOC.Cells["AA6"].Value = "INC RM";
                        //tabFOC.Cells["AB6"].Value = "COST/KM";

                        tabFOC.Cells["L7"].Value = "12C";
                        tabFOC.Cells["L8"].Value = "24C";
                        tabFOC.Cells["L9"].Value = "36C";
                        tabFOC.Cells["L10"].Value = "48C";
                        tabFOC.Cells["L11"].Value = "96C";
                        tabFOC.Cells["L12"].Value = "144C";
                        tabFOC.Cells["L13"].Value = "192C";
                        tabFOC.Cells["L14"].Value = "12C IB";
                        tabFOC.Cells["L15"].Value = "24C IB";
                        tabFOC.Cells["L16"].Value = "36C IB";
                        tabFOC.Cells["L17"].Value = "48C IB";
                        tabFOC.Cells["L18"].Value = "192C IB";

                        for (int i = 7; i < 19; i++)
                        {
                            tabFOC.Cells["M" + i].Value = 0.00;
                            tabFOC.Cells["K" + i].Style.Numberformat.Format = "0%";
                            tabFOC.Cells["M" + i].Style.Numberformat.Format = "0.00";
                        }

                        for (int m = 7; m < 11; m++)
                        {
                            tabFOC.Cells["N" + m].Formula = "=M" + m;
                            tabFOC.Cells["K" + m].Formula = "=IFERROR(M" + m + "/SUBTOTAL(9,$M$7:$M$10),)";
                            //tabFOC.Cells["O" + m].Formula = "=IFERROR((M" + m + "/SUM($M$7:$M$10,$M$14:$M$15)*($F$10)),)";
                            //tabFOC.Cells["Q" + m].Formula = "=IFERROR($F$11*(M" + m + "/(SUBTOTAL(9,$M$7:$M$10,$M$14:$M$15))),)";
                            //tabFOC.Cells["S" + m].Formula = "=SUM(O" + m + ":Q" + m + ")";
                            //tabFOC.Cells["T" + m].Formula = "=S" + m + "/M" + m;
                        }
                        for (int m = 11; m < 14; m++)
                        {
                            tabFOC.Cells["N" + m].Formula = "=M" + m;
                            tabFOC.Cells["K" + m].Formula = "=IFERROR(M" + m + "/SUBTOTAL(9,$M$11:$M$13),)";
                            //tabFOC.Cells["O" + m].Formula = "=IFERROR((M" + m + "/SUM($M$11:$M$13,$M$16:$M$18)*($E$10)),)";
                            //tabFOC.Cells["Q" + m].Formula = "=IFERROR($E$11*(M" + m + "/(SUBTOTAL(9,$M$11:$M$13,$M$16:$M$18))),)";
                            //tabFOC.Cells["S" + m].Formula = "=SUM(O" + m + ":Q" + m + ")";
                            //tabFOC.Cells["T" + m].Formula = "=S" + m + "/M" + m;
                        }
                        for (int m = 14; m < 16; m++)
                        {
                            tabFOC.Cells["N" + m].Formula = "=M" + m;
                            tabFOC.Cells["K" + m].Formula = "=IFERROR(M" + m + "/SUBTOTAL(9,$M$14:$M$15),)";
                            //tabFOC.Cells["O" + m].Formula = "=IFERROR((M" + m + "/SUM($M$7:$M$10,$M$14:$M$15)*($F$10)),)";
                            //tabFOC.Cells["Q" + m].Formula = "=IFERROR($F$11*(M" + m + "/(SUBTOTAL(9,$M$7:$M$10,$M$14:$M$15))),)";
                            //tabFOC.Cells["S" + m].Formula = "=SUM(O" + m + ":Q" + m + ")";
                            //tabFOC.Cells["T" + m].Formula = "=S" + m + "/M" + m;
                        }
                        for (int m = 16; m < 19; m++)
                        {
                            tabFOC.Cells["N" + m].Formula = "=M" + m;
                            tabFOC.Cells["K" + m].Formula = "=IFERROR(M" + m + "/SUBTOTAL(9,$M$16:$M$18),)";
                            //tabFOC.Cells["O" + m].Formula = "=IFERROR((M" + m + "/SUM($M$11:$M$13,$M$16:$M$18)*($E$10)),)";
                            //tabFOC.Cells["Q" + m].Formula = "=IFERROR($E$11*(M" + m + "/(SUBTOTAL(9,$M$11:$M$13,$M$16:$M$18))),)";
                            //tabFOC.Cells["S" + m].Formula = "=SUM(O" + m + ":Q" + m + ")";
                            //tabFOC.Cells["T" + m].Formula = "=S" + m + "/M" + m;
                        }

                        //tabFOC.Cells["V7"].Value = "INNERDUCT 21KM";
                        //tabFOC.Cells["V8"].Value = "INNERDUCT 16KM";
                        tabFOC.Cells["V11"].Value = "LOCATION FACTOR";
                        if (queryRegion.NATIONWIDE_GRP == "WEST")
                            tabFOC.Cells["W11"].Value = 1.00;
                        else if (queryRegion.NATIONWIDE_GRP == "EAST")
                            tabFOC.Cells["W11"].Value = 1.15;
                        tabFOC.Cells["V12"].Value = "VENDOR KABEL";

                        var checkFiberEside = (from a in ctxdata.GC_NETELEM
                                               where a.SCHEME_NAME == schemeName && a.G3E_FNO == 7200
                                               select a).Count();
                        if (checkFiberEside > 0)
                        {
                            var fiberEsideManufac = (from a in ctxdata.GC_NETELEM
                                                     join b in ctxdata.GC_FCBL on a.G3E_FID equals b.G3E_FID
                                                     where a.SCHEME_NAME == schemeName && a.G3E_FNO == 7200
                                                     select b.MANUFACTURER).FirstOrDefault();
                            tabFOC.Cells["W12"].Value = fiberEsideManufac;
                        }

                        tabFOC.Cells["V13"].Value = "VENDOR CLOSURE";

                        using (Entities1 ctxdata2 = new Entities1()) //check performance
                        {
                            var fiberSpliceManufac = (from a in ctxdata2.GC_NETELEM2
                                                      join b in ctxdata2.GC_FSPLICE on a.G3E_FID equals b.G3E_FID
                                                      where a.SCHEME_NAME == schemeName
                                                      select b.CONTRACTOR).FirstOrDefault();
                            tabFOC.Cells["W13"].Value = fiberSpliceManufac;

                            var odfManufac = (from a in ctxdata2.GC_NETELEM2
                                              join b in ctxdata2.GC_ODF on a.G3E_FID equals b.G3E_FID
                                              where a.SCHEME_NAME == schemeName
                                              select b.CONTRACTOR).FirstOrDefault();
                            tabFOC.Cells["W14"].Value = odfManufac;
                        }

                        tabFOC.Cells["V14"].Value = "VENDOR ODF";
                        tabFOC.Cells["V15"].Value = "SEM / SBH&SWK :";
                        if (queryRegion.NATIONWIDE_GRP == "WEST")
                            tabFOC.Cells["W15"].Value = "SEM";
                        else if (queryRegion.NATIONWIDE_GRP == "EAST")
                            tabFOC.Cells["W15"].Value = "SBH/SWK";
                        tabFOC.Cells["F8"].Value = "TOTAL";
                        tabFOC.Cells["C9"].Value = "MATERIALS";
                        tabFOC.Cells["C10"].Value = "INCIDENTAL";
                        tabFOC.Cells["C11"].Value = "LABOUR";
                        tabFOC.Cells["C13"].Value = "Overall Cost/F-Km =";
                        tabFOC.Cells["F12"].Formula = "=SUM(F9:F11)";
                        tabFOC.Cells["F13"].Formula = "=F12/SUBTOTAL(9,M7:M17)";
                        tabFOC.Cells["F13"].Style.Numberformat.Format = "0.00";

                        tabFOC.Cells["C20"].Value = "EXPENDITURE TYPE";
                        tabFOC.Cells["L20"].Value = "UNIT";
                        tabFOC.Cells["S20"].Value = "QTY";
                        tabFOC.Cells["Q20"].Value = "RATE";
                        tabFOC.Cells["U20"].Value = "UNIT PRICE";
                        tabFOC.Cells["W20"].Value = "COST";
                        tabFOC.Cells["X20"].Value = "REMARKS";
                        tabFOC.Cells["C22"].Value = "1";
                        tabFOC.Cells["C22"].Style.Font.Size = 11;
                        tabFOC.Cells["D22"].Value = "MATERIAL";
                        tabFOC.Cells["D22"].Style.Font.Size = 11;

                        
                        string strTableTgh = "1000000190|1000000708|1000000709|1000000710|1000000711|1000000832|1000000833|1000000714|1000000792|1000000793|1000000794|1000003444";

                        string[] tableTgh = strTableTgh.Split('|');

                        WebService._base.OSPBOQ_MAIN_EXCEL BOQ_REPORT = new WebService._base.OSPBOQ_MAIN_EXCEL();
                        BOQ_REPORT = myWebService.GetOSPBOQ_MAT_Excel2(0, 100, schemeName);
                        int count = BOQ_REPORT.BOQ_MAIN_Excel.Count;

                        int j = 24;
                        for (int i = 0; i < count; i++)
                        {

                            tabFOC.Cells["D" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID;
                            tabFOC.Cells["F" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_DESC;
                            tabFOC.Cells["L" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_UOM;
                            tabFOC.Cells["S" + j].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                            tabFOC.Cells["S" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            tabFOC.Cells["R" + j + ":S" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            tabFOC.Cells["S" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_MAT_PRICE;
                            tabFOC.Cells["S" + j].Style.Numberformat.Format = "0.00";
                            tabFOC.Cells["U" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_MATERIAL_VALUE;
                            tabFOC.Cells["W" + j].Style.Font.Bold = true;
                            tabFOC.Cells["W" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_TOTAL_VALUE;
                            
                            for (int l = 0; l < tableTgh.Count(); l++)
                            {
                                int m = l + 7;
                                if (BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID.Trim() == tableTgh[l])
                                    tabFOC.Cells["M"  + m].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_UOM;
                            }
                            j += 1;
                        }
                        
                        int k = j + 1;
                        /*tabFOC.Cells["D" + j + ":D" + k + ",S" + j + ":S" + k].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                        tabFOC.Cells["C" + j + ":D" + k + ",R" + j + ":S" + k].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabFOC.Cells["W" + j].Value = "0.00";
                        tabFOC.Cells["W" + k].Value = "0.00";*/
                        k++;
                        //tabFOC.Cells["D" + j + ":D" + k + ",S" + j + ":S" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        j = k + 1;
                        tabFOC.Cells["W" + k + ":W" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabFOC.Cells["V" + j + ":W" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabFOC.Cells["W" + j].Style.Font.Bold = true;
                        double totalMat = 0;
                        if (count > 0)
                            totalMat = Convert.ToDouble(BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_TOTAL_VALUE_TOTAL);
                        else
                            totalMat = Convert.ToDouble(String.Format("{0:0.00}", (0.00)));
                        tabFOC.Cells["W" + j].Value = totalMat;
                        tabFOC.Cells["F9"].Value = totalMat;
                        tabFOC.Cells["X" + j].Style.Font.Bold = true;
                        tabFOC.Cells["X" + j].Value = "Subtotal Materials Cost";
                        j += 2;
                        tabFOC.Cells["B" + j + ":AC" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        k = j + 1;
                        tabFOC.Cells["C" + j + ":D" + j + ",Q" + k].Style.Font.Bold = true;
                        tabFOC.Cells["C" + j + ":D" + j].Style.Font.Size = 11;
                        tabFOC.Cells["C" + j].Value = 2;
                        tabFOC.Cells["D" + j].Value = "INCIDENTALS";
                        j = k + 1;

                        BOQ_REPORT = myWebService.GetOSPBOQ_MAIN_Excel2(0, 100, schemeName);
                        count = BOQ_REPORT.BOQ_MAIN_Excel.Count;
                        int count2 = count;
                        decimal totIncidental = 0;
                        for (int i = 0; i < count; i++)
                        {
                            if (BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONTRACT_NO == null || BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONTRACT_NO == "")
                            {
                                tabFOC.Cells["D" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_ID;
                                tabFOC.Cells["F" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_DESC;
                                tabFOC.Cells["L" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_UOM;
                                tabFOC.Cells["Q" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_RATE_INDICATOR;
                                tabFOC.Cells["S" + j].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                                tabFOC.Cells["S" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                tabFOC.Cells["R" + j + ":S" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                tabFOC.Cells["S" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_QTY;
                                tabFOC.Cells["U" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_INSTALL_PRICE;
                                tabFOC.Cells["W" + j].Style.Font.Bold = true;
                                tabFOC.Cells["W" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_INSTALL_VALUE;
                                totIncidental = totIncidental + (Convert.ToDecimal(BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_INSTALL_PRICE) * Convert.ToDecimal(BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_QTY));
                                System.Diagnostics.Debug.WriteLine("Total Incidental" + totIncidental);
                            }
                            else
                            {
                                tabFOC.Cells["D" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONTRACT_NO;
                                tabFOC.Cells["E" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_ITEM_NO;
                                tabFOC.Cells["F" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONTRACT_DESC;
                                tabFOC.Cells["L" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_UOM;
                                tabFOC.Cells["S" + j].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                                tabFOC.Cells["S" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                tabFOC.Cells["R" + j + ":S" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                tabFOC.Cells["S" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_PU_QTY;
                                tabFOC.Cells["U" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_BQ_CONT_PRICE;
                                tabFOC.Cells["W" + j].Style.Font.Bold = true;
                                tabFOC.Cells["W" + j].Value = BOQ_REPORT.BOQ_MAIN_Excel[i].x_CONT_VALUE;
                            }
                            
                            j += 1;
                        }
                        
                        
                        k = j + 1;
                        /*tabFOC.Cells["D" + j + ":D" + k + ",Q" + j + ":Q" + k + ",S" + j + ":S" + k].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                        tabFOC.Cells["C" + j + ":D" + k + ",P" + j + ":Q" + k + ",R" + j + ":S" + k].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabFOC.Cells["W" + j].Value = "0.00";
                        tabFOC.Cells["W" + k].Value = "0.00";*/
                        k++;
                        //tabFOC.Cells["D" + j + ":D" + k + ",Q" + j + ":Q" + k + ",S" + j + ":S" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        j = k + 2;
                        k++;
                        tabFOC.Cells["D" + j].Style.Font.Bold = true;
                        tabFOC.Cells["D" + j].Value = "TNT : MILEAGE + MEAL + TOLL etc";
                        tabFOC.Cells["G" + j].Value = "AND TnT";
                        tabFOC.Cells["W" + k + ":W" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabFOC.Cells["V" + j + ":W" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                        WebService._base.SUMM_ProjCost estimateBP = new WebService._base.SUMM_ProjCost();
                        estimateBP = myWebService.Get_ProjCost(schemeName);
                        count = estimateBP.ProjCost.Count;

                        double SumLbrOT = 0;
                        double SumLbrSalary = 0;
                        double SumTNT_Mileage = 0;
                        double SumTotal = 0;
                        for (int i = 0; i < count; i++)
                        {
                            SumLbrOT = Convert.ToDouble(estimateBP.ProjCost[i].SUM_LBR_OT);
                            SumLbrSalary = Convert.ToDouble(estimateBP.ProjCost[i].SUM_LBR_SALARY);
                            SumTNT_Mileage = Convert.ToDouble(estimateBP.ProjCost[i].SUM_TNT_MILEAGE);
                            SumTotal = Convert.ToDouble(estimateBP.ProjCost[i].SUM_TOTAL);
                        }
                        tabFOC.Cells["W" + j].Style.Font.Bold = true;
                        tabFOC.Cells["W" + j].Value = SumTNT_Mileage;
                        j++;
                        k++;
                        tabFOC.Cells["G" + j].Value = "AD TnT";
                        tabFOC.Cells["W" + k + ":W" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabFOC.Cells["V" + j + ":W" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabFOC.Cells["W" + j].Style.Font.Bold = true;

                        double planNImple = 0;
                        if (queryRegion.REGION_NAME_GRN == "CENTRAL")
                            planNImple = 130;
                        else if (queryRegion.REGION_NAME_GRN == "NORTHERN")
                            planNImple = 560;
                        else if (queryRegion.REGION_NAME_GRN == "SOUTHERN")
                            planNImple = 500;
                        else if (queryRegion.REGION_NAME_GRN == "EASTERN")
                            planNImple = 540;
                        else if (queryRegion.REGION_NAME_GRN == "SABAH")
                            planNImple = 910;
                        else if (queryRegion.REGION_NAME_GRN == "SARAWAK")
                            planNImple = 760;

                        double totalInci = 0;
                        if (count2 > 0)
                            totalInci = Convert.ToDouble(BOQ_REPORT.BOQ_MAIN_Excel[count2 - 1].x_INSTALL_VALUE_TOTAL);
                        else
                            totalInci = Convert.ToDouble(String.Format("{0:0.00}", (0.00)));
                        tabFOC.Cells["F10"].Value = totalInci;
                        double matPlusInci = totalMat + totalInci;

                        double adMileage = 0;
                        if (matPlusInci <= 25000)
                            adMileage = 100;
                        else
                            adMileage = planNImple;

                        
                        
                        tabFOC.Cells["W" + j].Value = adMileage;
                        j += 2;
                        k = j + 1;
                        tabFOC.Cells["W" + j + ":W" + k].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tabFOC.Cells["V" + j + ":W" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabFOC.Cells["W" + j].Style.Font.Bold = true;
                        double inciPlusMileage = totalInci + SumTNT_Mileage + adMileage;
                        tabFOC.Cells["W" + j].Value = Convert.ToDouble(String.Format("{0:0.00}", inciPlusMileage));
                        tabFOC.Cells["X" + j].Style.Font.Bold = true;
                        tabFOC.Cells["X" + j].Value = "Subtotal Incidentals Cost";

                        j += 2;
                        tabFOC.Cells["B" + j + ":AC" + j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabFOC.Cells["C" + j + ":D" + j].Style.Font.Bold = true;
                        tabFOC.Cells["C" + j + ":D" + j].Style.Font.Size = 11;
                        tabFOC.Cells["C" + j].Value = 3;
                        tabFOC.Cells["D" + j].Value = "LABOUR";
                        j++;
                        k = j + 1;
                        tabFOC.Cells["E" + j].Value = "OVERTIME";
                        tabFOC.Cells["G" + j].Value = "AND Staff OT";
                        tabFOC.Cells["W" + j + ":W" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabFOC.Cells["V" + j + ":W" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabFOC.Cells["W" + j].Style.Font.Bold = true;
                        tabFOC.Cells["W" + j].Value = SumLbrOT;
                        j++;
                        k = j + 1;
                        tabFOC.Cells["E" + j].Value = "SALARY";
                        tabFOC.Cells["G" + j].Value = "AND Staff Salary";
                        tabFOC.Cells["W" + j + ":W" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabFOC.Cells["V" + j + ":W" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabFOC.Cells["W" + j].Style.Font.Bold = true;
                        tabFOC.Cells["W" + j].Value = SumLbrSalary;
                        j++;
                        k = j + 1;
                        tabFOC.Cells["E" + j].Value = "SALARY";
                        tabFOC.Cells["G" + j].Value = "Access Development Staff Salary";
                        tabFOC.Cells["W" + j + ":W" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabFOC.Cells["V" + j + ":W" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabFOC.Cells["W" + j].Style.Font.Bold = true;
                        tabFOC.Cells["W" + j].Value = 10;
                        j++;
                        k = j + 1;
                        tabFOC.Cells["W" + j + ":W" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabFOC.Cells["V" + j + ":W" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabFOC.Cells["W" + j].Style.Font.Bold = true;
                        tabFOC.Cells["W" + j].Value = SumLbrOT + SumLbrSalary + 10;
                        double totalLabour = SumLbrOT + SumLbrSalary + 10;
                        tabFOC.Cells["F11"].Value = totalLabour;
                        
                        

                        tabFOC.Cells["X" + j].Style.Font.Bold = true;
                        tabFOC.Cells["X" + j].Value = "Subtotal Labour Cost";

                        j += 2;
                        k = j + 1;
                        tabFOC.Cells["W" + j + ":W" + k].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tabFOC.Cells["V" + j + ":W" + j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tabFOC.Cells["W" + j].Style.Font.Bold = true;
                        double grandTot = matPlusInci + totalLabour + SumTNT_Mileage + adMileage;
                        tabFOC.Cells["W" + j].Value = Convert.ToDouble(String.Format("{0:0.00}", grandTot));
                        tabFOC.Cells["X" + j].Style.Font.Bold = true;
                        tabFOC.Cells["X" + j].Value = "GRAND TOTAL COST";
                       
                        j += 2;
                        tabFOC.Cells["B" + j + ":AC" + j].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        j++;
                        tabFOC.Cells["D" + j + ":Z" + j].Style.Font.Bold = true;
                        tabFOC.Cells["D" + j + ":Z" + j].Style.Font.Size = 10;
                        tabFOC.Cells["D" + j].Value = "PREPARED BY :";
                        tabFOC.Cells["M" + j].Value = "CHECKED BY:";
                        tabFOC.Cells["V" + j].Value = "ENDORSED BY:";
                        j += 2;
                        tabFOC.Cells["D" + j + ":F" + j + ",L" + j + ":P" + j + ",V" + j + ":X" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        j += 2;
                        tabFOC.Cells["B" + j + ":AC" + j].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

                       
                        var stream = new MemoryStream();
                        xlsFile.SaveAs(stream);

                        string fileName = "BOQReport_" + schemeName + ".xlsx";
                        string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                        stream.Position = 0;
                        return File(stream, contentType, fileName);
                        #endregion
                    }

                    
                    else
                    {
                        #region cost breakdown lama
                        ViewBag.JKH_CONTRACT = jkhContract;
                        ViewBag.date = DateTime.Now.ToString("dd/MM/yyyy");
                        ViewBag.schemeName = schemeName;
                        ViewBag.projNo = "";

                        WebService._base.SUMM_ProjCost estimateBP = new WebService._base.SUMM_ProjCost();
                        estimateBP = myWebService.Get_ProjCost(schemeName);

                        for (int i = 0; i < estimateBP.ProjCost.Count; i++)
                        {
                            ViewBag.SumLbrOT = estimateBP.ProjCost[i].SUM_LBR_OT;
                            ViewBag.SumLbrSalary = estimateBP.ProjCost[i].SUM_LBR_SALARY;
                            ViewBag.SumMaterial = estimateBP.ProjCost[i].SUM_MATERIAL;
                            if (estimateBP.ProjCost[i].SUM_JKH == "0.00")
                            {
                                ViewBag.SumJKH = estimateBP.ProjCost[i].SUM_CONTRACT;
                            }
                            else
                            {
                                ViewBag.SumJKH = estimateBP.ProjCost[i].SUM_JKH;
                            }
                            ViewBag.SumTNT_Mileage = estimateBP.ProjCost[i].SUM_TNT_MILEAGE;
                            ViewBag.SumMilling = estimateBP.ProjCost[i].SUM_MILLING;
                            ViewBag.SumMisc = estimateBP.ProjCost[i].SUM_MISC;
                            ViewBag.SumTotal = estimateBP.ProjCost[i].SUM_TOTAL;
                        }

                        var query = from a in ctxdata.G3E_JOB
                                    where a.SCHEME_NAME == schemeName
                                    select new { a.PLAN_START_DATE, a.PLAN_END_DATE };
                        foreach (var a in query)
                        {
                            ViewBag.StartDate = String.Format("{0:MM/dd/yyyy}", a.PLAN_START_DATE);
                            ViewBag.EndDate = String.Format("{0:MM/dd/yyyy}", a.PLAN_END_DATE);
                        }

                        if (excel_html == "EXCEL")
                        {
                            this.Response.AddHeader("Content-Disposition", "attachment; filename=BOQReport_" + schemeName + ".xls");
                            this.Response.ContentType = "application/vnd.ms-excel";
                        }
                        return View();
                        #endregion
                    }
                }
            }
        }
        #endregion

        public ActionResult ReportPDF(string id)
        {
            string schemeName1 = id;
            return new Rotativa.ActionAsPdf("ReportBOQPDF", new { schemeName1 }) { PageOrientation = Rotativa.Options.Orientation.Landscape, PageSize = Rotativa.Options.Size.A4 };
        }

        public ActionResult ReportBOQPDF(string schemeName1)
        {
            string[] arr = schemeName1.Split('|');
            string schemeName = arr[0];
            string excel_html = arr[1];

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            ViewBag.date = DateTime.Now.ToString("d/M/yyyy");
            ViewBag.schemeName = schemeName;
            //ViewBag.projNo = myWebService.BOQ_ProjNo(schemeName);
            using (Entities ctxdata = new Entities())
            {
                ViewBag.Exc = (from a in ctxdata.WV_EXC_MAST
                               join b in ctxdata.G3E_JOB on a.EXC_ABB.Trim() equals b.EXC_ABB.Trim()
                               where b.SCHEME_NAME == schemeName
                               select a.EXC_NAME).Single();


                ViewBag.projNo = (from b in ctxdata.G3E_JOB
                                  where b.SCHEME_NAME == schemeName
                                  select b.PROJECT_NO).Single();
            }

            WebService._base.OSPBOQ_MAIN_EXCEL BOQ_REPORT = new WebService._base.OSPBOQ_MAIN_EXCEL();
            string checkJKH_Contract = "JKH";
            BOQ_REPORT = myWebService.GetOSPBOQ_MAIN_Excel(0, 100, schemeName);
            int count = BOQ_REPORT.BOQ_MAIN_Excel.Count;
            ViewBag.Page = count / 16;

            WebService._base.OSPBOQ_MAIN_EXCEL[] BOQ_REPORT2 = new WebService._base.OSPBOQ_MAIN_EXCEL[count / 16];
            for (int i = 0; i < (count / 16); i++)
            {
                BOQ_REPORT2[i] = new WebService._base.OSPBOQ_MAIN_EXCEL();
                for (int j = (16 * i); j < 16 * (i + 1); j++)
                {
                    WebView.WebService._base.BOQ_MAIN_EXCEL BOQ_MAIN_EXCEL = new WebView.WebService._base.BOQ_MAIN_EXCEL();
                    BOQ_MAIN_EXCEL.x_TOTAL_PRICE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_TOTAL_PRICE;
                    BOQ_MAIN_EXCEL.x_INSTALL_VALUE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_INSTALL_VALUE;
                    BOQ_MAIN_EXCEL.x_BQ_INSTALL_PRICE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_BQ_INSTALL_PRICE;
                    BOQ_MAIN_EXCEL.x_MATERIAL_VALUE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_MATERIAL_VALUE;
                    BOQ_MAIN_EXCEL.x_BQ_MAT_PRICE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_BQ_MAT_PRICE;
                    BOQ_MAIN_EXCEL.x_PU_UOM = BOQ_REPORT.BOQ_MAIN_Excel[j].x_PU_UOM;
                    BOQ_MAIN_EXCEL.x_BQ_CONT_PRICE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_BQ_CONT_PRICE;
                    BOQ_MAIN_EXCEL.x_CONT_VALUE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_CONT_VALUE;
                    BOQ_MAIN_EXCEL.x_TOTAL_VALUE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_TOTAL_VALUE;
                    BOQ_MAIN_EXCEL.x_PU_DESC = BOQ_REPORT.BOQ_MAIN_Excel[j].x_PU_DESC; // PLANT UNIT
                    BOQ_MAIN_EXCEL.x_CONTRACT_NO = BOQ_REPORT.BOQ_MAIN_Excel[j].x_CONTRACT_NO;
                    BOQ_MAIN_EXCEL.x_CONTRACT_DESC = BOQ_REPORT.BOQ_MAIN_Excel[j].x_CONTRACT_DESC;
                    BOQ_MAIN_EXCEL.x_PU_ID = BOQ_REPORT.BOQ_MAIN_Excel[j].x_PU_ID;
                    BOQ_MAIN_EXCEL.x_ISP_OSP = BOQ_REPORT.BOQ_MAIN_Excel[j].x_ISP_OSP;
                    BOQ_MAIN_EXCEL.x_PU_QTY = BOQ_REPORT.BOQ_MAIN_Excel[j].x_PU_QTY;
                    BOQ_MAIN_EXCEL.x_RATE_INDICATOR = BOQ_REPORT.BOQ_MAIN_Excel[j].x_RATE_INDICATOR;
                    BOQ_MAIN_EXCEL.x_ITEM_NO = BOQ_REPORT.BOQ_MAIN_Excel[j].x_ITEM_NO;
                    BOQ_MAIN_EXCEL.x_NET_PRICE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_NET_PRICE;

                    BOQ_REPORT2[i].BOQ_MAIN_Excel.Add(BOQ_MAIN_EXCEL);
                }
            }

            ViewBag.TotalInstallValue = BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_INSTALL_VALUE_TOTAL;
            ViewBag.TotalMaterialValue = BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_MATERIAL_VALUE_TOTAL;
            ViewBag.TotalContractValue = BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_CONT_VALUE_TOTAL;
            ViewBag.TotalValue = BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_TOTAL_VALUE_TOTAL;

            if (myWebService.CheckBOQErrorPUID(schemeName) && myWebService.CheckBOQErrorMinMaterial(schemeName))
            {
                ViewBag.ErrorMsgInList = "Unmatched Plan Unit and Min Material";
                ViewBag.ErrorMsgInTotal = "WARNING! COST FIGURES ARE INACCURATE";
            }
            else if (myWebService.CheckBOQErrorPUID(schemeName) && myWebService.CheckBOQErrorContractNo(schemeName))
            {
                ViewBag.ErrorMsgInList = "Unmatched Contract No and Min Material";
                ViewBag.ErrorMsgInTotal = "WARNING! COST FIGURES ARE INACCURATE";
            }
            else if (myWebService.CheckBOQErrorPUID(schemeName))
            {
                ViewBag.ErrorMsgInList = "Unmatched Plan Unit";
                ViewBag.ErrorMsgInTotal = "WARNING! COST FIGURES ARE INACCURATE";
            }
            else if (myWebService.CheckBOQErrorContractNo(schemeName))
            {
                ViewBag.ErrorMsgInList = "Unmatched Contract No";
                ViewBag.ErrorMsgInTotal = "WARNING! COST FIGURES ARE INACCURATE";
            }
            else if (myWebService.CheckBOQErrorMinMaterial(schemeName))
            {
                ViewBag.ErrorMsgInList = "Unmatched Min Material";
                ViewBag.ErrorMsgInTotal = "WARNING! COST FIGURES ARE INACCURATE";
            }
            else
            {
                ViewBag.ErrorMsgInList = "";
                ViewBag.ErrorMsgInTotal = "";
            }
            for (int pageNum = 1; pageNum <= (count / 16); pageNum++)
            {
                ViewData["BOQReportData" + pageNum] = BOQ_REPORT2[pageNum - 1].BOQ_MAIN_Excel;
            }
            // this.Response.AddHeader("Content-Disposition", "Report.xls");
            if (excel_html == "EXCEL")
            {
                this.Response.AddHeader("Content-Disposition", "attachment; filename=" + schemeName + ".xls");
                this.Response.ContentType = "application/vnd.ms-excel";
            }
            return View();
        }


        public ActionResult ReportBOQ_Contract(string id)
        {
            string[] arr = id.Split('|');
            string schemeName = arr[0];
            string excel_html = arr[1];

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            ViewBag.date = DateTime.Now.ToString("d/M/yyyy");
            ViewBag.schemeName = schemeName;
           // ViewBag.projNo = "";//myWebService.BOQ_ProjNo(id);
            using (Entities ctxdata = new Entities())
            {
                int countExc = (from a in ctxdata.WV_EXC_MAST
                             join b in ctxdata.G3E_JOB on a.EXC_ABB.Trim() equals b.EXC_ABB.Trim()
                             where b.SCHEME_NAME == schemeName
                             select a.EXC_NAME).Count();

                if (countExc != 0)
                {
                    ViewBag.Exc = (from a in ctxdata.WV_EXC_MAST
                                   join b in ctxdata.G3E_JOB on a.EXC_ABB.Trim() equals b.EXC_ABB.Trim()
                                   where b.SCHEME_NAME == schemeName
                                   select a.EXC_NAME).Single();
                }
                else
                {
                    ViewBag.Exc = (from a in ctxdata.WV_EXC_MAST
                                   join b in ctxdata.WV_ISP_JOB on a.EXC_ABB.Trim() equals b.EXC_ABB.Trim()
                                   where b.SCHEME_NAME == schemeName
                                   select a.EXC_NAME).Single();
                }

                ViewBag.projNo = (from b in ctxdata.G3E_JOB
                                  where b.SCHEME_NAME == schemeName
                                  select b.PROJECT_NO).FirstOrDefault();
            }
            

            WebService._base.OSPBOQ_MAIN_EXCEL BOQ_REPORT = new WebService._base.OSPBOQ_MAIN_EXCEL();
            string checkJKH_Contract = "CONTRACT";
            BOQ_REPORT = myWebService.GetOSPBOQ_MAIN_Excel(0, 100, schemeName);
            int count = BOQ_REPORT.BOQ_MAIN_Excel.Count;
            ViewBag.Page = count/16;

            WebService._base.OSPBOQ_MAIN_EXCEL[] BOQ_REPORT2 = new WebService._base.OSPBOQ_MAIN_EXCEL[count / 16];

            for (int i = 0; i < (count / 16); i++)
            {
                BOQ_REPORT2[i] = new WebService._base.OSPBOQ_MAIN_EXCEL();
                for (int j = (16 * i); j < 16 * (i + 1); j++)
                {
                    WebView.WebService._base.BOQ_MAIN_EXCEL BOQ_MAIN_EXCEL = new WebView.WebService._base.BOQ_MAIN_EXCEL();
                    BOQ_MAIN_EXCEL.x_TOTAL_PRICE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_TOTAL_PRICE;
                    BOQ_MAIN_EXCEL.x_INSTALL_VALUE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_INSTALL_VALUE;
                    BOQ_MAIN_EXCEL.x_BQ_INSTALL_PRICE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_BQ_INSTALL_PRICE;
                    BOQ_MAIN_EXCEL.x_MATERIAL_VALUE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_MATERIAL_VALUE;
                    BOQ_MAIN_EXCEL.x_BQ_MAT_PRICE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_BQ_MAT_PRICE;
                    BOQ_MAIN_EXCEL.x_PU_UOM = BOQ_REPORT.BOQ_MAIN_Excel[j].x_PU_UOM;
                    BOQ_MAIN_EXCEL.x_BQ_CONT_PRICE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_BQ_CONT_PRICE;
                    BOQ_MAIN_EXCEL.x_CONT_VALUE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_CONT_VALUE;
                    BOQ_MAIN_EXCEL.x_TOTAL_VALUE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_TOTAL_VALUE;
                    BOQ_MAIN_EXCEL.x_PU_DESC = BOQ_REPORT.BOQ_MAIN_Excel[j].x_PU_DESC; // PLANT UNIT
                    BOQ_MAIN_EXCEL.x_CONTRACT_NO = BOQ_REPORT.BOQ_MAIN_Excel[j].x_CONTRACT_NO;
                    BOQ_MAIN_EXCEL.x_CONTRACT_DESC = BOQ_REPORT.BOQ_MAIN_Excel[j].x_CONTRACT_DESC;
                    BOQ_MAIN_EXCEL.x_PU_ID = BOQ_REPORT.BOQ_MAIN_Excel[j].x_PU_ID;
                    BOQ_MAIN_EXCEL.x_ISP_OSP = BOQ_REPORT.BOQ_MAIN_Excel[j].x_ISP_OSP;
                    BOQ_MAIN_EXCEL.x_PU_QTY = BOQ_REPORT.BOQ_MAIN_Excel[j].x_PU_QTY;
                    BOQ_MAIN_EXCEL.x_RATE_INDICATOR = BOQ_REPORT.BOQ_MAIN_Excel[j].x_RATE_INDICATOR;
                    BOQ_MAIN_EXCEL.x_ITEM_NO = BOQ_REPORT.BOQ_MAIN_Excel[j].x_ITEM_NO;
                    BOQ_MAIN_EXCEL.x_NET_PRICE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_NET_PRICE;

                    BOQ_REPORT2[i].BOQ_MAIN_Excel.Add(BOQ_MAIN_EXCEL);
                }
            }
            ViewBag.TotalInstallValue = BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_INSTALL_VALUE_TOTAL;
            ViewBag.TotalMaterialValue = BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_MATERIAL_VALUE_TOTAL;
            ViewBag.TotalContractValue = BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_CONT_VALUE_TOTAL;
            ViewBag.TotalValue = BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_CONT_VALUE_TOTAL;

            if (myWebService.CheckBOQErrorPUID(schemeName) && myWebService.CheckBOQErrorMinMaterial(schemeName))
            {
                ViewBag.ErrorMsgInList = "Unmatched Plan Unit and Min Material";
                ViewBag.ErrorMsgInTotal = "WARNING! COST FIGURES ARE INACCURATE";
            }
            else if (myWebService.CheckBOQErrorPUID(schemeName))
            {
                ViewBag.ErrorMsgInList = "Unmatched Plan Unit";
                ViewBag.ErrorMsgInTotal = "WARNING! COST FIGURES ARE INACCURATE";
            }
            else if (myWebService.CheckBOQErrorMinMaterial(schemeName))
            {
                ViewBag.ErrorMsgInList = "Unmatched Min Material";
                ViewBag.ErrorMsgInTotal = "WARNING! COST FIGURES ARE INACCURATE";
            }
            else
            {
                ViewBag.ErrorMsgInList = "";
                ViewBag.ErrorMsgInTotal = "";
            }

            for (int pageNum = 1; pageNum <= (count / 16); pageNum++)
            {
                ViewData["BOQReportData" + pageNum] = BOQ_REPORT2[pageNum - 1].BOQ_MAIN_Excel;
            }
            //string vB = "BOQReportData1";
            //ViewData["BOQReportData"] = BOQ_REPORT.BOQ_MAIN_Excel;
            // this.Response.AddHeader("Content-Disposition", "Report.xls");
            if (excel_html == "EXCEL")
            {
                this.Response.AddHeader("Content-Disposition", "attachment; filename=" + schemeName + ".xls");
                this.Response.ContentType = "application/vnd.ms-excel";
            }
            return View();
        }

        public ActionResult ReportProjectCostBreakdownOld(string id)
        {
            string[] arr = id.Split('|');
            string schemeName = arr[0];
            string excel_html = arr[1];
            string jkhContract = arr[2];

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

          
            ViewBag.JKH_CONTRACT = jkhContract;
            ViewBag.date = DateTime.Now.ToString("dd/MM/yyyy");
            ViewBag.schemeName = schemeName;
            ViewBag.projNo = "";//myWebService.BOQ_ProjNo(id);

            WebService._base.SUMM_ProjCost estimateBP = new WebService._base.SUMM_ProjCost();
            estimateBP = myWebService.Get_ProjCost(schemeName);

            for (int i = 0; i < estimateBP.ProjCost.Count; i++)
            {
                ViewBag.SumLbrOT = estimateBP.ProjCost[i].SUM_LBR_OT;
                ViewBag.SumLbrSalary = estimateBP.ProjCost[i].SUM_LBR_SALARY;
                ViewBag.SumMaterial = estimateBP.ProjCost[i].SUM_MATERIAL;
                if (estimateBP.ProjCost[i].SUM_JKH == "0.00")
                {
                    //ViewBag.SumJKH = "0.00";
                    ViewBag.SumJKH = estimateBP.ProjCost[i].SUM_CONTRACT;
                }
                else
                {
                    //ViewBag.SumContract = estimateBP.ProjCost[i].SUM_CONTRACT;
                    ViewBag.SumJKH = estimateBP.ProjCost[i].SUM_JKH;
                }
                ViewBag.SumTNT_Mileage = estimateBP.ProjCost[i].SUM_TNT_MILEAGE;
                ViewBag.SumMilling = estimateBP.ProjCost[i].SUM_MILLING;
                ViewBag.SumMisc = estimateBP.ProjCost[i].SUM_MISC;
                ViewBag.SumTotal = estimateBP.ProjCost[i].SUM_TOTAL;
                //if (estimateBP.ProjCost[i].SUM_CONTRACT == "0.00")
                //{
                //    ViewBag.SUM_CONTRACT = "";
                //}
                //else
                //{
                //    ViewBag.SumJKH = estimateBP.ProjCost[i].SUM_JKH;
                    //ViewBag.SumContract = estimateBP.ProjCost[i].SUM_CONTRACT;
                //}
               
            }

            using (Entities ctxdata = new Entities())
            {
                var query = from a in ctxdata.G3E_JOB
                            where a.SCHEME_NAME == schemeName
                            select new { a.PLAN_START_DATE, a.PLAN_END_DATE };
                foreach(var a in query)
                {
                    ViewBag.StartDate = String.Format("{0:MM/dd/yyyy}",a.PLAN_START_DATE);
                    ViewBag.EndDate = String.Format("{0:MM/dd/yyyy}", a.PLAN_END_DATE);
                }
            }

            if (excel_html == "EXCEL")
            {
                this.Response.AddHeader("Content-Disposition", "attachment; filename="+schemeName+".xls");
                this.Response.ContentType = "application/vnd.ms-excel";
            }
            return View();
        }

        public ActionResult ReportBussinessPlanEstimate(string id)
        {
            string[] arr = id.Split('/');
            string schemeName = arr[0];
            string excel_html = arr[1];

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            ViewBag.date = DateTime.Now.ToString("dd/MM/yyyy");
            ViewBag.schemeName = schemeName;
            ViewBag.projNo = "";//myWebService.BOQ_ProjNo(id);

            WebService._base.Estimate_BPlanEstimate estimateBP = new WebService._base.Estimate_BPlanEstimate();
            estimateBP = myWebService.GetBussinessPlan_Estimate(schemeName);

            for (int i = 0; i < estimateBP.BPlanEstimate.Count; i++)
            {
                ViewBag.EstBPJamBuruh = estimateBP.BPlanEstimate[i].BP_JAM_BURUH;
                ViewBag.EstBPJamNilai = estimateBP.BPlanEstimate[i].BP_JAM_NILAI;
                ViewBag.EstBPOTBuruh = estimateBP.BPlanEstimate[i].BP_OT_BURUH;
                ViewBag.EstBPOTNilai = estimateBP.BPlanEstimate[i].BP_OT_NILAI;
                ViewBag.EstBPJKHBahan = estimateBP.BPlanEstimate[i].BP_JKH_BAHAN;
                ViewBag.EstBPJKHPelbagai = estimateBP.BPlanEstimate[i].BP_JKH_PEL;
                ViewBag.EstBPTMBahan = estimateBP.BPlanEstimate[i].BP_TM_BAHAN;
                ViewBag.EstBPTMPelbagai = estimateBP.BPlanEstimate[i].BP_TM_PEL;
                ViewBag.EstBPTotalBuruh = estimateBP.BPlanEstimate[i].BP_TOTAL_BURUH;
                ViewBag.EstBPTotalNilai = estimateBP.BPlanEstimate[i].BP_TOTAL_NILAI;
                ViewBag.EstBPTotalBahan = estimateBP.BPlanEstimate[i].BP_TOTAL_BAHAN;
                ViewBag.EstBPTotalPelbagai = estimateBP.BPlanEstimate[i].BP_TOTAL_PELBAGAI;
                ViewBag.EstBPTotalJKH = estimateBP.BPlanEstimate[i].BP_TOTAL_JKH;
                ViewBag.EstBPTotalTM = estimateBP.BPlanEstimate[i].BP_TOTAL_TM;
                ViewBag.EstBPTotalKos = estimateBP.BPlanEstimate[i].BP_TOTAL_KOS;
                ViewBag.EstBPTotalTotal = estimateBP.BPlanEstimate[i].BP_TOTAL_TOTAL;
                ViewBag.EstBPStartDate = estimateBP.BPlanEstimate[i].BP_START_DATE;
                ViewBag.EstBPEndDate = estimateBP.BPlanEstimate[i].BP_END_DATE;
            }
            if (excel_html == "EXCEL")
            {
                this.Response.AddHeader("Content-Disposition", "attachment; filename=report.xls");
                this.Response.ContentType = "application/vnd.ms-excel";
            }
            return View();
        }

        public ActionResult ReportMaterialSupply(string id)
        {
            string[] arr = id.Split('|');
            string schemeName = arr[0];
            string excel_html = arr[1];

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            ViewBag.date = DateTime.Now.ToString("d/M/yyyy");
            ViewBag.schemeName = schemeName;
            System.Diagnostics.Debug.WriteLine("SCHEME: " + schemeName);
            using (Entities ctxdata = new Entities())
            {
                ViewBag.projNo = (from b in ctxdata.G3E_JOB
                                  where b.SCHEME_NAME == schemeName
                                  select b.PROJECT_NO).Single();

                ViewBag.Exc = (from a in ctxdata.WV_EXC_MAST
                               join b in ctxdata.G3E_JOB on a.EXC_ABB.Trim() equals b.EXC_ABB.Trim()
                               where b.SCHEME_NAME == schemeName
                               select a.EXC_NAME).Single();

                string projNo1 = ViewBag.projNo;

                ViewBag.projDesc = (from a in ctxdata.WV_GEM_PROJNO
                                    where a.PROJECT_NO.Trim() == projNo1
                                    select a.PROJ_DESC).FirstOrDefault();
            }

            WebService._base.OSPBOQ_MAIN_EXCEL BOQ_REPORT = new WebService._base.OSPBOQ_MAIN_EXCEL();
            string checkJKH_Contract = "JKH";
            BOQ_REPORT = myWebService.GetOSPBOQ_MAT_Excel(0, 100, schemeName);
            System.Diagnostics.Debug.WriteLine("REPORT: " + BOQ_REPORT);
            int count = BOQ_REPORT.BOQ_MAIN_Excel.Count;
            ViewBag.Page = count / 16;

            WebService._base.OSPBOQ_MAIN_EXCEL[] BOQ_REPORT2 = new WebService._base.OSPBOQ_MAIN_EXCEL[count / 16];
            for (int i = 0; i < (count / 16); i++)
            {
                BOQ_REPORT2[i] = new WebService._base.OSPBOQ_MAIN_EXCEL();
                for (int j = (16 * i); j < 16 * (i + 1); j++)
                {
                    WebView.WebService._base.BOQ_MAIN_EXCEL BOQ_MAIN_EXCEL = new WebView.WebService._base.BOQ_MAIN_EXCEL();
                    BOQ_MAIN_EXCEL.x_TOTAL_PRICE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_TOTAL_PRICE;
                    BOQ_MAIN_EXCEL.x_MATERIAL_VALUE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_MATERIAL_VALUE;
                    BOQ_MAIN_EXCEL.x_BQ_MAT_PRICE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_BQ_MAT_PRICE;
                    BOQ_MAIN_EXCEL.x_PU_UOM = BOQ_REPORT.BOQ_MAIN_Excel[j].x_PU_UOM;
                    BOQ_MAIN_EXCEL.x_TOTAL_VALUE = BOQ_REPORT.BOQ_MAIN_Excel[j].x_TOTAL_VALUE;
                    BOQ_MAIN_EXCEL.x_PU_DESC = BOQ_REPORT.BOQ_MAIN_Excel[j].x_PU_DESC; 
                    BOQ_MAIN_EXCEL.x_PU_ID = BOQ_REPORT.BOQ_MAIN_Excel[j].x_PU_ID;
                    BOQ_MAIN_EXCEL.x_ISP_OSP = BOQ_REPORT.BOQ_MAIN_Excel[j].x_ISP_OSP;

                    BOQ_REPORT2[i].BOQ_MAIN_Excel.Add(BOQ_MAIN_EXCEL);
                }
            }

            //ViewBag.TotalInstallValue = BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_INSTALL_VALUE_TOTAL;
            //ViewBag.TotalMaterialValue = BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_MATERIAL_VALUE_TOTAL;
            //ViewBag.TotalContractValue = BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_CONT_VALUE_TOTAL;
            if (count > 0)
            {
                ViewBag.TotalValue = BOQ_REPORT.BOQ_MAIN_Excel[count - 1].x_TOTAL_VALUE_TOTAL;
                for (int pageNum = 1; pageNum <= (count / 16); pageNum++)
                {
                    ViewData["BOQReportData" + pageNum] = BOQ_REPORT2[pageNum - 1].BOQ_MAIN_Excel;
                }
            }
            // this.Response.AddHeader("Content-Disposition", "Report.xls");
            if (excel_html == "EXCEL")
            {
                this.Response.AddHeader("Content-Disposition", "attachment; filename=" + schemeName + ".xls");
                this.Response.ContentType = "application/vnd.ms-excel";
            }
            return View();
        }

        public ActionResult UpdateDataLWH(string schemeName, string txtJBuruh, string txtJBiasa, string txtJKanan, string txtPTBiasa)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.Estimate newLWH = new WebService._base.Estimate();
            newLWH.LAB_BURUH = txtJBuruh;
            newLWH.LAB_BIASA = txtJBiasa;
            newLWH.LAB_KANAN = txtJKanan;
            newLWH.LAB_PEMBANTU = txtPTBiasa;
            success = myWebService.UpdateDataLWH(newLWH, schemeName);

            WebService._base.Estimate_Lab estimateLab = new WebService._base.Estimate_Lab();
            estimateLab = myWebService.GetEstimate_Lab(schemeName);

            string EstLabTotal = "";
            for (int i = 0; i < estimateLab.LabList.Count; i++)
            {

                EstLabTotal = estimateLab.LabList[i].LAB_EST_TOTAL;
            }

            return Json(new
            {
                txttotal = EstLabTotal,
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateDataSUM(string schemeName, string txtSumLbrOT, string txtSumLbrSalary, string txtSumTNT_Mileage, string txtSumMilling, string txtSumMisc)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.Estimate newSumm = new WebService._base.Estimate();
            newSumm.SUM_LBR_OT = txtSumLbrOT;
            newSumm.SUM_LBR_SALARY = txtSumLbrSalary;
            newSumm.SUM_MILLING = txtSumMilling;
            newSumm.SUM_MISC = txtSumMisc;
            newSumm.SUM_TNT_MILEAGE = txtSumTNT_Mileage;

            success = myWebService.UpdateDataSUM(newSumm, schemeName);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
            
        }

        public ActionResult UpdateMillageLabour(string schemeName, string millageLabour)
        {
            bool success = true;
            string[] arr = millageLabour.Split('|');
            string andTnt = arr[0];
            string monthSalary = arr[1];
            string totalOT = arr[2];
            int mileageOrLabour = 0;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.Estimate newSumm = new WebService._base.Estimate();
            if (totalOT == "^")
            {
                totalOT = String.Format("{0:0.00}", (0.00));
                mileageOrLabour = 1;
            }
            else
                totalOT = String.Format("{0:0.00}", totalOT);
            if (monthSalary == "^")
                monthSalary = String.Format("{0:0.00}", (0.00));
            else
                monthSalary = String.Format("{0:0.00}", monthSalary);
            if (andTnt == "^")
            {
                andTnt = String.Format("{0:0.00}", (0.00));
                mileageOrLabour = 2;
            }
            else
                andTnt = String.Format("{0:0.00}", andTnt);
            
            newSumm.SUM_LBR_OT = totalOT;
            newSumm.SUM_LBR_SALARY = monthSalary;
            newSumm.SUM_TNT_MILEAGE = andTnt;

            success = myWebService.UpdateMillageLabour(newSumm, schemeName, mileageOrLabour);

            return Json(new
            {
                Success = success,
                totalOT = totalOT,
                monthSalary = monthSalary,
                andTnt = andTnt
            }, JsonRequestBehavior.AllowGet);

        }

        public ActionResult UpdateDataMaintenance(string schemeName, string PU_ID, string PU_BILLRATE, string txtEstMainPUDesc, string txtEstMainPUQty, string txtEstMainOldMatPr)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.Estimate updateMain = new WebService._base.Estimate();
            updateMain.MAIN_PU_DESC = txtEstMainPUDesc;
            updateMain.MAIN_PU_QTY = txtEstMainPUQty;
            updateMain.MAIN_OLD_MAT_PR = txtEstMainOldMatPr;
            success = myWebService.UpdateDataMaintenance(updateMain, schemeName, PU_ID, PU_BILLRATE);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);

        }

        public ActionResult UpdateDataMaintenanceContract(string schemeName, string contractNo, string itemNo, string qty)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.UpdateDataMaintenanceContract(schemeName, contractNo, itemNo, qty);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);

        }

        public ActionResult AddDataMaintenance(string schemeName, string PU_ID, string PU_BILLRATE, string txtEstMainPUDesc, string txtEstMainPUQty, string txtEstMainOldMatPr, string txtEstMainMatPrice, string txtEstMainInstPrice)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.Estimate updateMain = new WebService._base.Estimate();
            updateMain.MAIN_PU_DESC = txtEstMainPUDesc;
            updateMain.MAIN_PU_QTY = txtEstMainPUQty;
            updateMain.MAIN_OLD_MAT_PR = txtEstMainOldMatPr;
            updateMain.MAIN_PU_INST_PR = txtEstMainInstPrice;
            updateMain.MAIN_PU_MAT_PR = txtEstMainMatPrice;

            success = myWebService.AddDataMaintenance(updateMain, schemeName, PU_ID, PU_BILLRATE);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);

        }

        public ActionResult AddDataMaintenanceContract(string schemeName, string ContractNo, string ItemNo, string txtQuantity, string txtEstMulItemNo)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            
            System.Diagnostics.Debug.WriteLine(schemeName + " : " + ContractNo + " : " + ItemNo + " : "+ txtQuantity);
            success = myWebService.AddDataMaintenanceContract(schemeName, ContractNo, ItemNo, txtQuantity);

            string[] check = txtEstMulItemNo.Split(','); // check item no
            string[] arr = txtEstMulItemNo.Split(',');

            if (ItemNo == "Multiple")
            {
                for (int i = 0; i < check.Length; i++)
                {
                    success = myWebService.CheckItemContractBOQ(schemeName, arr[i]);
                }
               
                if (success == true)
                { // if item no doesn't exist
                    for (int i = 0; i < arr.Length; i++)
                    {
                        success = myWebService.AddDataMaintenanceContract(schemeName, ContractNo, arr[i], txtQuantity);
                    }
                }
            }
            

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);

        }

        public ActionResult UpdateDataIncCost(string schemeName, string txtEstIncOtSupHours, string txtEstIncValue, string txtEstIncOtValue, string txtEstIncExecMileage, string txtEstIncNonExecMileage)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.Estimate newIncCost = new WebService._base.Estimate();
            newIncCost.INC_CTRT_VALUE = txtEstIncValue;
            newIncCost.INC_CTRT_OT_SUP_HOURS = txtEstIncOtSupHours;
            newIncCost.INC_CTRT_OT_VALUE = txtEstIncOtValue;
            newIncCost.INC_EXEC_MILEAGE = txtEstIncExecMileage;
            newIncCost.INC_NON_EXEC_MILEAGE = txtEstIncNonExecMileage;
            success = myWebService.UpdateDataINC(newIncCost, schemeName);

            WebService._base.Estimate_Inc estimateInc = new WebService._base.Estimate_Inc();
            estimateInc = myWebService.GetEstimate_Inc(schemeName);

            string EstIncTotalHours = "";
            string EstIncTotalValue = "";
            string EstIncTotalMileage = "";
            for (int i = 0; i < estimateInc.IncList.Count; i++)
            {

                EstIncTotalHours = estimateInc.IncList[i].INC_TOTAL_HOURS;
                EstIncTotalValue = estimateInc.IncList[i].INC_TOTAL_VALUE;
                EstIncTotalMileage = estimateInc.IncList[i].INC_TOTAL_MILEAGE;
            }


            return Json(new
            {
                txtEstIncTotalHours = EstIncTotalHours,
                txtEstIncTotalValue = EstIncTotalValue,
                txtEstIncTotalMileage = EstIncTotalMileage,
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateDataBussinessPlanEstimate(string schemeName, string txtEstBPJamBuruh, string txtEstBPJamNilai, string txtEstBPOTBuruh, string txtEstBPOTNilai, string txtEstBPTMPelbagai)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.Estimate newBP = new WebService._base.Estimate();
            newBP.BP_JAM_BURUH = txtEstBPJamBuruh;
            newBP.BP_JAM_NILAI = txtEstBPJamNilai;
            newBP.BP_OT_BURUH = txtEstBPOTBuruh;
            newBP.BP_OT_NILAI = txtEstBPOTNilai;
            newBP.BP_TM_PEL = txtEstBPTMPelbagai;
            success = myWebService.UpdateDataBPlanEstimate(newBP, schemeName);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult AddEst_Maintenance(string PU_BILLRATE, string schemeName, string txtEstMainPUDesc, string txtEstMainMatPrice, string txtEstMainInstPrice, string txtEstMainPUQty, string txtEstMainOldMatPr, string txtEstMainOldInstPr, string txtEstMainConstructBy, string txtEstMainRecvrQty)
        {
            string[] arr = PU_BILLRATE.Split('/');
            string txtEstMainPUId = arr[0];
            string txtEstMainBillRate = arr[1];

            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            int check = 0;
            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.WV_MAIN_DATA
                            where p.SCHEME_NAME == schemeName
                            select new { Text = p.PU_ID + "/" + p.BILL_RATE };

                foreach (var a in query)
                {
                    if (check == 0)
                    {
                        if (a.Text == PU_BILLRATE)
                            check = 1;
                    }
                }
            }
            if (check == 0)
            {
                WebService._base.Estimate newEstMain = new WebService._base.Estimate();
                newEstMain.SCHEME_NAME = schemeName;
                newEstMain.MAIN_PU_ID = txtEstMainPUId;
                newEstMain.MAIN_PU_DESC = txtEstMainPUDesc;
                newEstMain.MAIN_PU_MAT_PR = txtEstMainMatPrice;
                newEstMain.MAIN_PU_INST_PR = txtEstMainInstPrice;
                newEstMain.MAIN_PU_QTY = txtEstMainPUQty;
                newEstMain.MAIN_OLD_MAT_PR = txtEstMainOldMatPr;
                newEstMain.MAIN_OLD_INSTALL_PR = txtEstMainOldInstPr;
                newEstMain.MAIN_CONSTRUCT_BY = txtEstMainConstructBy;
                newEstMain.MAIN_RECVR_QTY = txtEstMainRecvrQty;
                newEstMain.MAIN_BILL_RATE = txtEstMainBillRate;

                success = myWebService.AddEstimate_Main(newEstMain);
            }
            return Json(new
            {
                Success = success,
                check = check
            }, JsonRequestBehavior.AllowGet);
        }
    }
}
