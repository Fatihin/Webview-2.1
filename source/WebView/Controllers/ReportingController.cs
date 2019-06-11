using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PagedList;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;

namespace WebView.Controllers
{
    public class ReportingController : Controller
    {
        //
        // GET: /Reporting/

        
        public ActionResult NetworkInfraReport(string reportType, string ptt, string excabb, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.DataNetInfraInfo NetInfraInfo1 = new WebService._base.DataNetInfraInfo();

            if (reportType != null || ptt != null)
            {
                if (reportType == "" || ptt == "")
                {
                    NetInfraInfo1 = myWebService.GetDataNetInfraInfo(null, null,null);
                    ViewBag.ReportType1 = reportType;
                    ViewBag.Ptt1 = ptt;
                    ViewBag.excabb1 = excabb;
                }
                else
                {
                    NetInfraInfo1 = myWebService.GetDataNetInfraInfo(reportType, ptt,excabb);
                    ViewBag.ReportType1 = reportType;
                    ViewBag.Ptt1 = ptt;
                    ViewBag.excabb1 = excabb;
                }
            }
            else
            {
                NetInfraInfo1 = myWebService.GetDataNetInfraInfo(null, null,null);
                ViewBag.ReportType1 = "";
                ViewBag.Ptt1 = "";
                ViewBag.excabb1 = "";
            }

            using (Entities1 ctxData1 = new Entities1())
            {
                List<SelectListItem> list1 = new List<SelectListItem>();
                var query1 = from p in ctxData1.REF_RPT_INFRA
                             select new { p.DESCRIPTION };
                list1.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query1.OrderBy(it => it.DESCRIPTION))
                {
                    list1.Add(new SelectListItem() { Text = a.DESCRIPTION, Value = a.DESCRIPTION });
                }
                ViewBag.ReportType = list1;
            }

            #region FAZ change table 17012019
            using (Entities1 ctxData2 = new Entities1())
            {
                List<SelectListItem> list2 = new List<SelectListItem>();
                //var query2 = from p in ctxData.WV_PTT_MAST
                //             select new { p.PTT_ID };
                var query2 = from p in ctxData2.REF_RPT_WV_PTT
                             select new { p.PTT };
                list2.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query2.OrderBy(it => it.PTT))
                {
                    list2.Add(new SelectListItem() { Text = a.PTT, Value = a.PTT });
                }
                ViewBag.Ptt = list2;
            }
            #endregion

            #region Fatihin 1102019 - add excabb
            using (Entities ctxData = new Entities())
            {
                List<SelectListItem> list3 = new List<SelectListItem>();
                var query = from p in ctxData.WV_EXC_MAST
                            where p.PTT_ID.Trim() == ptt.Trim()
                            select new { Text = p.EXC_NAME + " (" + p.EXC_ABB + ")", Value = p.EXC_ABB };

                foreach (var a in query)
                {
                    list3.Add(new SelectListItem() { Text = a.Text, Value = a.Value.ToString() });
                }

                ViewBag.excabb = list3;
            }
            #endregion

            int pageSize = 10;
            int pageNumber = (page ?? 1);
            return View(NetInfraInfo1.DataNetInfraInfoList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult GenerateReport(string reportType, string ptt, string excabb)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.DataNetInfraInfo NetInfraInfo1 = new WebService._base.DataNetInfraInfo();
            NetInfraInfo1 = myWebService.GetDataNetInfraInfo(reportType, ptt, excabb);

            using (var xlsFile = new ExcelPackage())
            {
                #region CABINET LATLONG
                if (reportType == "CABINET LATLONG")
                {
                    var cabLatLongTable = xlsFile.Workbook.Worksheets.Add("CABINET LATLONG");
                    cabLatLongTable.Cells.Style.Font.Size = 8;
                    cabLatLongTable.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cabLatLongTable.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    cabLatLongTable.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cabLatLongTable.Cells["A1:E1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    cabLatLongTable.Cells["A1:E1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    cabLatLongTable.Cells["A1:E1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    cabLatLongTable.Cells["A1:E1"].Style.Font.Bold = true;

                    cabLatLongTable.Cells["A1"].Value = "PTT";
                    cabLatLongTable.Cells["B1"].Value = "EXC_ABB";
                    cabLatLongTable.Cells["C1"].Value = "CAB_CODE";
                    cabLatLongTable.Cells["D1"].Value = "LATITUDE";
                    cabLatLongTable.Cells["E1"].Value = "LONGITUDE";
                    
                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        cabLatLongTable.Cells["A" + num + ":E" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        cabLatLongTable.Cells["A" + num + ":E" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        cabLatLongTable.Cells["A" + num].Value = a.PTT;
                        cabLatLongTable.Cells["B" + num].Value = a.EXC_ABB;
                        cabLatLongTable.Cells["C" + num].Value = a.CAB_CODE;
                        cabLatLongTable.Cells["D" + num].Value = a.LATITUDE;
                        cabLatLongTable.Cells["E" + num++].Value = a.LONGITUDE;
                    }
                    cabLatLongTable.Cells.AutoFitColumns();
                }
                #endregion

                #region DP LATLONG
                else if (reportType == "DP LATLONG")
                {
                    var dpLatLongTable = xlsFile.Workbook.Worksheets.Add("DP LATLONG");
                    dpLatLongTable.Cells.Style.Font.Size = 8;
                    dpLatLongTable.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    dpLatLongTable.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    dpLatLongTable.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    dpLatLongTable.Cells["A1:F1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    dpLatLongTable.Cells["A1:F1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    dpLatLongTable.Cells["A1:F1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    dpLatLongTable.Cells["A1:F1"].Style.Font.Bold = true;

                    dpLatLongTable.Cells["A1"].Value = "PTT";
                    dpLatLongTable.Cells["B1"].Value = "EXC_ABB";
                    dpLatLongTable.Cells["C1"].Value = "CAB_CODE";
                    dpLatLongTable.Cells["D1"].Value = "DP_CODE";
                    dpLatLongTable.Cells["E1"].Value = "LATITUDE";
                    dpLatLongTable.Cells["F1"].Value = "LONGITUDE";

                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        dpLatLongTable.Cells["A" + num + ":F" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        dpLatLongTable.Cells["A" + num + ":F" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        dpLatLongTable.Cells["A" + num].Value = a.PTT;
                        dpLatLongTable.Cells["B" + num].Value = a.EXC_ABB;
                        dpLatLongTable.Cells["C" + num].Value = a.CAB_CODE;
                        dpLatLongTable.Cells["D" + num].Value = a.DP_CODE;
                        dpLatLongTable.Cells["E" + num].Value = a.LATITUDE;
                        dpLatLongTable.Cells["F" + num++].Value = a.LONGITUDE;
                    }
                    dpLatLongTable.Cells.AutoFitColumns();
                }
                #endregion

                #region CABINET DISTANCE FROM EXCHANGE
                else if (reportType == "CABINET DISTANCE FROM EXCHANGE")
                {
                    var cabDistTable = xlsFile.Workbook.Worksheets.Add("CABINET DISTANCE FROM EXCHANGE");
                    cabDistTable.Cells.Style.Font.Size = 12;
                    cabDistTable.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cabDistTable.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    cabDistTable.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cabDistTable.Cells["A2:E2"].Style.Fill.BackgroundColor.SetColor(Color.OrangeRed);
                    cabDistTable.Cells["A2:E2"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    cabDistTable.Cells["A2:E2"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    cabDistTable.Cells["A2:E2"].Style.Font.Bold = true;

                    cabDistTable.Cells["A1:E1"].Merge = true;
                    cabDistTable.Cells["A1"].Value = "CABINET DISTANCE FROM EXCHANGE";
                    cabDistTable.Cells["A2"].Value = "PTT";
                    cabDistTable.Cells["B2"].Value = "EXC_ABB";
                    cabDistTable.Cells["C2"].Value = "CABINET TYPE";
                    cabDistTable.Cells["D2"].Value = "CABINET CODE";
                    cabDistTable.Cells["E2"].Value = "CABINET DISTANCE(m)";

                    int num = 3;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        cabDistTable.Cells["A" + num + ":E" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        cabDistTable.Cells["A" + num + ":E" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        cabDistTable.Cells["A" + num].Value = a.PTT;
                        cabDistTable.Cells["B" + num].Value = a.EXC_ABB;
                        cabDistTable.Cells["C" + num].Value = a.CAB_TYPE;
                        cabDistTable.Cells["D" + num].Value = a.CAB_CODE;
                        cabDistTable.Cells["E" + num++].Value = a.CAB_DISTANCE;
                    }
                    cabDistTable.Cells.AutoFitColumns();
                }
                #endregion

                #region DP DISTANCE FROM CABINET
                else if (reportType == "DP DISTANCE FROM CABINET")
                {
                    var dpDistTable = xlsFile.Workbook.Worksheets.Add("DP DISTANCE FROM CABINET");
                    dpDistTable.Cells.Style.Font.Size = 8;
                    dpDistTable.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    dpDistTable.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    dpDistTable.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    dpDistTable.Cells["A1:E1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    dpDistTable.Cells["A1:E1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    dpDistTable.Cells["A1:E1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    dpDistTable.Cells["A1:E1"].Style.Font.Bold = true;

                    dpDistTable.Cells["A1"].Value = "PTT";
                    dpDistTable.Cells["B1"].Value = "EXC_ABB";
                    dpDistTable.Cells["C1"].Value = "CAB_CODE";
                    dpDistTable.Cells["D1"].Value = "DP_CODE";
                    dpDistTable.Cells["E1"].Value = "DISTANCE";

                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        dpDistTable.Cells["A" + num + ":E" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        dpDistTable.Cells["A" + num + ":E" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        dpDistTable.Cells["A" + num].Value = a.PTT;
                        dpDistTable.Cells["B" + num].Value = a.EXC_ABB;
                        dpDistTable.Cells["C" + num].Value = a.CAB_CODE;
                        dpDistTable.Cells["D" + num].Value = a.DP_CODE;
                        dpDistTable.Cells["E" + num++].Value = a.DISTANCE;
                    }
                    dpDistTable.Cells.AutoFitColumns();
                }
                #endregion

                #region MANHOLE LATLONG
                else if (reportType == "MANHOLE LATLONG")
                {
                    var manholeLatLongTable = xlsFile.Workbook.Worksheets.Add("MANHOLE LATLONG");
                    manholeLatLongTable.Cells.Style.Font.Size = 8;
                    manholeLatLongTable.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    manholeLatLongTable.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    manholeLatLongTable.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    manholeLatLongTable.Cells["A1:J1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    manholeLatLongTable.Cells["A1:J1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    manholeLatLongTable.Cells["A1:J1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    manholeLatLongTable.Cells["A1:J1"].Style.Font.Bold = true;

                    manholeLatLongTable.Cells["A1"].Value = "PTT";
                    manholeLatLongTable.Cells["B1"].Value = "EXC_ABB";
                    manholeLatLongTable.Cells["C1"].Value = "STATE";
                    manholeLatLongTable.Cells["D1"].Value = "REGION";
                    manholeLatLongTable.Cells["E1"].Value = "ZONE";
                    manholeLatLongTable.Cells["F1"].Value = "MANHOLE_FID";
                    manholeLatLongTable.Cells["G1"].Value = "MANHOLE_ID";
                    manholeLatLongTable.Cells["H1"].Value = "MANHOLE_TYPE";
                    manholeLatLongTable.Cells["I1"].Value = "LATITUDE";
                    manholeLatLongTable.Cells["J1"].Value = "LONGITUDE";

                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        manholeLatLongTable.Cells["A" + num + ":J" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        manholeLatLongTable.Cells["A" + num + ":J" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        manholeLatLongTable.Cells["A" + num].Value = a.PTT;
                        manholeLatLongTable.Cells["B" + num].Value = a.EXC_ABB;
                        manholeLatLongTable.Cells["C" + num].Value = a.STATE;
                        manholeLatLongTable.Cells["D" + num].Value = a.REGION;
                        manholeLatLongTable.Cells["E" + num].Value = a.ZONE;
                        manholeLatLongTable.Cells["F" + num].Value = a.MANHOLE_FID;
                        manholeLatLongTable.Cells["G" + num].Value = a.MANHOLE_ID;
                        manholeLatLongTable.Cells["H" + num].Value = a.MANHOLE_TYPE;
                        manholeLatLongTable.Cells["I" + num].Value = a.LATITUDE;
                        manholeLatLongTable.Cells["J" + num++].Value = a.LONGITUDE;
                    }
                    manholeLatLongTable.Cells.AutoFitColumns();
                }
                #endregion

                #region POLE LATLONG
                else if (reportType == "POLE LATLONG")
                {
                    var poleLatLongTable = xlsFile.Workbook.Worksheets.Add("POLE LATLONG");
                    poleLatLongTable.Cells.Style.Font.Size = 8;
                    poleLatLongTable.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    poleLatLongTable.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    poleLatLongTable.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    poleLatLongTable.Cells["A1:I1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    poleLatLongTable.Cells["A1:I1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    poleLatLongTable.Cells["A1:I1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    poleLatLongTable.Cells["A1:I1"].Style.Font.Bold = true;

                    poleLatLongTable.Cells["A1"].Value = "PTT";
                    poleLatLongTable.Cells["B1"].Value = "EXC_ABB";
                    poleLatLongTable.Cells["C1"].Value = "STATE";
                    poleLatLongTable.Cells["D1"].Value = "REGION";
                    poleLatLongTable.Cells["E1"].Value = "ZONE";
                    poleLatLongTable.Cells["F1"].Value = "POLE_FID";
                    poleLatLongTable.Cells["G1"].Value = "MANHOLE_TYPE";
                    poleLatLongTable.Cells["H1"].Value = "LATITUDE";
                    poleLatLongTable.Cells["I1"].Value = "LONGITUDE";

                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        poleLatLongTable.Cells["A" + num + ":I" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        poleLatLongTable.Cells["A" + num + ":I" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        poleLatLongTable.Cells["A" + num].Value = a.PTT;
                        poleLatLongTable.Cells["B" + num].Value = a.EXC_ABB;
                        poleLatLongTable.Cells["C" + num].Value = a.STATE;
                        poleLatLongTable.Cells["D" + num].Value = a.REGION;
                        poleLatLongTable.Cells["E" + num].Value = a.ZONE;
                        poleLatLongTable.Cells["F" + num].Value = a.POLE_FID;
                        poleLatLongTable.Cells["G" + num].Value = a.MANHOLE_TYPE;
                        poleLatLongTable.Cells["H" + num].Value = a.LATITUDE;
                        poleLatLongTable.Cells["I" + num++].Value = a.LONGITUDE;
                    }
                    poleLatLongTable.Cells.AutoFitColumns();
                }
                #endregion

                #region FEATURE STATE
                else if (reportType == "FEATURE STATE")
                {
                    var featStateTable = xlsFile.Workbook.Worksheets.Add("FEATURE STATE");
                    featStateTable.Cells.Style.Font.Size = 8;
                    featStateTable.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    featStateTable.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    featStateTable.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    featStateTable.Cells["A1:H1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    featStateTable.Cells["A1:H1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    featStateTable.Cells["A1:H1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    featStateTable.Cells["A1:H1"].Style.Font.Bold = true;

                    featStateTable.Cells["A1"].Value = "PTT";
                    featStateTable.Cells["B1"].Value = "EXC_ABB";
                    featStateTable.Cells["C1"].Value = "STATE";
                    featStateTable.Cells["D1"].Value = "REGION";
                    featStateTable.Cells["E1"].Value = "FEATURE_TYPE";
                    featStateTable.Cells["F1"].Value = "FEATURE_CODE";
                    featStateTable.Cells["G1"].Value = "FEATURE_FID";
                    featStateTable.Cells["H1"].Value = "FEATURE_STATE";

                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        featStateTable.Cells["A" + num + ":H" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        featStateTable.Cells["A" + num + ":H" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        featStateTable.Cells["A" + num].Value = a.PTT;
                        featStateTable.Cells["B" + num].Value = a.EXC_ABB;
                        featStateTable.Cells["C" + num].Value = a.STATE;
                        featStateTable.Cells["D" + num].Value = a.REGION;
                        featStateTable.Cells["E" + num].Value = a.FEATURE_TYPE;
                        featStateTable.Cells["F" + num].Value = a.FEATURE_CODE;
                        featStateTable.Cells["G" + num].Value = a.FEATURE_FID;
                        featStateTable.Cells["H" + num++].Value = a.FEATURE_STATE;
                    }
                    featStateTable.Cells.AutoFitColumns();
                }
                #endregion

                #region PROJECT NUMBER REPORT
                else if (reportType == "PROJECT NUMBER REPORT")
                {
                    var projNoTable = xlsFile.Workbook.Worksheets.Add("PROJECT NUMBER REPORT");
                    projNoTable.Cells.Style.Font.Size = 8;
                    projNoTable.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    projNoTable.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    projNoTable.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    projNoTable.Cells["A1:L1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    projNoTable.Cells["A1:L1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    projNoTable.Cells["A1:L1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projNoTable.Cells["A1:L1"].Style.Font.Bold = true;

                    projNoTable.Cells["A1"].Value = "PTT";
                    projNoTable.Cells["B1"].Value = "EXC_ABB";
                    projNoTable.Cells["C1"].Value = "PROJECT_NO";
                    projNoTable.Cells["D1"].Value = "WBS_NUM";
                    projNoTable.Cells["E1"].Value = "PROJ_DESC";
                    projNoTable.Cells["F1"].Value = "WBS_DESC";
                    projNoTable.Cells["G1"].Value = "SCHEME_NAME";
                    projNoTable.Cells["H1"].Value = "JOB_CREATED";
                    projNoTable.Cells["I1"].Value = "START_DATE";
                    projNoTable.Cells["J1"].Value = "END_DATE";
                    projNoTable.Cells["K1"].Value = "CREATED_DATE";
                    projNoTable.Cells["L1"].Value = "JOB_STATE";

                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        projNoTable.Cells["A" + num + ":L" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        projNoTable.Cells["A" + num + ":L" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        projNoTable.Cells["A" + num].Value = a.PTT;
                        projNoTable.Cells["B" + num].Value = a.EXC_ABB;
                        projNoTable.Cells["C" + num].Value = a.PROJECT_NO;
                        projNoTable.Cells["D" + num].Value = a.WBS_NUM;
                        projNoTable.Cells["E" + num].Value = a.PROJ_DESC;
                        projNoTable.Cells["F" + num].Value = a.WBS_DESC;
                        projNoTable.Cells["G" + num].Value = a.SCHEME_NAME;
                        projNoTable.Cells["H" + num].Value = a.JOB_CREATED;
                        projNoTable.Cells["I" + num].Value = a.START_DATE;
                        projNoTable.Cells["J" + num].Value = a.END_DATE;
                        projNoTable.Cells["K" + num].Value = a.CREATED_DATE;
                        projNoTable.Cells["L" + num++].Value = a.JOB_STATE;
                    }
                    projNoTable.Cells.AutoFitColumns();
                }
                #endregion

                #region PROJECT NUMBER MATCHING WITH JOB NUMBER
                else if (reportType == "PROJECT NUMBER MATCHING WITH JOB NUMBER")
                {
                    var projMatchJobTable = xlsFile.Workbook.Worksheets.Add("PROJECT NUMBER MATCHING WITH JOB NUMBER");
                    projMatchJobTable.Cells.Style.Font.Size = 8;
                    projMatchJobTable.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    projMatchJobTable.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    projMatchJobTable.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    projMatchJobTable.Cells["A1:H1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    projMatchJobTable.Cells["A1:H1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    projMatchJobTable.Cells["A1:H1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projMatchJobTable.Cells["A1:H1"].Style.Font.Bold = true;

                    projMatchJobTable.Cells["A1"].Value = "PTT";
                    projMatchJobTable.Cells["B1"].Value = "EXC_ABB";
                    projMatchJobTable.Cells["C1"].Value = "PROJECT_NO";
                    projMatchJobTable.Cells["D1"].Value = "WBS_NUM";
                    projMatchJobTable.Cells["E1"].Value = "PROJ_DESC";
                    projMatchJobTable.Cells["F1"].Value = "WBS_DESC";
                    projMatchJobTable.Cells["G1"].Value = "SCHEME_NAME";
                    projMatchJobTable.Cells["H1"].Value = "CREATED_DATE";

                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        projMatchJobTable.Cells["A" + num + ":H" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        projMatchJobTable.Cells["A" + num + ":H" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        projMatchJobTable.Cells["A" + num].Value = a.PTT;
                        projMatchJobTable.Cells["B" + num].Value = a.EXC_ABB;
                        projMatchJobTable.Cells["C" + num].Value = a.PROJECT_NO;
                        projMatchJobTable.Cells["D" + num].Value = a.WBS_NUM;
                        projMatchJobTable.Cells["E" + num].Value = a.PROJ_DESC;
                        projMatchJobTable.Cells["F" + num].Value = a.WBS_DESC;
                        projMatchJobTable.Cells["G" + num].Value = a.SCHEME_NAME;
                        projMatchJobTable.Cells["H" + num++].Value = a.CREATED_DATE;
                    }
                    projMatchJobTable.Cells.AutoFitColumns();
                }
                #endregion

                #region OLT MAPPING
                else if (reportType == "OLT MAPPING")
                {
                    var projOltMapping = xlsFile.Workbook.Worksheets.Add("OLT MAPPING");
                    projOltMapping.Cells.Style.Font.Size = 8;
                    projOltMapping.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    projOltMapping.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    projOltMapping.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    projOltMapping.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    projOltMapping.Row(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    projOltMapping.Cells["A1:A3"].Merge = true;
                    projOltMapping.Cells["A1:A3"].Value = "SEGMENT";
                    projOltMapping.Cells["B1:B3"].Merge = true;
                    projOltMapping.Cells["B1:B3"].Value = "EXC_ABB";
                    projOltMapping.Cells["C1:F1"].Merge = true;
                    projOltMapping.Cells["C1"].Value = "OLT INFORMATION ";
                    projOltMapping.Cells["G1"].Value = "ODF INFORMATION";
                    projOltMapping.Cells["S1"].Value = "FDC INFORMATION";
                    projOltMapping.Cells["V1"].Value = "FDP INFORMATION	";	
                    projOltMapping.Cells["G1:R1"].Merge = true;
                    projOltMapping.Cells["S1:U1"].Merge = true;
                    projOltMapping.Cells["V1:X1"].Merge = true;
                    projOltMapping.Cells["C2:C3"].Merge = true;
                    projOltMapping.Cells["C2"].Value = "OLT ID";
                    projOltMapping.Cells["D2:D3"].Merge = true;
                    projOltMapping.Cells["D2"].Value = "SHELF NO";
                    projOltMapping.Cells["E2:E3"].Merge = true;
                    projOltMapping.Cells["E2"].Value = "SLOT NO";
                    projOltMapping.Cells["F2:F3"].Merge = true;
                    projOltMapping.Cells["F2"].Value = "PLOT NO";
                    projOltMapping.Cells["G2:L2"].Merge = true;
                    projOltMapping.Cells["G2"].Value = "MAIN";
                    projOltMapping.Cells["M2:R2"].Merge = true;
                    projOltMapping.Cells["M2"].Value = "PROTECTION";
                    projOltMapping.Cells["G3"].Value = "ODF ID";
                    projOltMapping.Cells["H3"].Value = "CABLE CODE";
                    projOltMapping.Cells["I3"].Value = "NO OF CORE";
                    projOltMapping.Cells["J3"].Value = "CABLE CORE NO";
                    projOltMapping.Cells["K3"].Value = "ODF SHELF NO";
                    projOltMapping.Cells["L3"].Value = "ODF PORT NO";
                    projOltMapping.Cells["M3"].Value = "ODF ID";
                    projOltMapping.Cells["N3"].Value = "CABLE CODE";
                    projOltMapping.Cells["O3"].Value = "NO OF CORE";
                    projOltMapping.Cells["P3"].Value = "CABLE CORE NO";
                    projOltMapping.Cells["Q3"].Value = "ODF SHELF NO";
                    projOltMapping.Cells["R3"].Value = "ODF PORT NO";
                    projOltMapping.Cells["S2:S3"].Merge = true;
                    projOltMapping.Cells["S2"].Value = "FDC CODE";
                    projOltMapping.Cells["T2:T3"].Merge = true;
                    projOltMapping.Cells["T2"].Value = "SPLITTER CODE";
                    projOltMapping.Cells["U2:U3"].Merge = true;
                    projOltMapping.Cells["U2"].Value = "SPLITTER TYPE";
                    projOltMapping.Cells["V2:V3"].Merge = true;
                    projOltMapping.Cells["V2"].Value = "DSIDE CODE";
                    projOltMapping.Cells["W2:W3"].Merge = true;
                    projOltMapping.Cells["W2"].Value = "DSIDE CORE";
                    projOltMapping.Cells["X2:X3"].Merge = true;
                    projOltMapping.Cells["X2"].Value = "FDP CODE";
                    projOltMapping.Cells["A1:A3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["B1:B3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["C2:C3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["D2:D3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["E2:E3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["F2:F3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["G3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["H3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["I3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["J3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["K3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["L2:L3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["M3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["N3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["O3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["P3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["Q3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["R1:R3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["S2:S3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["T2:T3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["U1:U3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["V2:V3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["W2:W3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["X1:X3"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["C1:X1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["G2:R2"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["A3:X3"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    projOltMapping.Cells["A1"].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                    projOltMapping.Cells["B1"].Style.Fill.BackgroundColor.SetColor(Color.Aqua);
                    projOltMapping.Cells["C1"].Style.Fill.BackgroundColor.SetColor(Color.Orange);
                    projOltMapping.Cells["G1"].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    projOltMapping.Cells["S1"].Style.Fill.BackgroundColor.SetColor(Color.Cornsilk);
                    projOltMapping.Cells["S2:U3"].Style.Fill.BackgroundColor.SetColor(Color.Lime);
                    projOltMapping.Cells["V1"].Style.Fill.BackgroundColor.SetColor(Color.DarkOrange);
                    projOltMapping.Cells["V1"].Style.Fill.BackgroundColor.SetColor(Color.PaleTurquoise);
                    projOltMapping.Cells["C2:F3"].Style.Fill.BackgroundColor.SetColor(Color.Salmon);
                    projOltMapping.Cells["G2:L2"].Style.Fill.BackgroundColor.SetColor(Color.BlueViolet);
                    projOltMapping.Cells["G3:L3"].Style.Fill.BackgroundColor.SetColor(Color.Gray);
                    projOltMapping.Cells["M2:R2"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    projOltMapping.Cells["M3:R3"].Style.Fill.BackgroundColor.SetColor(Color.LemonChiffon);
                    projOltMapping.Cells["A1:X3"].Style.Font.Bold = true;
                   
                    int num = 4;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        projOltMapping.Cells["A" + num + ":X" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        projOltMapping.Cells["A" + num + ":X" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        projOltMapping.Cells["A" + num].Value = a.PTT;
                        projOltMapping.Cells["B" + num].Value = a.EXC_ABB;
                        projOltMapping.Cells["C" + num].Value = a.OLT_ID;
                        projOltMapping.Cells["D" + num].Value = a.OLT_SHELF;
                        projOltMapping.Cells["E" + num].Value = a.OLT_SLOT;
                        projOltMapping.Cells["F" + num].Value = a.OLT_PORT;

                        projOltMapping.Cells["G" + num].Value = a.ODF_ID;
                        projOltMapping.Cells["H" + num].Value = a.ESIDE_CBL_CODE;
                        projOltMapping.Cells["I" + num].Value = a.ESIDE_CBL_SIZE;
                        projOltMapping.Cells["J" + num].Value = a.ESIDE_CBL_CORE;
                        projOltMapping.Cells["K" + num].Value = a.ODF_SHELF;
                        projOltMapping.Cells["L" + num].Value = a.ODF_PORT;

                        projOltMapping.Cells["M" + num].Value = a.ODF_ID_P;
                        projOltMapping.Cells["N" + num].Value = a.ESIDE_CBL_CODE_P;
                        projOltMapping.Cells["O" + num].Value = a.ESIDE_CBL_SIZE;
                        projOltMapping.Cells["P" + num].Value = a.ESIDE_CBL_CORE_P;
                        projOltMapping.Cells["Q" + num].Value = a.ODF_SHELF_P;
                        projOltMapping.Cells["R" + num].Value = a.ODF_PORT_P;

                        projOltMapping.Cells["S" + num].Value = a.FDC_CODE;
                        projOltMapping.Cells["T" + num].Value = a.FDC_SP_CODE;
                        projOltMapping.Cells["U" + num].Value = a.FDC_SP_TYPE;
                        projOltMapping.Cells["V" + num].Value = a.DSIDE_CBL_CODE;
                        projOltMapping.Cells["W" + num].Value = a.DSIDE_CORE;
                        projOltMapping.Cells["X" + num++].Value = a.FDP_CODE;

                    }
                    projOltMapping.Cells.AutoFitColumns();
                }
                #endregion

                #region CORECOUNTESIDE
                
                else if (reportType == "CORE COUNT REPORT (ESIDE)")
                {
                    var projCoreCountEside = xlsFile.Workbook.Worksheets.Add("CORE COUNT ESIDE");
                    projCoreCountEside.Cells.Style.Font.Size = 8;
                    projCoreCountEside.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    projCoreCountEside.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    projCoreCountEside.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    projCoreCountEside.Cells["A1:H1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    projCoreCountEside.Cells["A1:H1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    projCoreCountEside.Cells["A1:H1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projCoreCountEside.Cells["A1:H1"].Style.Font.Bold = true;

                    projCoreCountEside.Cells["A1"].Value = "PTT";
                    projCoreCountEside.Cells["B1"].Value = "EXC_ABB";
                    projCoreCountEside.Cells["C1"].Value = "CABLE_CODE";
                    projCoreCountEside.Cells["D1"].Value = "CABLE_SIZE";
                    projCoreCountEside.Cells["E1"].Value = "TERM CODE";
                    projCoreCountEside.Cells["F1"].Value = "CORE STATUS";
                    projCoreCountEside.Cells["G1"].Value = "LOW CORE";
                    projCoreCountEside.Cells["H1"].Value = "HIGH CORE";

                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        projCoreCountEside.Cells["A" + num + ":H" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        projCoreCountEside.Cells["A" + num + ":H" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        projCoreCountEside.Cells["A" + num].Value = a.PTT;
                        projCoreCountEside.Cells["B" + num].Value = a.EXC_ABB;
                        projCoreCountEside.Cells["C" + num].Value = a.CABLE_CODE;
                        projCoreCountEside.Cells["D" + num].Value = a.CABLE_SIZE;
                        projCoreCountEside.Cells["E" + num].Value = a.TERM_CODE;
                        projCoreCountEside.Cells["F" + num].Value = a.CORE_STATUS;
                        projCoreCountEside.Cells["G" + num].Value = a.LOW_CORE;
                        projCoreCountEside.Cells["H" + num++].Value = a.HIGH_CORE;
                    }
                    projCoreCountEside.Cells.AutoFitColumns();
                }
                
                #endregion

                #region CORECOUNTDSIDE
                else if (reportType == "CORE COUNT REPORT (DSIDE)")
                {
                    var projCoreCountDside = xlsFile.Workbook.Worksheets.Add("CORE COUNT DSIDE");
                    projCoreCountDside.Cells.Style.Font.Size = 8;
                    projCoreCountDside.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    projCoreCountDside.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    projCoreCountDside.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    projCoreCountDside.Cells["A1:I1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    projCoreCountDside.Cells["A1:I1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    projCoreCountDside.Cells["A1:I1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projCoreCountDside.Cells["A1:I1"].Style.Font.Bold = true;

                    projCoreCountDside.Cells["A1"].Value = "PTT";
                    projCoreCountDside.Cells["B1"].Value = "EXC_ABB";
                    projCoreCountDside.Cells["C1"].Value = "FDC_CODE";
                    projCoreCountDside.Cells["D1"].Value = "CABLE_CODE";
                    projCoreCountDside.Cells["E1"].Value = "CABLE_SIZE";
                    projCoreCountDside.Cells["F1"].Value = "FDP_CODE";
                    projCoreCountDside.Cells["G1"].Value = "CORE_STATUS";
                    projCoreCountDside.Cells["H1"].Value = "LOW CORE";
                    projCoreCountDside.Cells["I1"].Value = "HIGH CORE";

   
                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        projCoreCountDside.Cells["A" + num + ":I" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        projCoreCountDside.Cells["A" + num + ":I" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        projCoreCountDside.Cells["A" + num].Value = a.PTT;
                        projCoreCountDside.Cells["B" + num].Value = a.EXC_ABB;
                        projCoreCountDside.Cells["C" + num].Value = a.FDC_CODE;
                        projCoreCountDside.Cells["D" + num].Value = a.CABLE_CODE;
                        projCoreCountDside.Cells["E" + num].Value = a.CABLE_SIZE;
                        projCoreCountDside.Cells["F" + num].Value = a.FDP_CODE;
                        projCoreCountDside.Cells["G" + num].Value = a.CORE_STATUS;
                        projCoreCountDside.Cells["H" + num].Value = a.LOW_CORE;
                        projCoreCountDside.Cells["I" + num++].Value = a.HIGH_CORE;
                    }
                    projCoreCountDside.Cells.AutoFitColumns();
                }
                #endregion

                #region FMU-REDMARK
                else if (reportType == "FMU - COUNT OF REDMARK")
                {
                    var projFMURedmark = xlsFile.Workbook.Worksheets.Add("FMU - COUNT OF REDMARK");
                    projFMURedmark.Cells.Style.Font.Size = 8;
                    projFMURedmark.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    projFMURedmark.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    projFMURedmark.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    projFMURedmark.Cells["A1:J1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    projFMURedmark.Cells["A1:J1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    projFMURedmark.Cells["A1:J1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projFMURedmark.Cells["A1:J1"].Style.Font.Bold = true;

                    projFMURedmark.Cells["A1"].Value = "USERNAME";
                    projFMURedmark.Cells["B1"].Value = "FULL_NAME";
                    projFMURedmark.Cells["C1"].Value = "GRP_ID";
                    projFMURedmark.Cells["D1"].Value = "GRPNAME";
                    projFMURedmark.Cells["E1"].Value = "PTT_STATE";
                    projFMURedmark.Cells["F1"].Value = "EXC_ABB";
                    projFMURedmark.Cells["G1"].Value = "G3E_DESCRIPTION";
                    projFMURedmark.Cells["H1"].Value = "G3E_CREATION";
                    projFMURedmark.Cells["I1"].Value = "JOB_STATE";
                    projFMURedmark.Cells["J1"].Value = "SCHEME_NAME";
        


                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        projFMURedmark.Cells["A" + num + ":J" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        projFMURedmark.Cells["A" + num + ":J" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        projFMURedmark.Cells["A" + num].Value = a.USERNAME;
                        projFMURedmark.Cells["B" + num].Value = a.FULL_NAME;
                        projFMURedmark.Cells["C" + num].Value = a.GRP_ID;
                        projFMURedmark.Cells["D" + num].Value = a.GRPNAME;
                        projFMURedmark.Cells["E" + num].Value = a.PTT_STATE;
                        projFMURedmark.Cells["F" + num].Value = a.EXC_ABB;
                        projFMURedmark.Cells["G" + num].Value = a.G3E_DESCRIPTION;
                        projFMURedmark.Cells["H" + num].Value = a.G3E_CREATION;
                        projFMURedmark.Cells["I" + num].Value = a.JOB_STATE;
                        projFMURedmark.Cells["J" + num++].Value = a.SCHEME_NAME;
                    }
                    projFMURedmark.Cells.AutoFitColumns();
                }
#endregion

                #region USERUSAGE
                else if (reportType == "USER USAGE TRACKING LOGIN")
                {
                    var projUserUsage = xlsFile.Workbook.Worksheets.Add("USER USAGE TRACKING LOGIN");
                    projUserUsage.Cells.Style.Font.Size = 8;
                    projUserUsage.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    projUserUsage.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    projUserUsage.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    projUserUsage.Cells["A1:F1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    projUserUsage.Cells["A1:F1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    projUserUsage.Cells["A1:F1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projUserUsage.Cells["A1:F1"].Style.Font.Bold = true;

                    projUserUsage.Cells["A1"].Value = "USERNAME";
                    projUserUsage.Cells["B1"].Value = "FULL_NAME";
                    projUserUsage.Cells["C1"].Value = "PTT_STATE";
                    projUserUsage.Cells["D1"].Value = "GRP_ID";
                    projUserUsage.Cells["E1"].Value = "GRPNAME";
                    projUserUsage.Cells["F1"].Value = "LOGOFF_TIME";
 

                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        projUserUsage.Cells["A" + num + ":F" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        projUserUsage.Cells["A" + num + ":F" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        projUserUsage.Cells["A" + num].Value = a.USERNAME;
                        projUserUsage.Cells["B" + num].Value = a.FULL_NAME;
                        projUserUsage.Cells["C" + num].Value = a.PTT_STATE;
                        projUserUsage.Cells["D" + num].Value = a.GRP_ID;
                        projUserUsage.Cells["E" + num].Value = a.GRPNAME;
                        projUserUsage.Cells["F" + num++].Value = a.LOGOFF_TIME;
                       
                    }
                    projUserUsage.Cells.AutoFitColumns();
                }
                #endregion
                #region PROPERTY COUNT CABINET
                else if (reportType == "PROPERTY COUNT (CABINET)")
                {
                    var projPropCountCab = xlsFile.Workbook.Worksheets.Add("PROPERTY COUNT CABINET");
                    projPropCountCab.Cells.Style.Font.Size = 8;
                    projPropCountCab.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    projPropCountCab.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    projPropCountCab.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    projPropCountCab.Cells["A1:F1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    projPropCountCab.Cells["A1:F1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    projPropCountCab.Cells["A1:F1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projPropCountCab.Cells["A1:F1"].Style.Font.Bold = true;

                    projPropCountCab.Cells["A1"].Value = "PTT";
                    projPropCountCab.Cells["B1"].Value = "EXC_ABB";
                    projPropCountCab.Cells["C1"].Value = "CABINET NUM";
                    projPropCountCab.Cells["D1"].Value = "PROPERTY_TYPE";
                    projPropCountCab.Cells["E1"].Value = "STATE";
                    projPropCountCab.Cells["F1"].Value = "TOTAL_COUNT";

                   
                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        projPropCountCab.Cells["A" + num + ":F" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        projPropCountCab.Cells["A" + num + ":F" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        projPropCountCab.Cells["A" + num].Value = a.PTT;
                        projPropCountCab.Cells["B" + num].Value = a.EXC_ABB;
                        projPropCountCab.Cells["C" + num].Value = a.CAB_NUM;
                        projPropCountCab.Cells["D" + num].Value = a.PROPERTY_TYPE;
                        projPropCountCab.Cells["E" + num].Value = a.STATE;
                        projPropCountCab.Cells["F" + num++].Value = a.TOTAL_COUNT;

                    }
                    projPropCountCab.Cells.AutoFitColumns();
                }
                #endregion
                #region PROPERTY COUNT FDC
                else if (reportType == "PROPERTY COUNT (FDC)")
                {
                    var projPropCountFDC = xlsFile.Workbook.Worksheets.Add("PROPERTY COUNT FDC");
                    projPropCountFDC.Cells.Style.Font.Size = 8;
                    projPropCountFDC.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    projPropCountFDC.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    projPropCountFDC.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    projPropCountFDC.Cells["A1:F1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    projPropCountFDC.Cells["A1:F1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    projPropCountFDC.Cells["A1:F1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projPropCountFDC.Cells["A1:F1"].Style.Font.Bold = true;

                    projPropCountFDC.Cells["A1"].Value = "PTT";
                    projPropCountFDC.Cells["B1"].Value = "EXC_ABB";
                    projPropCountFDC.Cells["C1"].Value = "FDC NUM";
                    projPropCountFDC.Cells["D1"].Value = "PROPERTY_TYPE";
                    projPropCountFDC.Cells["E1"].Value = "STATE";
                    projPropCountFDC.Cells["F1"].Value = "TOTAL_COUNT";


                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        projPropCountFDC.Cells["A" + num + ":F" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        projPropCountFDC.Cells["A" + num + ":F" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        projPropCountFDC.Cells["A" + num].Value = a.PTT;
                        projPropCountFDC.Cells["B" + num].Value = a.EXC_ABB;
                        projPropCountFDC.Cells["C" + num].Value = a.FDC_NUM;
                        projPropCountFDC.Cells["D" + num].Value = a.PROPERTY_TYPE;
                        projPropCountFDC.Cells["E" + num].Value = a.STATE;
                        projPropCountFDC.Cells["F" + num++].Value = a.TOTAL_COUNT;

                    }
                    projPropCountFDC.Cells.AutoFitColumns();
                }
                #endregion
                #region PROPERTY COUNT MSAN
                else if (reportType == "PROPERTY COUNT (MSAN)")
                {
                    var projPropCountMSAN = xlsFile.Workbook.Worksheets.Add("PROPERTY COUNT MSAN");
                    projPropCountMSAN.Cells.Style.Font.Size = 8;
                    projPropCountMSAN.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    projPropCountMSAN.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    projPropCountMSAN.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    projPropCountMSAN.Cells["A1:F1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    projPropCountMSAN.Cells["A1:F1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    projPropCountMSAN.Cells["A1:F1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projPropCountMSAN.Cells["A1:F1"].Style.Font.Bold = true;

                    projPropCountMSAN.Cells["A1"].Value = "PTT";
                    projPropCountMSAN.Cells["B1"].Value = "EXC_ABB";
                    projPropCountMSAN.Cells["C1"].Value = "MSAN NUM";
                    projPropCountMSAN.Cells["D1"].Value = "PROPERTY_TYPE";
                    projPropCountMSAN.Cells["E1"].Value = "STATE";
                    projPropCountMSAN.Cells["F1"].Value = "TOTAL_COUNT";


                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        projPropCountMSAN.Cells["A" + num + ":F" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        projPropCountMSAN.Cells["A" + num + ":F" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        projPropCountMSAN.Cells["A" + num].Value = a.PTT;
                        projPropCountMSAN.Cells["B" + num].Value = a.EXC_ABB;
                        projPropCountMSAN.Cells["C" + num].Value = a.MSAN_NUM;
                        projPropCountMSAN.Cells["D" + num].Value = a.PROPERTY_TYPE;
                        projPropCountMSAN.Cells["E" + num].Value = a.STATE;
                        projPropCountMSAN.Cells["F" + num++].Value = a.TOTAL_COUNT;

                    }
                    projPropCountMSAN.Cells.AutoFitColumns();
                }
                #endregion
                #region PROPERTY COUNT RT
                else if (reportType == "PROPERTY COUNT (RT)")
                {
                    var projPropCountRT = xlsFile.Workbook.Worksheets.Add("PROPERTY COUNT RT");
                    projPropCountRT.Cells.Style.Font.Size = 8;
                    projPropCountRT.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    projPropCountRT.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    projPropCountRT.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    projPropCountRT.Cells["A1:F1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    projPropCountRT.Cells["A1:F1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    projPropCountRT.Cells["A1:F1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projPropCountRT.Cells["A1:F1"].Style.Font.Bold = true;

                    projPropCountRT.Cells["A1"].Value = "PTT";
                    projPropCountRT.Cells["B1"].Value = "EXC_ABB";
                    projPropCountRT.Cells["C1"].Value = "RT NUM";
                    projPropCountRT.Cells["D1"].Value = "PROPERTY_TYPE";
                    projPropCountRT.Cells["E1"].Value = "STATE";
                    projPropCountRT.Cells["F1"].Value = "TOTAL_COUNT";


                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        projPropCountRT.Cells["A" + num + ":F" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        projPropCountRT.Cells["A" + num + ":F" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        projPropCountRT.Cells["A" + num].Value = a.PTT;
                        projPropCountRT.Cells["B" + num].Value = a.EXC_ABB;
                        projPropCountRT.Cells["C" + num].Value = a.RT_NUM;
                        projPropCountRT.Cells["D" + num].Value = a.PROPERTY_TYPE;
                        projPropCountRT.Cells["E" + num].Value = a.STATE;
                        projPropCountRT.Cells["F" + num++].Value = a.TOTAL_COUNT;

                    }
                    projPropCountRT.Cells.AutoFitColumns();
                }
                #endregion

                #region FEATURE COUNT
                else if (reportType == "FEATURE COUNT")
                {
                    var projFeatCount = xlsFile.Workbook.Worksheets.Add("FEATURE COUNT");
                    projFeatCount.Cells.Style.Font.Size = 8;
                    projFeatCount.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    projFeatCount.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    projFeatCount.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    projFeatCount.Cells["A1:D1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    projFeatCount.Cells["A1:D1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    projFeatCount.Cells["A1:D1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projFeatCount.Cells["A1:D1"].Style.Font.Bold = true;

                    projFeatCount.Cells["A1"].Value = "PTT";
                    projFeatCount.Cells["B1"].Value = "EXC_ABB";
                    projFeatCount.Cells["C1"].Value = "FEATURE NAME";
                    projFeatCount.Cells["D1"].Value = "TOTAL COUNT";
                   
                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        projFeatCount.Cells["A" + num + ":D" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        projFeatCount.Cells["A" + num + ":D" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        projFeatCount.Cells["A" + num].Value = a.PTT;
                        projFeatCount.Cells["B" + num].Value = a.EXC_ABB;
                        projFeatCount.Cells["C" + num].Value = a.FEATURE_NAME;
                        projFeatCount.Cells["D" + num++].Value = a.TOTAL_COUNT;
                        

                    }
                    projFeatCount.Cells.AutoFitColumns();
                }
#endregion
                #region CIVIL
                else if (reportType == "CIVIL REPORT")
                {
                    var projCivilRep = xlsFile.Workbook.Worksheets.Add("CIVIL REPORT");
                    projCivilRep.Cells.Style.Font.Size = 8;
                    projCivilRep.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    projCivilRep.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    projCivilRep.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    projCivilRep.Cells["A1:D1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    projCivilRep.Cells["A1:D1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    projCivilRep.Cells["A1:D1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projCivilRep.Cells["A1:D1"].Style.Font.Bold = true;

                    projCivilRep.Cells["A1"].Value = "PTT";
                    projCivilRep.Cells["B1"].Value = "EXC_ABB";
                    projCivilRep.Cells["C1"].Value = "FEATURE ";
                    projCivilRep.Cells["D1"].Value = "TOTAL COUNT";

                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        projCivilRep.Cells["A" + num + ":D" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        projCivilRep.Cells["A" + num + ":D" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        projCivilRep.Cells["A" + num].Value = a.PTT;
                        projCivilRep.Cells["B" + num].Value = a.EXC_ABB;
                        projCivilRep.Cells["C" + num].Value = a.FEATURE;
                        projCivilRep.Cells["D" + num++].Value = a.TOTAL_COUNT;


                    }
                    projCivilRep.Cells.AutoFitColumns();
                }
                #endregion
                #region FIBER ESIDE
                else if (reportType == "FIBER REPORT (ESIDE)")
                {

                    var projFiberES = xlsFile.Workbook.Worksheets.Add("FIBER ESIDE");
                    projFiberES.Cells.Style.Font.Size = 8;
                    projFiberES.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    projFiberES.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    projFiberES.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    projFiberES.Cells["A1:F1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    projFiberES.Cells["A1:F1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    projFiberES.Cells["A1:F1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projFiberES.Cells["A1:F1"].Style.Font.Bold = true;

                    projFiberES.Cells["A1"].Value = "PTT";
                    projFiberES.Cells["B1"].Value = "EXC_ABB";
                    projFiberES.Cells["C1"].Value = "CABLE_CODE ";
                    projFiberES.Cells["D1"].Value = "CABLE_TYPE";
                    projFiberES.Cells["E1"].Value = "CABLE_SIZE";
                    projFiberES.Cells["F1"].Value = "LENGTH_KM"; 
                   

                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        projFiberES.Cells["A" + num + ":F" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        projFiberES.Cells["A" + num + ":F" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        projFiberES.Cells["A" + num].Value = a.PTT;
                        projFiberES.Cells["B" + num].Value = a.EXC_ABB;
                        projFiberES.Cells["C" + num].Value = a.CABLE_CODE;
                        projFiberES.Cells["D" + num].Value = a.CABLE_TYPE;
                        projFiberES.Cells["E" + num].Value = a.CABLE_SIZE;
                        projFiberES.Cells["F" + num++].Value = a.LENGTH_KM;


                    }
                    projFiberES.Cells.AutoFitColumns();
                }
                #endregion
                #region FIBER DSIDE
                else if (reportType == "FIBER REPORT (DSIDE)")
                {

                    var projFiberDS = xlsFile.Workbook.Worksheets.Add("FIBER DSIDE");
                    projFiberDS.Cells.Style.Font.Size = 8;
                    projFiberDS.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    projFiberDS.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    projFiberDS.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    projFiberDS.Cells["A1:F1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    projFiberDS.Cells["A1:F1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    projFiberDS.Cells["A1:F1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projFiberDS.Cells["A1:F1"].Style.Font.Bold = true;

                    projFiberDS.Cells["A1"].Value = "PTT";
                    projFiberDS.Cells["B1"].Value = "EXC_ABB";
                    projFiberDS.Cells["C1"].Value = "CABLE_CODE ";
                    projFiberDS.Cells["D1"].Value = "CABLE_TYPE";
                    projFiberDS.Cells["E1"].Value = "CABLE_SIZE";
                    projFiberDS.Cells["F1"].Value = "LENGTH_KM";


                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        projFiberDS.Cells["A" + num + ":F" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        projFiberDS.Cells["A" + num + ":F" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        projFiberDS.Cells["A" + num].Value = a.PTT;
                        projFiberDS.Cells["B" + num].Value = a.EXC_ABB;
                        projFiberDS.Cells["C" + num].Value = a.CABLE_CODE;
                        projFiberDS.Cells["D" + num].Value = a.CABLE_TYPE;
                        projFiberDS.Cells["E" + num].Value = a.CABLE_SIZE;
                        projFiberDS.Cells["F" + num++].Value = a.LENGTH_KM;


                    }
                    projFiberDS.Cells.AutoFitColumns();
                }
                #endregion
                #region COPPER ESIDE
                else if (reportType == "COPPER REPORT (ESIDE)")
                {

                    var projCopES = xlsFile.Workbook.Worksheets.Add("COPPER ESIDE");
                    projCopES.Cells.Style.Font.Size = 8;
                    projCopES.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    projCopES.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    projCopES.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    projCopES.Cells["A1:F1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    projCopES.Cells["A1:F1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    projCopES.Cells["A1:F1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projCopES.Cells["A1:F1"].Style.Font.Bold = true;

                    projCopES.Cells["A1"].Value = "PTT";
                    projCopES.Cells["B1"].Value = "EXC_ABB";
                    projCopES.Cells["C1"].Value = "CABLE_CODE ";
                    projCopES.Cells["D1"].Value = "CABLE_TYPE";
                    projCopES.Cells["E1"].Value = "CABLE_SIZE";
                    projCopES.Cells["F1"].Value = "LENGTH_KM";


                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        projCopES.Cells["A" + num + ":F" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        projCopES.Cells["A" + num + ":F" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        projCopES.Cells["A" + num].Value = a.PTT;
                        projCopES.Cells["B" + num].Value = a.EXC_ABB;
                        projCopES.Cells["C" + num].Value = a.CABLE_CODE;
                        projCopES.Cells["D" + num].Value = a.CABLE_TYPE;
                        projCopES.Cells["E" + num].Value = a.CABLE_SIZE;
                        projCopES.Cells["F" + num++].Value = a.LENGTH_KM;


                    }
                    projCopES.Cells.AutoFitColumns();
                }
                #endregion
                #region COPPER DSIDE
                else if (reportType == "COPPER REPORT (DSIDE)")
                {

                    var projCopDS = xlsFile.Workbook.Worksheets.Add("COPPER DSIDE");
                    projCopDS.Cells.Style.Font.Size = 8;
                    projCopDS.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    projCopDS.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                    projCopDS.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    projCopDS.Cells["A1:F1"].Style.Fill.BackgroundColor.SetColor(Color.Cyan);
                    projCopDS.Cells["A1:F1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    projCopDS.Cells["A1:F1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    projCopDS.Cells["A1:F1"].Style.Font.Bold = true;

                    projCopDS.Cells["A1"].Value = "PTT";
                    projCopDS.Cells["B1"].Value = "EXC_ABB";
                    projCopDS.Cells["C1"].Value = "CABLE_CODE ";
                    projCopDS.Cells["D1"].Value = "CABLE_TYPE";
                    projCopDS.Cells["E1"].Value = "CABLE_SIZE";
                    projCopDS.Cells["F1"].Value = "LENGTH_KM";


                    int num = 2;
                    foreach (var a in NetInfraInfo1.DataNetInfraInfoList)
                    {
                        projCopDS.Cells["A" + num + ":F" + num].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        projCopDS.Cells["A" + num + ":F" + num].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        projCopDS.Cells["A" + num].Value = a.PTT;
                        projCopDS.Cells["B" + num].Value = a.EXC_ABB;
                        projCopDS.Cells["C" + num].Value = a.CABLE_CODE;
                        projCopDS.Cells["D" + num].Value = a.CABLE_TYPE;
                        projCopDS.Cells["E" + num].Value = a.CABLE_SIZE;
                        projCopDS.Cells["F" + num++].Value = a.LENGTH_KM;


                    }
                    projCopDS.Cells.AutoFitColumns();
                }
                #endregion
               

                var stream = new MemoryStream();
                xlsFile.SaveAs(stream);

                string fileName = reportType + ".xlsx";
                string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                stream.Position = 0;
                return File(stream, contentType, fileName);
            }
        }

        [HttpPost]
        public ActionResult GetEXCABB(string PTT)
        {
            string PuList = "";
            string Mast = "";
            //string PuList2 = "";
            using (Entities ctxData = new Entities())
            {
                List<SelectListItem> list = new List<SelectListItem>();
                var queryEXC = from p in ctxData.WV_EXC_MAST
                               where p.PTT_ID.Trim() == PTT.Trim()
                               orderby p.EXC_ABB
                               select new { p.EXC_ABB, p.EXC_NAME };
                
                foreach (var a in queryEXC.Distinct().OrderBy(it => it.EXC_NAME))
                {
                    PuList = PuList + a.EXC_NAME + " (" + a.EXC_ABB.Trim() + ") :  " + a.EXC_ABB + "|";
                }
                ViewBag.excabb = list;
            }

            return Json(new
            {
                Success = true,
                PuList = PuList
            }, JsonRequestBehavior.AllowGet); //
        }

      

    }
}
