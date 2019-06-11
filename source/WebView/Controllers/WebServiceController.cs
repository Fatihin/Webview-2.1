using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using System.Configuration;
using  System.Xml.Linq;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Serialization;
using Oracle.DataAccess.Types;
using Oracle.DataAccess.Client;
using WebView.Library;
using System.ServiceModel;
using System.Data;

namespace WebView.Controllers
{
/// <summary>
/// This class is an alternative when you can't use Service References. It allows you to invoke Web Methods on a given Web Service URL.
/// Based on the code from http://stackoverflow.com/questions/9482773/web-service-without-adding-a-reference
/// </summary>
public class WebService2
{

    private string connString = ConfigurationManager.AppSettings.Get("connString");
    public string Url { get; private set; }
    public string Method { get; private set; }
    public Dictionary<string, string> Params = new Dictionary<string, string>();
    public Dictionary<string, string> Node = new Dictionary<string, string>();
    public XDocument ResponseSOAP = XDocument.Parse("<root/>");
    public XDocument ResultXML = XDocument.Parse("<root/>");
    public string ResultString = String.Empty;

    public WebService2()
    {
        Url = String.Empty;
        Method = String.Empty;
    }
    public WebService2(string baseUrl)
    {
        Url = baseUrl;
        Method = String.Empty;
    }
    public WebService2(string baseUrl, string methodName)
    {
        Url = baseUrl;
        Method = methodName;
    }

    // Public API

    /// <summary>
    /// Adds a parameter to the WebMethod invocation.
    /// </summary>
    /// <param name="name">Name of the WebMethod parameter (case sensitive)</param>
    /// <param name="value">Value to pass to the paramenter</param>
    public void AddParameter(string name, string value)
    {
        Params.Add(name, value);
    }

    public void NodeParameter(string name, string value)
    {
        Node.Add(name, value);
    }

    public void Invoke()
    {
        //Invoke(Method, true);
    }

    /// <summary>
    /// Using the base url, invokes the WebMethod with the given name
    /// </summary>
    /// <param name="methodName">Web Method name</param>
    public void Invoke(string methodName)
    {

        Tools tool = new Tools();
        int idNEPSBI_PROC;
        using (EntitiesNetworkElement data_nepsbi = new EntitiesNetworkElement())
        {
            var id_process = (from p in data_nepsbi.BI_PROC_GRN_OSP
                              select p).OrderByDescending(p => p.PROC_ID).First();
            idNEPSBI_PROC = Convert.ToInt32(id_process.PROC_ID);
        }
        string jobid = "";
        string userId = "";
        foreach (var param in Params)
        {
            if (param.Key == "UserId")
            {
                userId = param.Value;
            }
            if (param.Key == "JobId")
            {
                jobid = param.Value;
            }
        }
      
        Invoke(methodName, true, idNEPSBI_PROC);
    }

    /// <summary>
    /// Cleans all internal data used in the last invocation, except the WebService's URL.
    /// This avoids creating a new WebService object when the URL you want to use is the same.
    /// </summary>
    public void CleanLastInvoke()
    {
        ResponseSOAP = ResultXML = null;
        ResultString = Method = String.Empty;
        Params = new Dictionary<string, string>();
    }

    #region Helper Methods

    /// <summary>
    /// Checks if the WebService's URL and the WebMethod's name are valid. If not, throws ArgumentNullException.
    /// </summary>
    /// <param name="methodName">Web Method name (optional)</param>
    private void AssertCanInvoke(string methodName = "")
    {
        if (Url == String.Empty)
            throw new ArgumentNullException("You tried to invoke a webservice without specifying the WebService's URL.");
        if ((methodName == "") && (Method == String.Empty))
            throw new ArgumentNullException("You tried to invoke a webservice without specifying the WebMethod.");
    }

    private void ExtractResult(string methodName, int procid, int req_id)
    {
        WebService2 WS2 = new WebService2();
        // Selects just the elements with namespace http://tempuri.org/ (i.e. ignores SOAP namespace)
        XmlNamespaceManager namespMan = new XmlNamespaceManager(new NameTable());
        namespMan.AddNamespace("nep", "http://www.tm.com.my/hsbb/eai/NEPSLoadPathConsumer/");

        XElement webMethodResult = ResponseSOAP.XPathSelectElement("//nep:NEPSLoadPathConsumerResponse", namespMan);
        System.Diagnostics.Debug.WriteLine(webMethodResult.Value.ToString());
        // If the result is an XML, return it and convert it to string
        string status = "";
        string code = "";
        string msg = "";
        if (webMethodResult.FirstNode.NodeType == XmlNodeType.Element)
        {
            foreach (XNode ex in ((XElement)webMethodResult).Nodes())
            {
                ResultXML = XDocument.Parse(ex.ToString());
                ResultXML = Utils.RemoveNamespaces(ResultXML);
                XElement a = ResultXML.Root;
                System.Diagnostics.Debug.WriteLine(a.FirstNode.ToString());
                System.Diagnostics.Debug.WriteLine(a.Elements().Count());
                foreach (XNode no in ((XElement)a).Nodes())
                {
                    XDocument ResultXMLN = XDocument.Parse(no.ToString());
                    XElement b = ResultXMLN.Root;
                    System.Diagnostics.Debug.WriteLine(b.Elements().Count());
                    System.Diagnostics.Debug.WriteLine(no.ToString());
                    if (b.Elements().Count() == 0)
                    {
                        string[] arr = no.ToString().Split('>');
                        string[] arr2 = arr[1].Split('<');
                        string[] arr3 = arr[0].Split('<');
                        System.Diagnostics.Debug.WriteLine(arr2[0]);
                        WS2.NodeParameter(arr3[1], arr2[0]);
                        if (arr3[1].Contains("Status"))
                        {
                            status = arr2[0];
                        }
                        else if (arr3[1].Contains("MessageCode"))
                        {
                            code = arr2[0];
                        }
                        else if (arr3[1].Contains("Message"))
                        {
                            msg = arr2[0];
                        }
                    }
                    else
                    {
                        foreach (XNode no2 in ((XElement)a).Nodes())
                        {
                            string[] arr = no2.ToString().Split('>');
                            string[] arr2 = arr[1].Split('<');
                            System.Diagnostics.Debug.WriteLine(arr2[0]);
                        }

                    }
                    ResultString = ResultXML.ToString();
                    System.Diagnostics.Debug.WriteLine(ResultString);
                    
                 
                    
                }
                
            }
            int idNEPSBI_RES;
            using (EntitiesNetworkElement data_nepsbi = new EntitiesNetworkElement())
            {
                var id_process = (from p in data_nepsbi.BI_GRNOSPDLDPATHCONS_RES
                                  select p).OrderByDescending(p => p.REQUEST_ID).First();
                idNEPSBI_RES = Convert.ToInt32(id_process.REQUEST_ID) + 1;
            }
            Tools tool = new Tools();
            string sqlStr = "insert into NEPSBI.BI_GRNOSPDLDPATHCONS_RES(REPLY_ID,PROC_ID,REQUEST_ID,TIME_RETURNED,CALLSTATUS_STATUS,CALLSTATUS_MSGCODE,CALLSTATUS_MSG) values";
            OracleParameter[] oraPrm = new OracleParameter[7];
            oraPrm[0] = new OracleParameter("v_REPLY_ID", OracleDbType.Int32);
            oraPrm[0].Value = idNEPSBI_RES;
            oraPrm[1] = new OracleParameter("v_PROC_ID", OracleDbType.Int32);
            oraPrm[1].Value = procid;
            oraPrm[2] = new OracleParameter("v_REQUEST_ID", OracleDbType.Int32);
            oraPrm[2].Value = req_id;
            oraPrm[3] = new OracleParameter("v_TIME_RETURNED", OracleDbType.TimeStamp);
            oraPrm[3].Value = DateTime.Now;
            oraPrm[4] = new OracleParameter("v_CALLSTATUS_STATUS", OracleDbType.Varchar2);
            oraPrm[4].Value = status;
            oraPrm[5] = new OracleParameter("v_CALLSTATUS_MSGCODE", OracleDbType.Varchar2);
            oraPrm[5].Value = code;
            oraPrm[6] = new OracleParameter("v_CALLSTATUS_MSG", OracleDbType.Varchar2);
            oraPrm[6].Value = msg;
            string result2 = tool.ExecuteStored(connString, sqlStr, CommandType.StoredProcedure, oraPrm, false);
            System.Diagnostics.Debug.WriteLine("result :" + result2);
        }
        // If the result is a string, return it and convert it to XML (creating a root node to wrap the result)
        else
        {
            ResultString = webMethodResult.FirstNode.ToString();
            ResultXML = XDocument.Parse("<root>" + ResultString + "</root>");
        }
    }

    /// <summary>
    /// Invokes a Web Method, with its parameters encoded or not.
    /// </summary>
    /// <param name="methodName">Name of the web method you want to call (case sensitive)</param>
    /// <param name="encode">Do you want to encode your parameters? (default: true)</param>
    private void Invoke(string methodName, bool encode, int proc_id)
    {
        Tools tool = new Tools();

        bool success = true;
     
        AssertCanInvoke(methodName);

        string jobid = "";
        string userId = "";
        foreach (var param in Params)
        {
            if (param.Key == "UserId")
            {
                userId = param.Value;
            }
            if (param.Key == "JobId")
            {
                jobid = param.Value;
            }
        }
        string soapStr ="";
        int idNEPSBI_REQ;
        using (Entities7 ctxData = new Entities7())
        {
            var pathData = (from p in ctxData.WV_LOAD_PATH_CONSUMER_ODP
                            where p.JOBID == jobid
                            select p);
            foreach (var data_val in pathData)
            {
                soapStr =
                @"<?xml version=""1.0"" encoding=""utf-8""?>
                <soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
                   xmlns:xsd=""http://www.w3.org/2001/XMLSchema""
                   xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
                   xmlns:nep=""http://www.tm.com.my/hsbb/eai/NEPSLoadPathConsumer/"">
                  <soap:Body>
                    <nep:NEPSLoadPathConsumerRequest>
                        <!--Optional:-->
                        <eventName>NEPSLoadPathConsumer</eventName>
                        <UserId>" + userId + @"</UserId>
                        <EquipmentASide>
                        <!--Optional:-->
                        <AName>" + data_val.ANAME + @"</AName>
                        <!--Optional:-->
                        <ASite>" + data_val.ASITE + @"</ASite>
                        <!--Optional:-->
                        <AType>" + data_val.ATYPE + @"</AType>
                        <ESide>
                            <!--Optional:-->
                            <ACard2>" + data_val.ACARD2 + @"</ACard2>
                            <!--Optional:-->
                            <APort2>" + data_val.APORT2 + @"</APort2>
                        </ESide>
                        <Splitter>
                            <!--Optional:-->
                            <ACard3>" + data_val.ACARD3 + @"</ACard3>
                            <!--Optional:-->
                            <APort3>" + data_val.APORT3 + @"</APort3>
                        </Splitter>
                        </EquipmentASide>
                        <EquipmentZSide>
                        <!--Optional:-->
            
                        <ZName>" + data_val.ZNAME + @"</ZName><ZSite>" + data_val.ZSITE + @"</ZSite>
                        <!--Optional:-->
            
                        <!--Optional:-->
                        <ZType>" + data_val.ZTYPE + @"</ZType>
                        <!--Optional:-->
                        <ZCard>" + data_val.ZCARD + @"</ZCard><ZPort>" + data_val.ZPORT + @"</ZPort>
                        </EquipmentZSide>
                        <EquipDP>
                        <DPName>" + data_val.DPNAME + @"</DPName>
                        <DPSite>" + data_val.DPSITE + @"</DPSite><DPPort>" + data_val.DPPORT + @"</DPPort>
                        <!--Optional:-->
            
                        </EquipDP>
                    </nep:NEPSLoadPathConsumerRequest>

                  </soap:Body>
                </soap:Envelope>";

                
               
                using (EntitiesNetworkElement data_nepsbi = new EntitiesNetworkElement())
                {
                    var id_process = (from p in data_nepsbi.BI_GRNOSPDLDPATHCONS_REQ
                                      select p).OrderByDescending(p => p.REQUEST_ID).First();
                    idNEPSBI_REQ = Convert.ToInt32(id_process.REQUEST_ID) + 1;
                }
                string sqlStr = "insert into NEPSBI.BI_GRNOSPDLDPATHCONS_REQ(REQUEST_ID,PROC_ID,TIME_SENT,EVENTNAME,USERID,ANAME ,ASITE ,ATYPE ,ACARD2,APORT2,ACARD3,APORT3,ZNAME ,ZSITE ,ZTYPE ,ZCARD ,ZPORT ,DPNAME,DPSITE,DPPORT) values";
                OracleParameter[] oraPrm = new OracleParameter[20];

                oraPrm[0] = new OracleParameter("v_REQUEST_ID", OracleDbType.Int32);
                oraPrm[0].Value = idNEPSBI_REQ;
                oraPrm[1] = new OracleParameter("v_PROC_ID", OracleDbType.Varchar2);
                oraPrm[1].Value = proc_id;
                oraPrm[2] = new OracleParameter("v_TIME_SENT", OracleDbType.TimeStamp);
                oraPrm[2].Value = DateTime.Now;
                oraPrm[3] = new OracleParameter("v_EVENTNAME", OracleDbType.Varchar2);
                oraPrm[3].Value = "NEPSLoadPathConsumer";
                oraPrm[4] = new OracleParameter("v_CLASS", OracleDbType.Varchar2);
                oraPrm[4].Value = "Webview.Granite.LoadPath";
                oraPrm[5] = new OracleParameter("v_ANAME", OracleDbType.Varchar2);
                oraPrm[5].Value = data_val.ANAME;
                oraPrm[6] = new OracleParameter("v_ASITE", OracleDbType.Varchar2);
                oraPrm[6].Value = data_val.ASITE;
                oraPrm[7] = new OracleParameter("v_ATYPE", OracleDbType.Varchar2);
                oraPrm[7].Value = data_val.ATYPE;
                oraPrm[8] = new OracleParameter("v_ACARD2", OracleDbType.Varchar2);
                oraPrm[8].Value = data_val.ACARD2;
                oraPrm[9] = new OracleParameter("v_APORT2", OracleDbType.Varchar2);
                oraPrm[9].Value = data_val.APORT2;
                oraPrm[10] = new OracleParameter("v_ACARD3", OracleDbType.Varchar2);
                oraPrm[10].Value = data_val.ACARD3;
                oraPrm[11] = new OracleParameter("v_APORT3", OracleDbType.Varchar2);
                oraPrm[11].Value = data_val.APORT3;
                oraPrm[12] = new OracleParameter("v_ZNAME", OracleDbType.Varchar2);
                oraPrm[12].Value = data_val.ZNAME;
                oraPrm[13] = new OracleParameter("v_ZSITE", OracleDbType.Varchar2);
                oraPrm[13].Value = data_val.ZSITE;
                oraPrm[14] = new OracleParameter("v_ZTYPE", OracleDbType.Varchar2);
                oraPrm[14].Value = data_val.ZTYPE;
                oraPrm[15] = new OracleParameter("v_ZCARD", OracleDbType.Varchar2);
                oraPrm[15].Value = data_val.ZCARD;
                oraPrm[16] = new OracleParameter("v_ZPORT", OracleDbType.Varchar2);
                oraPrm[16].Value = data_val.ZPORT;
                oraPrm[17] = new OracleParameter("v_DPNAME", OracleDbType.Varchar2);
                oraPrm[17].Value = data_val.DPNAME;
                oraPrm[18] = new OracleParameter("v_DPSITE", OracleDbType.Varchar2);
                oraPrm[18].Value = data_val.DPSITE;
                oraPrm[19] = new OracleParameter("v_DPPORT", OracleDbType.Varchar2);
                oraPrm[19].Value = data_val.DPPORT;
                string result2 = tool.ExecuteStored(connString, sqlStr, CommandType.StoredProcedure, oraPrm, false);
                System.Diagnostics.Debug.WriteLine("result :" + result2);

               
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(Url);
                req.Headers.Add("SOAPAction", "\"http://tempuri.org/" + methodName + "\"");
                req.ContentType = "text/xml;charset=\"utf-8\"";
                req.Accept = "text/xml";
                req.Method = "POST";
                //Stream stm = req.GetRequestStream();
                //StreamWriter stmw = new StreamWriter(stm);
                //stmw.Write(soapStr);
               
                ////StreamReader objReader = new StreamReader(req.GetResponse().GetResponseStream());
                //WebResponse wr = req.GetResponse();
                //Stream receiveStream = wr.GetResponseStream();
                //StreamReader reader = new StreamReader(receiveStream);
                //string result = reader.ReadToEnd();
                //ResponseSOAP = XDocument.Parse(Utils.UnescapeString(result));
                //ExtractResult(methodName, proc_id, idNEPSBI_REQ);

                //req.Timeout = 300000;
                //req.AllowWriteStreamBuffering = true;
                using (Stream stm = req.GetRequestStream())
                {
                    //string postValues = "";
                    //foreach (var param in Params)
                    //{
                    //    if (encode) postValues += string.Format("<{0}>{1}</{0}>", HttpUtility.HtmlEncode(param.Key), HttpUtility.HtmlEncode(param.Value));
                    //    else postValues += string.Format("<{0}>{1}</{0}>", param.Key, param.Value);
                    //}

                    //soapStr = string.Format(soapStr, methodName, postValues);

                    using (StreamWriter stmw = new StreamWriter(stm))
                    {
                        //Response.BufferOutput = true;
                        stmw.Write(soapStr);
                        stmw.Close();
                    }
                }
                using (StreamReader responseReader = new StreamReader(req.GetResponse().GetResponseStream()))
                {
                    string result = responseReader.ReadToEnd();
                    ResponseSOAP = XDocument.Parse(Utils.UnescapeString(result));
                    ExtractResult(methodName, proc_id, idNEPSBI_REQ);
                }
            }
        }
    }
    #endregion
}

public static class Utils
{
    /// <summary>
    /// Remove all xmlns:* instances from the passed XmlDocument to simplify our xpath expressions
    /// </summary>
    public static XDocument RemoveNamespaces(XDocument oldXml)
    {
        // FROM: http://social.msdn.microsoft.com/Forums/en-US/bed57335-827a-4731-b6da-a7636ac29f21/xdocument-remove-namespace?forum=linqprojectgeneral
        try
        {
            XDocument newXml = XDocument.Parse(Regex.Replace(
                oldXml.ToString(),
                @"(xmlns:?[^=]*=[""][^""]*[""])",
                "",
                RegexOptions.IgnoreCase | RegexOptions.Multiline)
            );
            return newXml;
        }
        catch (XmlException error)
        {
            throw new XmlException(error.Message + " at Utils.RemoveNamespaces");
        } 
    }

    /// <summary>
    /// Remove all xmlns:* instances from the passed XmlDocument to simplify our xpath expressions
    /// </summary>
    public static XDocument RemoveNamespaces(string oldXml)
    {
        XDocument newXml = XDocument.Parse(oldXml);
        return RemoveNamespaces(newXml);
    }

    /// <summary>
    /// Converts a string that has been HTML-enconded for HTTP transmission into a decoded string.
    /// </summary>
    /// <param name="escapedString">String to decode.</param>
    /// <returns>Decoded (unescaped) string.</returns>
    public static string UnescapeString(string escapedString)
    {
        return HttpUtility.HtmlDecode(escapedString);
    }
    }
}
