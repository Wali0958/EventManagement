using ClosedXML.Excel;
using EventManagement.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Web;
using System.Web.Mvc;
using static Microsoft.IO.RecyclableMemoryStreamManager;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.EMMA;
using QRCoder;
using System.Drawing.Imaging;
using System.Drawing;
using System.Text;
using DocumentFormat.OpenXml.ExtendedProperties;
using static iTextSharp.text.pdf.PdfDocument;

namespace EventManagement.Controllers
{
    public class HomeController : Controller
    {
        Master_Helper masterHelper = new Master_Helper();
        DataSet ds = new DataSet();
        bool Ispri = true;
        #region seach badge or print
        public ActionResult MasterSearchBadge()
        {
            TotalCount();
            return View();
        }
        public JsonResult SearchResult(string text)
        {
            int eventId = Convert.ToInt32(Session["EventId"]);
            int loginId = Convert.ToInt32(Session["LoginId"]);

            if (string.IsNullOrWhiteSpace(text))
                return Json(new { success = false, message = "Please enter a keyword to search." }, JsonRequestBehavior.AllowGet);

            var ds = masterHelper.SearchData("SearchData", text, eventId, loginId);
            var table = ds?.Tables[0];

            if (table == null || table.Rows.Count == 0)
                return Json(new { success = false, message = "No results found." }, JsonRequestBehavior.AllowGet);

            try
            {
                var sb = new StringBuilder();

                foreach (DataRow row in table.Rows)
                {
                    bool isPaid = row["PaymentStatus"].ToString().ToLower() == "true";
                    bool isPrinted = row["IsPrint"].ToString() == "No";

                    sb.Append($@"
                <tr>
                    <td>{row["BPathCode"]}</td>
                    <td>
                        <span class='val-name'>{row["Name"]}</span>
                        <div class='clearfix mt-2'></div>{row["Mobile"]}
                        <div class='clearfix mt-2'></div>{row["Email"]}
                        <div class='clearfix mt-2'></div>{row["Country"]}
                    </td>
                    <td><span class='val-sudo d-none'>{row["BadgeId"]}</span>{row["Category"]}</td>
                    <td>{row["Designation"]}</td>
                    <td>{row["Company"]}</td>
                    <td>
                        <img src='../Content/images/{(isPaid ? "ok" : "cross")}.png' height='70' width='70' />
                    </td>
                    <td>{row["IsPrint"]} / {row["PrintOn"]}</td>
                    <td>
                        <button type='button' onclick=""{(isPrinted ? $"Printstatus('{row["BadgeId"]}')" : $"Reprint('{row["BadgeId"]}')")}"" 
                            class='btn btn-{(isPrinted ? "success" : "warning")}'>{(isPrinted ? "Print" : "Re-Print")}</button>
                        &nbsp;<button type='button' class='btn btn-primary' onclick='Editcode(this)'>Edit</button>
                    </td>
                </tr>");
                }

                return Json(new { success = true, mData = sb.ToString() }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "Error: " + ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        public JsonResult PrintBadgeCodeSearch(string BadgeId)
        {
            int EventId = Convert.ToInt32(Session["EventId"].ToString());
            if (BadgeId.ToString() != "")
            {
                ds = masterHelper.SearchData("SearchDataNew", BadgeId.ToString(), EventId, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    try
                    {
                        //ds.Tables[0].Rows[0]["IsKitAllow"].ToString() == "Yes" || ds.Tables[0].Rows[0]["IsPaymentAllow"].ToString() == "Yes" &&
                        if (ds.Tables[0].Rows[0]["PaymentStatus"].ToString().ToLower() == "false")
                        {
                            StringBuilder sb = new StringBuilder();
                            sb.Append("<div class='row'>");
                            if (ds.Tables[0].Rows[0]["IsPaymentAllow"].ToString() == "Yes" && ds.Tables[0].Rows[0]["IsKitAllow"].ToString() == "Yes")
                            {
                                sb.Append("<input type='hidden' id='hfboth' value='1'></input>");
                            }
                            else
                            {
                                sb.Append("<input type='hidden' id='hfboth' value='0'></input>");
                            }
                            //if (ds.Tables[0].Rows[0]["IsKitAllow"].ToString() == "Yes")
                            //{
                            //    if (ds.Tables[0].Rows[0]["KitCollect"].ToString().ToLower() == "yes")
                            //    {
                            //        sb.Append("<div class='col-sm-12'><p>Please select kit status</p><select class='form-select' id='ddlkitstatus' name='KitCollect' required=''><option value='true' selected=''>Yes</option><option value='false'>No</option></select></div>");

                            //    }
                            //    else
                            //    {
                            //        sb.Append("<div class='col-sm-12'><p>Please select kit status</p><select class='form-select' id='ddlkitstatus' name='KitCollect' required=''><option value='true'>Yes</option><option value='false' selected=''>No</option></select></div>");
                            //    }
                            //    sb.Append("<div class='clesrfix mt-2'></div>");
                            //    if (ds.Tables[0].Rows[0]["CollectBy"].ToString().ToLower() != "")
                            //    {
                            //        sb.Append("<div class='col-sm-12'><p>Please enter collect by</p><input type='text' value=" + ds.Tables[0].Rows[0]["CollectBy"].ToString() + " class='form-control' id='CollectBy' name='CollectBy' maxlength='35' placeholder='Please enter collected by name' required=''/></div>");
                            //    }
                            //    else
                            //    {
                            //        sb.Append("<div class='col-sm-12'><p>Please enter collect by</p><input type='text' class='form-control' id='CollectBy' name='CollectBy' maxlength='35' placeholder='Please enter collected by name' required=''/></div>");
                            //    }
                            //}
                            //sb.Append("<div class='clesrfix mt-2'></div>");
                            //if (ds.Tables[0].Rows[0]["IsPaymentAllow"].ToString() == "Yes")
                            //{
                            //    if (ds.Tables[0].Rows[0]["PaymentStatus"].ToString() == "1")
                            //    {
                            //        sb.Append("<div class='col-sm-12'><p>Please select payment status</p><select class='form-select' id='ddlpayment' name='PaymentStatus' required=''><option value='true' selected=''>Yes</option><option value='false'>No</option></select></div>");
                            //    }
                            //    else
                            //    {
                            //        sb.Append("<div class='col-sm-12'><p>Please select payment status</p><select class='form-select' id='ddlpayment' name='PaymentStatus' required=''><option value='true'>Yes</option><option value='false' selected=''>No</option></select></div>");
                            //    }

                            //}
                            // sb.Append("<div class='clesrfix mt-2'></div>");
                            if (ds.Tables[0].Rows[0]["Remarks"].ToString() == "")
                            {
                                sb.Append("<div class='col-sm-12'><p>Remarks</p><textarea id='commentsname'Remarks' rows='4' cols='100' placeholder='Type your Remarks...' required=''></textarea></div>");
                            }
                            else
                            {
                                sb.Append("<div class='col-sm-12'><p>Remarks</p><textarea id='commentsname'Remarks' rows='4' cols='100' placeholder='Type your Remarks...' required=''>" + ds.Tables[0].Rows[0]["Remarks"].ToString() + "</textarea></div>");

                            }
                            sb.Append("<div class='clesrfix mt-2'></div>");
                            sb.Append("<div class='col-sm-12'><div class='text-center'><button type='submit' id='btnprintpop' onclick=Printthroughpopup('" + ds.Tables[0].Rows[0]["BadgeId"].ToString() + "')  class='btn btn-warning text-center'>Print</button>&nbsp;&nbsp;<button type='button' class='btn btn-secondary pull-right mr-2' data-bs-dismiss='modal'>Close</button></div></div>");
                            sb.Append("</div>");
                            sb.Append("</tr>");
                            string mPrintStatusCheck = sb.ToString();
                            TotalCount();
                            return Json(new { success = true, mData = mPrintStatusCheck, mData1 = "update status or then print" }, JsonRequestBehavior.AllowGet);
                        }
                        else
                        {
                            try
                            {
                                if (EventId != 0 && BadgeId.ToString() != "0")
                                {
                                    ds = masterHelper.BadgePrintDataBulk("DetailBadgePrint", EventId, Convert.ToInt32(BadgeId), Convert.ToInt16(Session["LoginId"]));
                                    if (ds != null && ds.Tables[0].Rows.Count > 0)
                                    {
                                        try
                                        {
                                            StringBuilder sb = new StringBuilder();
                                            sb.Append("<input type='hidden' id='badgeid' value='" + BadgeId + "'>");
                                            sb.Append("<div id=divprint style='display:none;'>");
                                            if (ds.Tables[1].Rows.Count > 0)
                                            {
                                                for (int i = 0; ds.Tables[1].Rows.Count > i; i++)
                                                {
                                                    for (int j = 0; ds.Tables[1].Rows.Count > j; j++)
                                                    {
                                                        string mstyleloopcode = $"margin-top: {ds.Tables[1].Rows[j]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[j]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[j]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[j]["MarginRight"].ToString().ToLower()}px;";

                                                        if (ds.Tables[1].Rows[j]["columName"].ToString().ToLower() == "photo" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "photo")
                                                        {
                                                            sb.Append("<div style=\"" + mstyleloopcode + "\"><img src='../Content/BadgePhoto/" + Session["EventName"].ToString() + "/" + ds.Tables[0].Rows[0]["Photo"].ToString() + "' height='75px' width='75px'></div>");
                                                        }
                                                    }
                                                    for (int k = 0; ds.Tables[1].Rows.Count > k; k++)
                                                    {
                                                        string mstyleloopcode = $"text-align: center; margin-top: {ds.Tables[1].Rows[k]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[k]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[k]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[k]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[1].Rows[k]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[1].Rows[k]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[1].Rows[k]["FontWeight"].ToString().ToLower()};";

                                                        if (ds.Tables[1].Rows[k]["columName"].ToString().ToLower() == "name" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "name")
                                                        {
                                                            sb.Append("<div style=\"" + mstyleloopcode + "\">" + ds.Tables[0].Rows[0]["Name"].ToString() + "</div>");
                                                        }
                                                    }
                                                    for (int l = 0; ds.Tables[1].Rows.Count > l; l++)
                                                    {
                                                        string mstyleloopcode = $"text-align: center; margin-top: {ds.Tables[1].Rows[l]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[l]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[l]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[l]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[1].Rows[l]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[1].Rows[l]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[1].Rows[l]["FontWeight"].ToString().ToLower()};";

                                                        if (ds.Tables[1].Rows[l]["columName"].ToString().ToLower() == "designation" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "designation")
                                                        {
                                                            sb.Append("<div style=\"" + mstyleloopcode + "\">" + ds.Tables[0].Rows[0]["Designation"].ToString() + "</div>");
                                                        }
                                                    }
                                                    for (int m = 0; ds.Tables[1].Rows.Count > m; m++)
                                                    {
                                                        string mstyleloopcode = $"text-align: center; margin-top: {ds.Tables[1].Rows[m]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[m]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[m]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[m]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[1].Rows[m]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[1].Rows[m]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[1].Rows[m]["FontWeight"].ToString().ToLower()};";

                                                        if (ds.Tables[1].Rows[m]["columName"].ToString().ToLower() == "company" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "company")
                                                        {
                                                            sb.Append("<div style=\"" + mstyleloopcode + "\">" + ds.Tables[0].Rows[0]["Company"].ToString() + "</div>");
                                                        }
                                                    }
                                                    for (int n = 0; ds.Tables[1].Rows.Count > n; n++)
                                                    {
                                                        string mstyleloopcode = $"text-align: center; margin-top: {ds.Tables[1].Rows[n]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[n]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[n]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[n]["MarginRight"].ToString().ToLower()}px;";

                                                        if (ds.Tables[1].Rows[n]["columName"].ToString().ToLower() == "barcode" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "barcode")
                                                        {
                                                            sb.Append("<div style=\"" + mstyleloopcode + "\"><img src='../Content/QRCodesScanner/" + Session["EventName"].ToString() + "/" +
                                                                ds.Tables[0].Rows[0]["BPath"].ToString() + "' height='75px' width='75px'>" +
                                                               "<div class=clearfix></div><label style = 'font-size:9px; padding-bottom:20px  !important;' > " + ds.Tables[0].Rows[0]["BPathCode"].ToString() + "</label>" +
                                                                "</div>");
                                                        }
                                                    }
                                                    for (int o = 0; ds.Tables[1].Rows.Count > o; o++)
                                                    {
                                                        string mstyleloopcode = $"margin-top: {ds.Tables[1].Rows[o]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[o]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[o]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[o]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[1].Rows[o]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[1].Rows[o]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[1].Rows[o]["FontWeight"].ToString().ToLower()};";

                                                        if (ds.Tables[1].Rows[o]["columName"].ToString().ToLower() == "category" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "category")
                                                        {
                                                            sb.Append("<div style=\"" + mstyleloopcode + "\" id='mcat'>" + ds.Tables[0].Rows[0]["Category"].ToString() + "</div>");
                                                        }
                                                    }
                                                }
                                            }
                                            sb.Append("</div>");
                                            sb.Append("</div>");
                                            string mBadge = sb.ToString();
                                            TotalCount();
                                            return Json(new { success = true, mData = mBadge, mData1 = "Only Print" }, JsonRequestBehavior.AllowGet);
                                        }
                                        catch (Exception ex)
                                        {
                                            TempData["Msg"] = ex.Message + " Oops some error occured";
                                            return Json(new { success = false, message = TempData["Msg"] });
                                        }
                                    }
                                    else
                                    {
                                        TempData["Msg"] = "Oops some error occured";
                                        return Json(new { success = false, message = TempData["Msg"] });
                                    }
                                }
                                else
                                {
                                    TempData["Msg"] = "Please fill all detsils.";
                                    return Json(new { success = false, message = TempData["Msg"] });
                                }

                            }
                            catch (Exception ex)
                            {
                                TempData["Msg"] = ex + "Oops some error occured.";
                                return Json(new { success = false, message = TempData["Msg"] });
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        TempData["Msg"] = ex.Message + " Oops some error occured";
                        return Json(new { success = false, message = TempData["Msg"] });
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                    return Json(new { success = false, message = TempData["Msg"] });
                }
            }
            else
            {
                TempData["Msg"] = "Please fill keyword to find data";
                return Json(new
                {
                    success = false,
                    message = TempData["Msg"]
                });
            }
        }
        string newbadgeid;
        [HttpGet]
        public JsonResult MultiplePrintBadgeCodeSearch(List<int> BadgeId)
        {

            if (BadgeId.Count > 0)
            {
                for (int g = 0; g < BadgeId.Count; g++)
                {
                    newbadgeid = newbadgeid + "," + BadgeId[g];
                }
            }
            int EventId = Convert.ToInt32(Session["EventId"].ToString());
            if (BadgeId.ToString() != "")
            {
                DataSet dss = masterHelper.SearchData("MultiSearchDataNew", newbadgeid.ToString().Substring(1), EventId, Convert.ToInt16(Session["LoginId"]));
                if (dss != null && dss.Tables[0].Rows.Count > 0)
                {
                    try
                    {
                        StringBuilder sb = new StringBuilder();
                        for (int z = 0; dss.Tables[0].Rows.Count > z; z++)
                        {
                            ds = masterHelper.BadgePrintDataBulk("DetailBadgePrint", EventId, Convert.ToInt32(dss.Tables[0].Rows[z]["BadgeId"]), Convert.ToInt16(Session["LoginId"]));
                            if (ds != null && ds.Tables[0].Rows.Count > 0)
                            {
                                try
                                {
                                    sb.Append("<input type='hidden' id='badgeid' value='" + dss.Tables[0].Rows[z]["BadgeId"] + "'>");
                                    sb.Append("<div id='divprint_" + dss.Tables[0].Rows[z]["BadgeId"] + "' class='print-page' style='display:none;'>");
                                    if (ds.Tables[1].Rows.Count > 0)
                                    {
                                        for (int i = 0; ds.Tables[1].Rows.Count > i; i++)
                                        {
                                            for (int j = 0; ds.Tables[1].Rows.Count > j; j++)
                                            {
                                                string mstyleloopcode = $"margin-top: {ds.Tables[1].Rows[j]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[j]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[j]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[j]["MarginRight"].ToString().ToLower()}px;";

                                                if (ds.Tables[1].Rows[j]["columName"].ToString().ToLower() == "photo" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "photo")
                                                {
                                                    sb.Append("<div style=\"" + mstyleloopcode + "\"><img src='../Content/BadgePhoto/" + Session["EventName"].ToString() + "/" + ds.Tables[0].Rows[0]["Photo"].ToString() + "' height='75px' width='75px'></div>");
                                                }
                                            }
                                            for (int k = 0; ds.Tables[1].Rows.Count > k; k++)
                                            {
                                                string mstyleloopcode = $"text-align: center; margin-top: {ds.Tables[1].Rows[k]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[k]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[k]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[k]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[1].Rows[k]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[1].Rows[k]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[1].Rows[k]["FontWeight"].ToString().ToLower()};";

                                                if (ds.Tables[1].Rows[k]["columName"].ToString().ToLower() == "name" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "name")
                                                {
                                                    sb.Append("<div style=\"" + mstyleloopcode + "\">" + ds.Tables[0].Rows[0]["Name"].ToString() + "</div>");
                                                }
                                            }
                                            for (int l = 0; ds.Tables[1].Rows.Count > l; l++)
                                            {
                                                string mstyleloopcode = $"text-align: center; margin-top: {ds.Tables[1].Rows[l]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[l]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[l]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[l]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[1].Rows[l]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[1].Rows[l]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[1].Rows[l]["FontWeight"].ToString().ToLower()};";

                                                if (ds.Tables[1].Rows[l]["columName"].ToString().ToLower() == "designation" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "designation")
                                                {
                                                    sb.Append("<div style=\"" + mstyleloopcode + "\">" + ds.Tables[0].Rows[0]["Designation"].ToString() + "</div>");
                                                }
                                            }
                                            for (int m = 0; ds.Tables[1].Rows.Count > m; m++)
                                            {
                                                string mstyleloopcode = $"text-align: center; margin-top: {ds.Tables[1].Rows[m]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[m]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[m]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[m]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[1].Rows[m]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[1].Rows[m]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[1].Rows[m]["FontWeight"].ToString().ToLower()};";

                                                if (ds.Tables[1].Rows[m]["columName"].ToString().ToLower() == "company" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "company")
                                                {
                                                    sb.Append("<div style=\"" + mstyleloopcode + "\">" + ds.Tables[0].Rows[0]["Company"].ToString() + "</div>");
                                                }
                                            }
                                            for (int n = 0; ds.Tables[1].Rows.Count > n; n++)
                                            {
                                                string mstyleloopcode = $"text-align: center; margin-top: {ds.Tables[1].Rows[n]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[n]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[n]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[n]["MarginRight"].ToString().ToLower()}px;";

                                                if (ds.Tables[1].Rows[n]["columName"].ToString().ToLower() == "barcode" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "barcode")
                                                {
                                                    sb.Append("<div style=\"" + mstyleloopcode + "\"><img src='../Content/QRCodesScanner/" + Session["EventName"].ToString() + "/" +
                                                        ds.Tables[0].Rows[0]["BPath"].ToString() + "' height='75px' width='75px'>" +
                                                          "<div class=clearfix></div><label style = 'font-size:9px; padding-bottom:20px  !important;' > " + ds.Tables[0].Rows[0]["BPathCode"].ToString() + "</label>" +
                                                            "</div>");
                                                }
                                            }
                                            for (int o = 0; ds.Tables[1].Rows.Count > o; o++)
                                            {
                                                string mstyleloopcode = $"margin-top: {ds.Tables[1].Rows[o]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[o]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[o]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[o]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[1].Rows[o]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[1].Rows[o]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[1].Rows[o]["FontWeight"].ToString().ToLower()};";

                                                if (ds.Tables[1].Rows[o]["columName"].ToString().ToLower() == "category" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "category")
                                                {
                                                    sb.Append("<div style=\"" + mstyleloopcode + "\" id='mcat'>" + ds.Tables[0].Rows[0]["Category"].ToString() + "</div>");
                                                }
                                            }
                                        }

                                    }
                                    sb.Append("</div>");
                                    sb.Append("</div>");
                                    sb.Append("<div style='page-break-after: always;'></div>");
                                }
                                catch (Exception ex)
                                {
                                    TempData["Msg"] = ex.Message + " Oops some error occured";
                                    return Json(new { success = false, message = TempData["Msg"] });
                                }
                            }
                            else
                            {
                                TempData["Msg"] = "Oops some error occured";
                                return Json(new { success = false, message = TempData["Msg"] });
                            }
                        }
                        string mBadge = sb.ToString();
                        return Json(new { success = true, mData = mBadge }, JsonRequestBehavior.AllowGet);
                    }
                    catch (Exception ex)
                    {
                        TempData["Msg"] = ex + "Oops some error occured.";
                        return Json(new { success = false, message = TempData["Msg"] });
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                    return Json(new { success = false, message = TempData["Msg"] });
                }
            }
            else
            {
                TempData["Msg"] = "Please fill keyword to find data";
                return Json(new
                {
                    success = false,
                    message = TempData["Msg"]
                });
            }
        }
        [HttpPost]
        public JsonResult PrintBadgeCodeAfterupdate(string BadgeId, string kitstatus, string CollectBy, string Payment, string commentsname)
        {
            try
            {
                int EventId = Convert.ToInt32(Session["EventId"].ToString());
                if (EventId != 0 && BadgeId.ToString() != "0")
                {
                    ds = masterHelper.BadgePrintDataBulk2("DetailBadgePrint2", EventId, Convert.ToInt32(BadgeId), Convert.ToInt16(Session["LoginId"]),
                                                                                                        kitstatus, CollectBy, Payment, commentsname);
                    if (ds != null && ds.Tables[0].Rows.Count > 0)
                    {
                        try
                        {
                            StringBuilder sb = new StringBuilder();
                            sb.Append("<input type='hidden' id='badgeid' value='" + BadgeId + "'>");
                            sb.Append("<div id=divprint style='display:none;'>");
                            if (ds.Tables[1].Rows.Count > 0)
                            {
                                for (int i = 0; ds.Tables[1].Rows.Count > i; i++)
                                {
                                    for (int j = 0; ds.Tables[1].Rows.Count > j; j++)
                                    {
                                        string mstyleloopcode = $"margin-top: {ds.Tables[1].Rows[j]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[j]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[j]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[j]["MarginRight"].ToString().ToLower()}px;";

                                        if (ds.Tables[1].Rows[j]["columName"].ToString().ToLower() == "photo" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "photo")
                                        {
                                            sb.Append("<div style=\"" + mstyleloopcode + "\"><img src='../Content/BadgePhoto/" + Session["EventName"].ToString() + "/" + ds.Tables[0].Rows[0]["Photo"].ToString() + "' height='75px' width='75px'></div>");
                                        }
                                    }
                                    for (int k = 0; ds.Tables[1].Rows.Count > k; k++)
                                    {
                                        string mstyleloopcode = $"text-align: center; margin-top: {ds.Tables[1].Rows[k]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[k]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[k]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[k]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[1].Rows[k]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[1].Rows[k]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[1].Rows[k]["FontWeight"].ToString().ToLower()};";

                                        if (ds.Tables[1].Rows[k]["columName"].ToString().ToLower() == "name" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "name")
                                        {
                                            sb.Append("<div style=\"" + mstyleloopcode + "\">" + ds.Tables[0].Rows[0]["Name"].ToString() + "</div>");
                                        }
                                    }
                                    for (int l = 0; ds.Tables[1].Rows.Count > l; l++)
                                    {
                                        string mstyleloopcode = $"text-align: center; margin-top: {ds.Tables[1].Rows[l]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[l]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[l]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[l]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[1].Rows[l]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[1].Rows[l]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[1].Rows[l]["FontWeight"].ToString().ToLower()};";

                                        if (ds.Tables[1].Rows[l]["columName"].ToString().ToLower() == "designation" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "designation")
                                        {
                                            sb.Append("<div style=\"" + mstyleloopcode + "\">" + ds.Tables[0].Rows[0]["Designation"].ToString() + "</div>");
                                        }
                                    }
                                    for (int m = 0; ds.Tables[1].Rows.Count > m; m++)
                                    {
                                        string mstyleloopcode = $"text-align: center; margin-top: {ds.Tables[1].Rows[m]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[m]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[m]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[m]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[1].Rows[m]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[1].Rows[m]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[1].Rows[m]["FontWeight"].ToString().ToLower()};";

                                        if (ds.Tables[1].Rows[m]["columName"].ToString().ToLower() == "company" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "company")
                                        {
                                            sb.Append("<div style=\"" + mstyleloopcode + "\">" + ds.Tables[0].Rows[0]["Company"].ToString() + "</div>");
                                        }
                                    }
                                    for (int n = 0; ds.Tables[1].Rows.Count > n; n++)
                                    {
                                        string mstyleloopcode = $"text-align: center; margin-top: {ds.Tables[1].Rows[n]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[n]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[n]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[n]["MarginRight"].ToString().ToLower()}px;";

                                        if (ds.Tables[1].Rows[n]["columName"].ToString().ToLower() == "barcode" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "barcode")
                                        {
                                            sb.Append("<div style=\"" + mstyleloopcode + "\"><img src='../Content/QRCodesScanner/" + Session["EventName"].ToString() + "/" +
                                                ds.Tables[0].Rows[0]["BPath"].ToString() + "' height='75px' width='75px'>" +
                                                  "<div class=clearfix></div><label style = 'font-size:9px; padding-bottom:20px  !important;' > " + ds.Tables[0].Rows[0]["BPathCode"].ToString() + "</label>" +
                                                  "</div>");

                                        }
                                    }
                                    for (int o = 0; ds.Tables[1].Rows.Count > o; o++)
                                    {
                                        string mstyleloopcode = $"margin-top: {ds.Tables[1].Rows[o]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[o]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[o]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[o]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[1].Rows[o]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[1].Rows[o]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[1].Rows[o]["FontWeight"].ToString().ToLower()};";

                                        if (ds.Tables[1].Rows[o]["columName"].ToString().ToLower() == "category" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "category")
                                        {
                                            sb.Append("<div style=\"" + mstyleloopcode + "\" id='mcat'>" + ds.Tables[0].Rows[0]["Category"].ToString() + "</div>");
                                        }
                                    }
                                }
                            }
                            sb.Append("</div>");
                            sb.Append("</div>");
                            string mBadge = sb.ToString();
                            return Json(new { success = true, mData = mBadge }, JsonRequestBehavior.AllowGet);
                        }
                        catch (Exception ex)
                        {
                            TempData["Msg"] = ex.Message + " Oops some error occured";
                            return Json(new { success = false, message = TempData["Msg"] });
                        }
                    }
                    else
                    {
                        TempData["Msg"] = "Oops some error occured";
                        return Json(new { success = false, message = TempData["Msg"] });
                    }
                }
                else
                {
                    TempData["Msg"] = "Please fill all detsils.";
                    return Json(new { success = false, message = TempData["Msg"] });
                }

            }
            catch (Exception ex)
            {
                TempData["Msg"] = ex + "Oops some error occured.";
                return Json(new { success = false, message = TempData["Msg"] });
            }
        }
        [HttpPost]
        public JsonResult UpdateBadge(int BadgeId, string Name, string Company, string Designation)
        {
            try
            {
                if (BadgeId.ToString() != "0")
                {
                    ds = masterHelper.BadgeUpdate("UpdateBadgePr", Convert.ToInt32(BadgeId), Name, Company, Designation);
                    if (ds != null && ds.Tables[0].Rows.Count > 0)
                    {
                        return Json(new { success = true });
                    }
                    else
                    {
                        TempData["Msg"] = "Oops some error occured";
                        return Json(new { success = false, message = TempData["Msg"] });
                    }
                }
                else
                {
                    TempData["Msg"] = "Please fill all detsils.";
                    return Json(new { success = false, message = TempData["Msg"] });
                }

            }
            catch (Exception ex)
            {
                TempData["Msg"] = ex + "Oops some error occured.";
                return Json(new { success = false, message = TempData["Msg"] });
            }
        }
        #endregion
        #region Index Page Bulk Print home
        public ActionResult Index(GetPrintBadge model)
        {
            try
            {
                if (Session["EventId"] != null)
                {
                    TotalCount();
                    GetPrintBadge EC = new GetPrintBadge();
                    EC.EventId = Convert.ToInt32(Session["EventId"].ToString());
                    EC.IsPrint = model.IsPrint;
                    ds = masterHelper.FetchHomePage("GetAllBadgeSearch", EC, Convert.ToInt16(Session["LoginId"]));
                    if (ds != null && ds.Tables[0].Rows.Count > 0)
                    {
                        EC.BadgeCompany = CommonFunctions.ToList<company>(ds.Tables[0]);
                    }
                    if (ds != null && ds.Tables[1].Rows.Count > 0)
                    {
                        EC.EventWiseCategory = CommonFunctions.ToList<EventCatNew>(ds.Tables[1]);
                    }
                    //if (ds != null && ds.Tables[2].Rows.Count > 0)
                    //{
                    //    EC.ListBadgeCompany = CommonFunctions.ToList<GetPrintBadge>(ds.Tables[2]);
                    //}
                    return View(EC);
                }
                else
                {
                    return RedirectToAction("Login", "Home");
                }
            }
            catch (Exception ex)
            {
                ex.Message.ToString().Trim();
                return RedirectToAction("Login", "Home");
            }
        }
        [HttpGet]
        public ActionResult GetBadgeByFilter(int IsPrint = 0, string category = "", string Company = "")
        {
            string mquery = "";
            GetPrintBadge model = new GetPrintBadge();
            try
            {
                if (Session["EventId"] != null)
                {
                    model.EventId = Convert.ToInt32(Session["EventId"].ToString());
                    if (IsPrint.ToString() == "1")
                    {
                        model.IsPrint = "true";
                    }
                    else
                    {
                        model.IsPrint = "false";
                    }
                    model.Category = category;
                    model.Company = Company;
                    ds = masterHelper.FetchHomePage("GetAllBadgeSearch", model, Convert.ToInt16(Session["LoginId"]));
                    if (ds != null && ds.Tables[2].Rows.Count > 0)
                    {
                        model.ListBadgeCompany = CommonFunctions.ToList<GetPrintBadge>(ds.Tables[2]);
                    }
                    return PartialView("_BadgeFilterData", model.ListBadgeCompany);
                }
                else
                {
                    TempData["Msg"] = "Session Error.Please relogin";
                    return new HttpStatusCodeResult(401, "Session Expired");
                }
            }
            catch (Exception ex)
            {
                TempData["Msg"] = ex + " Oops some error occured. Redirected to login page";
                return new HttpStatusCodeResult(500, "Internal Error: " + ex.Message);
            }

        }
        [HttpGet]
        public JsonResult PrintBadgeBulk(int BadgeId)
        {
            try
            {
                DataSet mds;
                if (Session["EventId"] == null)
                {
                    TempData["Msg"] = "Session expired.Please relogin";
                    return Json(new { success = false, message = TempData["Msg"] });
                }
                else
                {
                    int EventId = Convert.ToInt16(Session["EventId"].ToString());

                    if (EventId != 0 && BadgeId != 0)
                    {
                        ds = masterHelper.BadgePrintDataBulk("DetailBadgePrint", EventId, BadgeId, Convert.ToInt16(Session["LoginId"]));
                        if (ds != null && ds.Tables[0].Rows.Count > 0)
                        {
                            try
                            {
                                StringBuilder sb = new StringBuilder();
                                sb.Append("<input type='hidden' id='badgeid' value='" + BadgeId + "'>");
                                sb.Append("<div id=divprint style='display:none;'>");
                                if (ds.Tables[1].Rows.Count > 0)
                                {
                                    for (int i = 0; ds.Tables[1].Rows.Count > i; i++)
                                    {
                                        for (int j = 0; ds.Tables[1].Rows.Count > j; j++)
                                        {
                                            string mstyleloopcode = $"margin-top: {ds.Tables[1].Rows[j]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[j]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[j]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[j]["MarginRight"].ToString().ToLower()}px;";

                                            if (ds.Tables[1].Rows[j]["columName"].ToString().ToLower() == "photo" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "photo")
                                            {
                                                sb.Append("<div style=\"" + mstyleloopcode + "\"><img src='../Content/BadgePhoto/" + Session["EventName"].ToString() + "/" + ds.Tables[0].Rows[0]["Photo"].ToString() + "' height='75px' width='75px'></div>");
                                            }
                                        }
                                        for (int k = 0; ds.Tables[1].Rows.Count > k; k++)
                                        {
                                            string mstyleloopcode = $"text-align: center; margin-top: {ds.Tables[1].Rows[k]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[k]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[k]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[k]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[1].Rows[k]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[1].Rows[k]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[1].Rows[k]["FontWeight"].ToString().ToLower()};";

                                            if (ds.Tables[1].Rows[k]["columName"].ToString().ToLower() == "name" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "name")
                                            {
                                                sb.Append("<div style=\"" + mstyleloopcode + "\">" + ds.Tables[0].Rows[0]["Name"].ToString() + "</div>");
                                            }
                                        }
                                        for (int l = 0; ds.Tables[1].Rows.Count > l; l++)
                                        {
                                            string mstyleloopcode = $"text-align: center; margin-top: {ds.Tables[1].Rows[l]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[l]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[l]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[l]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[1].Rows[l]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[1].Rows[l]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[1].Rows[l]["FontWeight"].ToString().ToLower()};";

                                            if (ds.Tables[1].Rows[l]["columName"].ToString().ToLower() == "designation" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "designation")
                                            {
                                                sb.Append("<div style=\"" + mstyleloopcode + "\">" + ds.Tables[0].Rows[0]["Designation"].ToString() + "</div>");
                                            }
                                        }
                                        for (int m = 0; ds.Tables[1].Rows.Count > m; m++)
                                        {
                                            string mstyleloopcode = $"text-align: center; margin-top: {ds.Tables[1].Rows[m]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[m]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[m]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[m]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[1].Rows[m]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[1].Rows[m]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[1].Rows[m]["FontWeight"].ToString().ToLower()};";

                                            if (ds.Tables[1].Rows[m]["columName"].ToString().ToLower() == "company" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "company")
                                            {
                                                sb.Append("<div style=\"" + mstyleloopcode + "\">" + ds.Tables[0].Rows[0]["Company"].ToString() + "</div>");
                                            }
                                        }
                                        for (int n = 0; ds.Tables[1].Rows.Count > n; n++)
                                        {
                                            string mstyleloopcode = $"text-align: center; margin-top: {ds.Tables[1].Rows[n]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[n]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[n]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[n]["MarginRight"].ToString().ToLower()}px;";

                                            if (ds.Tables[1].Rows[n]["columName"].ToString().ToLower() == "barcode" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "barcode")
                                            {
                                                sb.Append("<div style=\"" + mstyleloopcode + "\"><img src='../Content/QRCodesScanner/" + Session["EventName"].ToString() + "/" +
                                                    ds.Tables[0].Rows[0]["BPath"].ToString() + "' height='75px' width='75px'>" +
                                                    "<div class=clearfix></div><label style='font-size:9px; padding-bottom:20px  !important;'>" + ds.Tables[0].Rows[0]["BPathCode"].ToString() + "</label></div>");
                                            }
                                        }
                                        for (int o = 0; ds.Tables[1].Rows.Count > o; o++)
                                        {
                                            string mstyleloopcode = $"margin-top: {ds.Tables[1].Rows[o]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[1].Rows[o]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[1].Rows[o]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[1].Rows[o]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[1].Rows[o]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[1].Rows[o]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[1].Rows[o]["FontWeight"].ToString().ToLower()};";

                                            if (ds.Tables[1].Rows[o]["columName"].ToString().ToLower() == "category" && ds.Tables[1].Rows[i]["columName"].ToString().ToLower() == "category")
                                            {
                                                sb.Append("<div style=\"" + mstyleloopcode + "\" id='mcat'>" + ds.Tables[0].Rows[0]["Category"].ToString() + "</div>");
                                            }
                                        }

                                    }
                                }
                                sb.Append("</div>");
                                sb.Append("</div>");
                                string mBadge = sb.ToString();
                                TotalCount();
                                return Json(new { success = true, mData = mBadge }, JsonRequestBehavior.AllowGet);
                            }
                            catch (Exception ex)
                            {
                                TempData["Msg"] = ex.Message + " Oops some error occured";
                                return Json(new { success = false, message = TempData["Msg"] });
                            }
                        }
                        else
                        {
                            TempData["Msg"] = "Oops some error occured";
                            return Json(new { success = false, message = TempData["Msg"] });
                        }
                    }
                    else
                    {
                        TempData["Msg"] = "Please fill all detsils.";
                        return Json(new { success = false, message = TempData["Msg"] });
                    }
                }
            }
            catch (Exception ex)
            {
                TempData["Msg"] = ex + "Oops some error occured.";
                return Json(new { success = false, message = TempData["Msg"] });
            }
        }
        #endregion
        #region login code
        [HttpPost]
        [ValidateAntiForgeryToken]
        [OutputCache(NoStore = true, Duration = 0)]
        public ActionResult LoginCode(LoginModels model)
        {
            try
            {
                string ActionAttempt = "Login";
                DataSet ds = masterHelper.FetchLogin(ActionAttempt, model);
                if (ds.Tables[1].Rows[0]["DB_STATUS"].ToString() == "S")
                {
                    string redirect = "";
                    string mcontroler = "";
                    Session["LoginId"] = ds.Tables[0].Rows[0]["LoginId"].ToString();
                    Session["UserEmail"] = model.UserName;
                    Session["RoleType"] = ds.Tables[0].Rows[0]["RoleAbbr"].ToString();
                    if (ds.Tables[0].Rows[0]["EventId"].ToString() != "")
                    {
                        Session["EventId"] = ds.Tables[0].Rows[0]["EventId"].ToString();
                        Session["EventName"] = ds.Tables[0].Rows[0]["EventName"].ToString();
                    }
                    if (ds.Tables[0].Rows[0]["RoleAbbr"].ToString().Trim() == "GS")
                    {
                        redirect = "Scanning";
                        mcontroler = "Home";

                    }
                    else if (ds.Tables[0].Rows[0]["RoleAbbr"].ToString().Trim() == "S")
                    {
                        redirect = "MasterSearchBadge";
                        mcontroler = "Home";
                    }
                    else if (ds.Tables[0].Rows[0]["RoleAbbr"].ToString().Trim() == "A" || ds.Tables[0].Rows[0]["RoleAbbr"].ToString().Trim() == "M")
                    {
                        redirect = "AdminDashboard";
                        mcontroler = "Admin";
                    }
                    return RedirectToAction(redirect, mcontroler);
                }
                else
                {
                    TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    return RedirectToAction("Login", "Home");
                }
            }
            catch (Exception ex)
            {
                TempData["Msg"] = "!!! Something went wrong !!!. Please contact admin. " + ex.Message.ToString();
                return RedirectToAction("Login", "Home");
            }
        }
        public ActionResult Login()
        {
            return View();
        }
        public ActionResult LogOut()
        {
            Session.Abandon();
            Session.Clear();
            Session.RemoveAll();
            return View("Login");
        }
        #endregion
        #region master page add
        public ActionResult AddNewBadge()
        {
            try
            {
                if (Session["EventId"] != null)
                {
                    TotalCount();
                    AddBadgeCode EC = new AddBadgeCode();
                    EC.EventId = Convert.ToInt32(Session["EventId"].ToString());
                    ds = masterHelper.FetchAddBadgePrintNew("GetBadgeData4Show", EC, Convert.ToInt16(Session["LoginId"]));
                    if (ds != null && ds.Tables[0].Rows.Count > 0)
                    {
                        EC.masterBadgeSave = CommonFunctions.ToList<AddBadgeCode>(ds.Tables[0]);
                        if (ds != null && ds.Tables[0].Rows.Count > 0)
                        {
                            EC.EventCatAll = CommonFunctions.ToList<EventCatNew>(ds.Tables[1]);
                        }
                        if (ds != null && ds.Tables[0].Rows.Count > 0)
                        {
                            EC.mastercountry = CommonFunctions.ToList<EventCountry>(ds.Tables[2]);
                        }
                    }
                    return View(EC);
                }
                else
                {
                    return RedirectToAction("Login", "Home");
                }
            }
            catch (Exception ex)
            {
                ex.Message.ToString().Trim();
                // return RedirectToAction("Login", "Home");
                throw ex;
            }
        }
        
        [HttpPost]
        public JsonResult MasterBadgeSaveCodeandPrint(allbadgebyevent modal,HttpPostedFileBase Photo,string SelectedFields,string Name,string Designation,string Company,
        string Category,string Mobile,string Email,string Country,string Remarks)
        {
            try
            {
                if (Session["EventId"] == null)
                    return Json(new { success = false, message = "Session expired." });
                modal.EventId = Convert.ToInt32(Session["EventId"]);
                modal.Name = Name;
                modal.Designation = Designation;
                modal.Company = Company;
                modal.Category = Category;
                modal.Mobile = Mobile;
                modal.Email = Email;
                modal.Country = Country;
                modal.Remarks = Remarks;
                if (Photo != null && Photo.ContentLength > 0)
                {
                    string folder = Server.MapPath("~/Content/BadgePhoto/" + Session["EventName"]);
                    if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
                    string ext = Path.GetExtension(Photo.FileName);
                    string filename = DateTime.Now.Ticks + ext;
                    string full = Path.Combine(folder, filename);
                    Photo.SaveAs(full);
                    modal.Photo = filename;
                }
                DataSet ds = masterHelper.SaveMasterBadgeOnebyOne(
                    "SaveBadge", modal, Convert.ToInt32(Session["LoginId"]));
                if (ds == null || ds.Tables[0].Rows[0]["DB_STATUS"].ToString() != "S")
                    return Json(new { success = false, message = "Save failed" });
                int badgeId = Convert.ToInt32(ds.Tables[1].Rows[0]["LastInsertedID"]);
                string qrCode = ds.Tables[1].Rows[0]["BPathCode"].ToString();
                string qrFolder = Server.MapPath("~/Content/QRCodesScanner/" + Session["EventName"]);
                if (!Directory.Exists(qrFolder)) Directory.CreateDirectory(qrFolder);
                string qrFile = qrCode + ".png";
                string qrFull = Path.Combine(qrFolder, qrFile);
                using (QRCodeGenerator qg = new QRCodeGenerator())
                using (QRCodeData qd = qg.CreateQrCode(qrCode, QRCodeGenerator.ECCLevel.Q))
                using (QRCode qr = new QRCode(qd))
                using (Bitmap bmp = qr.GetGraphic(20))
                {
                    bmp.Save(qrFull, ImageFormat.Png);
                }
                modal.BadgeId = badgeId;
                modal.BPath = qrFile;
                masterHelper.SaveMasterBadgeOnebyOne("UpdateQrCodePath", modal,
                    Convert.ToInt32(Session["LoginId"]));
                BadgePrintVM vm = new BadgePrintVM
                {
                    BadgeId = badgeId,
                    Name = modal.Name,
                    Company = modal.Company,
                    Category = modal.Category,
                    Designation = modal.Designation,
                    Mobile = modal.Mobile,
                    Email = modal.Email,
                    Country = modal.Country,
                    Remarks = modal.Remarks,
                    PhotoPath = "/Content/BadgePhoto/" + Session["EventName"] + "/" + modal.Photo,
                    QRPath = "/Content/QRCodesScanner/" + Session["EventName"] + "/" + qrFile,
                    QRText = qrCode,
                    BgImage = "/Content/EventBadge/" + ds.Tables[3].Rows[0]["EventBadgePhoto"],
                    BadgeWidth = Convert.ToInt32(ds.Tables[3].Rows[0]["BadgeWidth"]),
                    BadgeHeight = Convert.ToInt32(ds.Tables[3].Rows[0]["BadgeHeight"]),
                    LayoutTable = ds.Tables[2],
                    FieldsToPrint = SelectedFields.Split(',').ToList()
                };
                string html = RenderPartialViewToString("_AddBadge", vm);
                return Json(new { success = true, html = html });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        protected string RenderPartialViewToString(string viewName, object model)
        {
            if (string.IsNullOrEmpty(viewName))
                viewName = ControllerContext.RouteData.GetRequiredString("action");
            ViewData.Model = model;
            using (StringWriter sw = new StringWriter())
            {
                ViewEngineResult viewResult = ViewEngines.Engines
                    .FindPartialView(ControllerContext, viewName);
                if (viewResult.View == null)
                    throw new FileNotFoundException("Partial view '" + viewName + "' not found.");
                ViewContext viewContext = new ViewContext(ControllerContext,viewResult.View,ViewData,TempData,sw);
                viewResult.View.Render(viewContext, sw);
                return sw.ToString();
            }
        }

        [HttpPost]
        public JsonResult updatebadgeafterprint(int BadgeId)
        {
            try
            {
                ds = masterHelper.PrintBadgeCode("UpdateStatusofPrint", BadgeId, "", "", Convert.ToInt16(Session["EventId"]), "");
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    return Json(new { success = true }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return Json(new { success = false }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false }, JsonRequestBehavior.AllowGet);
            }
        }
        public JsonResult updatebadgeafterReprint(int BadgeId)
        {
            try
            {
                ds = masterHelper.PrintBadgeCode("UpdateStatusofRePrint", BadgeId, "", "", Convert.ToInt16(Session["EventId"]), "");
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    return Json(new { success = true }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return Json(new { success = false }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false }, JsonRequestBehavior.AllowGet);
            }
        }
        public JsonResult updatebadgeafterbulkprintstatus(List<int> multiprint, int mstatus)
        {
            try
            {
                if (multiprint.Count > 0)
                {
                    for (int g = 0; g < multiprint.Count; g++)
                    {
                        newbadgeid = newbadgeid + "," + multiprint[g];
                    }
                }
                if (mstatus == 0)
                {
                    ds = masterHelper.PrintBadgeCode("MultiprintStatusupdate", 0, "", "", Convert.ToInt16(Session["EventId"]), newbadgeid.Substring(1));
                }
                else
                {
                    ds = masterHelper.PrintBadgeCode("Multireprint", 0, "", "", Convert.ToInt16(Session["EventId"]), newbadgeid.Substring(1));
                }
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    return Json(new { success = true }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return Json(new { success = false }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false }, JsonRequestBehavior.AllowGet);
            }
        }
        #endregion
        #region scanning
        public ActionResult Scanning()
        {
            TotalCount();
            return View();
        }

        [HttpGet]
        public ActionResult ScanningData(string input)
        {
            try
            {
                if (Session["EventId"] != null)
                {
                    BadgeDetailsScanning EC = new BadgeDetailsScanning();
                    EC.SearchText = input;
                    EC.EventId = Convert.ToInt32(Session["EventId"].ToString());
                    ds = masterHelper.FetchIdByBatch("DetailsByBadge", EC, Convert.ToInt16(Session["LoginId"]));
                    if (ds != null && ds.Tables[0].Rows.Count > 0)
                    {
                        EC.mBadgeList = CommonFunctions.ToList<BadgeDetailsScanning>(ds.Tables[0]);
                    }
                    ViewBag.KitTotal = ds.Tables[1].Rows[0]["TotalBadge"].ToString();
                    return PartialView("_BadgeShow", EC.mBadgeList);
                }
                else
                {
                    TempData["Msg"] = "Session Error.Please relogin";
                    return new HttpStatusCodeResult(401, "Session Expired");
                }
            }
            catch (Exception ex)
            {
                TempData["Msg"] = ex + " Oops some error occured. Redirected to login page";
                return new HttpStatusCodeResult(500, "Internal Error: " + ex.Message);
            }
        }

        [HttpPost]
        public ActionResult updatebadgekitstatus(int BadgeId, string CollectBy, string Remarks)
        {
            try
            {
                if (BadgeId != 0 && CollectBy != null)
                {
                    ds = masterHelper.PrintBadgeCode("UpdateStatusofKit", BadgeId, CollectBy, Remarks, Convert.ToInt16(Session["EventId"]), "");
                    if (ds != null && ds.Tables[0].Rows.Count > 0)
                    {
                        return RedirectToAction("Scanning", "Home");
                    }
                    else
                    {
                        TempData["Message"] = "Oops some error occured.";
                        return RedirectToAction("Scanning", "Home");
                    }
                }
                else
                {
                    TempData["Message"] = "Please fill valid details";
                    return RedirectToAction("Scanning", "Home");
                }
            }
            catch (Exception ex)
            {
                TempData["Message"] = ex.Message + "Oops some error occured";
                return RedirectToAction("Scanning", "Home");
            }
        }
        #endregion
        #region reoprt
        public ActionResult Report()
        {
            TotalCount();
            allbadgebyevent EC = new allbadgebyevent();
            ds = masterHelper.FetchBadgeEvent("ViewAllbadgeEvent", EC, Convert.ToInt16(Session["LoginId"]));
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                EC.EventExcelUpload = CommonFunctions.ToList<mExcel>(ds.Tables[0]);
            }
            return View(EC);
        }
        public ActionResult DownloadExcelEvent(int EventId, string EventName)
        {
            allbadgebyevent EC = new allbadgebyevent();
            EC.EventId = EventId;
            ds = masterHelper.FetchBadgeEvent("GetAllBadgeofEvent", EC, Convert.ToInt16(Session["LoginId"]));
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("EventData_" + EventId + DateTime.Now.ToString("dd-MM-yy"));
                    worksheet.Cell(1, 1).Value = "Category";
                    worksheet.Cell(1, 2).Value = "QRCodeNo";
                    worksheet.Cell(1, 3).Value = "Name";
                    worksheet.Cell(1, 4).Value = "Designation";
                    worksheet.Cell(1, 5).Value = "Company";
                    worksheet.Cell(1, 6).Value = "Mobile";
                    worksheet.Cell(1, 7).Value = "Email";
                    worksheet.Cell(1, 8).Value = "Country";
                    //worksheet.Cell(1, 9).Value = "ID Type";
                    //worksheet.Cell(1, 10).Value = "ID Number";
                    worksheet.Cell(1, 9).Value = "Print/Reprint Status";
                    worksheet.Cell(1, 10).Value = "Visit On";
                    worksheet.Cell(1, 11).Value = "Kit Status";
                    worksheet.Cell(1, 12).Value = "Payment Status";
                    worksheet.Cell(1, 13).Value = "Record Insert on";
                    worksheet.Cell(1, 14).Value = "Remarks";

                    var headerRange = worksheet.Range("A1:N1");
                    headerRange.Style.Fill.BackgroundColor = XLColor.Yellow;
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    int row = 2;
                    foreach (DataRow Event in ds.Tables[0].Rows)
                    {
                        worksheet.Cell(row, 1).Value = Event[2].ToString();
                        worksheet.Cell(row, 2).Value = Event[26].ToString();
                        worksheet.Cell(row, 3).Value = Event[4].ToString();
                        worksheet.Cell(row, 4).Value = Event[5].ToString();
                        worksheet.Cell(row, 5).Value = Event[6].ToString();
                        worksheet.Cell(row, 6).Value = Event[7].ToString();
                        worksheet.Cell(row, 7).Value = Event[8].ToString();
                        worksheet.Cell(row, 8).Value = Event[9].ToString();
                        //worksheet.Cell(row, 9).Value = Event[10].ToString();
                        //worksheet.Cell(row, 10).Value = Event[11].ToString();
                        worksheet.Cell(row, 9).Value = Event[13].ToString() + "/" + Event[15].ToString();
                        worksheet.Cell(row, 10).Value = Event[21].ToString();
                        worksheet.Cell(row, 11).Value = Event[22].ToString() + "/" + Event[24].ToString() + "/" + Event[23].ToString();
                        if(Event[28].ToString().ToLower()=="true")
                        { 
                        worksheet.Cell(row, 12).Value = "Yes" + "/" + Event[27].ToString();
                        }
                        else
                        {
                            worksheet.Cell(row, 12).Value = "No" + "/" + Event[27].ToString();
                        }
                        worksheet.Cell(row, 13).Value = Event[25].ToString();
                        worksheet.Cell(row, 14).Value = Event[17].ToString();
                        row++;
                    }

                    var usedRange = worksheet.RangeUsed();
                    usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        stream.Position = 0;
                        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "EventData_" + EventName + ".xlsx");
                    }
                }
            }
            return RedirectToAction("Report", "Home");
        }

        public ActionResult DownloadExcelEventYesOnly(int EventId, string EventName)
        {
            allbadgebyevent EC = new allbadgebyevent();
            EC.EventId = EventId;
            ds = masterHelper.FetchBadgeEvent("GetAllBadgeofEventYesOnly", EC, Convert.ToInt16(Session["LoginId"]));
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("EventData_" + EventId + DateTime.Now.ToString("dd-MM-yy"));
                    worksheet.Cell(1, 1).Value = "Category";
                    worksheet.Cell(1, 2).Value = "QRCodeNo";
                    worksheet.Cell(1, 3).Value = "Name";
                    worksheet.Cell(1, 4).Value = "Designation";
                    worksheet.Cell(1, 5).Value = "Company";
                    worksheet.Cell(1, 6).Value = "Mobile";
                    worksheet.Cell(1, 7).Value = "Email";
                    worksheet.Cell(1, 8).Value = "Country";
                   // worksheet.Cell(1, 9).Value = "ID Type";
                   // worksheet.Cell(1, 10).Value = "ID Number";
                    worksheet.Cell(1, 9).Value = "Print/Reprint Status";
                    worksheet.Cell(1, 10).Value = "Visit On";
                    worksheet.Cell(1, 11).Value = "Kit Status";
                    worksheet.Cell(1, 12).Value = "Payment Status";
                    worksheet.Cell(1, 13).Value = "Record Insert on";
                    worksheet.Cell(1, 14).Value = "Remarks";

                    var headerRange = worksheet.Range("A1:N1");
                    headerRange.Style.Fill.BackgroundColor = XLColor.Yellow;
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


                    int row = 2;
                    foreach (DataRow Event in ds.Tables[0].Rows)
                    {
                        worksheet.Cell(row, 1).Value = Event[2].ToString();
                        worksheet.Cell(row, 2).Value = Event[26].ToString();
                        worksheet.Cell(row, 3).Value = Event[4].ToString();
                        worksheet.Cell(row, 4).Value = Event[5].ToString();
                        worksheet.Cell(row, 5).Value = Event[6].ToString();
                        worksheet.Cell(row, 6).Value = Event[7].ToString();
                        worksheet.Cell(row, 7).Value = Event[8].ToString();
                        worksheet.Cell(row, 8).Value = Event[9].ToString();
                      //  worksheet.Cell(row, 9).Value = Event[10].ToString();
                      //  worksheet.Cell(row, 10).Value = Event[11].ToString();
                        worksheet.Cell(row, 9).Value = Event[13].ToString() + "/" + Event[15].ToString();
                        worksheet.Cell(row, 10).Value = Event[21].ToString();
                        worksheet.Cell(row, 11).Value = Event[22].ToString() + "/" + Event[24].ToString() + "/" + Event[23].ToString();
                        if (Event[28].ToString().ToLower() == "true")
                        {
                            worksheet.Cell(row, 12).Value = "Yes" + "/" + Event[27].ToString();
                        }
                        else
                        {
                            worksheet.Cell(row, 12).Value = "No" + "/" + Event[27].ToString();
                        }
                        worksheet.Cell(row, 13).Value = Event[25].ToString();
                        worksheet.Cell(row, 14).Value = Event[17].ToString();
                        row++;
                    }

                    var usedRange = worksheet.RangeUsed();
                    usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;


                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        stream.Position = 0;

                        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "EventData_" + EventName + ".xlsx");
                    }
                }
            }
            return RedirectToAction("Report", "Home");
        }

        public ActionResult EmailExcel(int EventId, string EventName, string SendtoEmailId)
        {
            allbadgebyevent EC = new allbadgebyevent();
            EC.EventId = EventId;
            ds = masterHelper.FetchBadgeEvent("GetAllBadgeofEvent", EC, Convert.ToInt16(Session["LoginId"]));
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("EventData_" + EventName);
                    worksheet.Cell(1, 1).Value = "Category";
                    worksheet.Cell(1, 2).Value = "QRCodeNo";
                    worksheet.Cell(1, 3).Value = "Name";
                    worksheet.Cell(1, 4).Value = "Designation";
                    worksheet.Cell(1, 5).Value = "Company";
                    worksheet.Cell(1, 6).Value = "Mobile";
                    worksheet.Cell(1, 7).Value = "Email";
                    worksheet.Cell(1, 8).Value = "Country";
                    worksheet.Cell(1, 9).Value = "Print/Reprint Status";
                    worksheet.Cell(1, 10).Value = "Visit On";
                    worksheet.Cell(1, 11).Value = "Kit Status";
                    worksheet.Cell(1, 12).Value = "Payment Status";
                    worksheet.Cell(1, 13).Value = "Record Insert on";
                    worksheet.Cell(1, 14).Value = "Remarks";

                    var headerRange = worksheet.Range("A1:N1");
                    headerRange.Style.Fill.BackgroundColor = XLColor.Yellow;
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


                    int row = 2;
                    foreach (DataRow Event in ds.Tables[0].Rows)
                    {
                        worksheet.Cell(row, 1).Value = Event[2].ToString();
                        worksheet.Cell(row, 2).Value = Event[26].ToString();
                        worksheet.Cell(row, 3).Value = Event[4].ToString();
                        worksheet.Cell(row, 4).Value = Event[5].ToString();
                        worksheet.Cell(row, 5).Value = Event[6].ToString();
                        worksheet.Cell(row, 6).Value = Event[7].ToString();
                        worksheet.Cell(row, 7).Value = Event[8].ToString();
                        worksheet.Cell(row, 8).Value = Event[9].ToString();
                        worksheet.Cell(row, 9).Value = Event[13].ToString() + "/" + Event[15].ToString();
                        worksheet.Cell(row, 10).Value = Event[21].ToString();
                        worksheet.Cell(row, 11).Value = Event[22].ToString() + "/" + Event[24].ToString() + "/" + Event[23].ToString();
                        if (Event[28].ToString().ToLower() == "true")
                        {
                            worksheet.Cell(row, 12).Value = "Yes" + "/" + Event[27].ToString();
                        }
                        else
                        {
                            worksheet.Cell(row, 12).Value = "No" + "/" + Event[27].ToString();
                        }
                        worksheet.Cell(row, 13).Value = Event[25].ToString();
                        worksheet.Cell(1, 14).Value = Event[17].ToString();
                        row++;
                    }

                    var usedRange = worksheet.RangeUsed();
                    usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        stream.Position = 0;
                        var message = new MailMessage();
                        message.From = new MailAddress("scan.namaste@gmail.com");
                        message.To.Add(SendtoEmailId);
                        message.Subject = "Event Badge Report - " + EventName;
                        message.Body = "Please find attached the Excel report for event: " + EventName;
                        stream.Position = 0;
                        var attachment = new Attachment(stream, "EventData_" + EventName + ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                        message.Attachments.Add(attachment);
                        var smtp = new SmtpClient("smtp.gmail.com", 587)
                        {
                            Credentials = new NetworkCredential("scan.namaste@gmail.com", "atrhkgnvnfvhpwda"), //Namaste1771
                            EnableSsl = true
                        };
                        smtp.Send(message);
                    }
                    TempData["Message"] = "Email sent with Excel attachment.";
                }
            }
            else
            {
                TempData["Message"] = "No data found to send.";
            }
            return RedirectToAction("Report", "Home");
        }
        #endregion
        #region   total count
        protected void TotalCount()
        {
            try
            {
                if (Session["EventId"] != null)
                {
                    string ActionAttempt = "CountTotal";
                    DataSet ds = masterHelper.FetchCountData(ActionAttempt, Convert.ToInt32(Session["EventId"].ToString()));
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        ViewBag.TotalBadgePrint = ds.Tables[0].Rows[0]["TotalBadgePrint"].ToString();
                        ViewBag.TotalKit = ds.Tables[0].Rows[0]["TotalKit"].ToString();
                        ViewBag.TotalBadge = ds.Tables[0].Rows[0]["TotalBadge"].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                ex.Message.ToString();
            }
        }
        #endregion
        #region Kiosk
        public ActionResult PrintwithKiosk()
        {
            Kiosk EC = new Kiosk();
            ds = masterHelper.FetchEventKiosk("Kiosk", EC, Convert.ToInt16(Session["LoginId"]));
            if (ds != null && ds.Tables.Count > 1 && ds.Tables[1].Rows.Count > 0)
            {
                DataView dv = new DataView(ds.Tables[1]);
                dv.RowFilter = "EventId = " + Session["EventId"];
                DataTable dtnew = dv.ToTable();
                if (dtnew.Rows.Count > 0)
                {
                    ViewBag.bgimg = dtnew.Rows[0]["KioskBg"].ToString();
                    ViewBag.Title = dtnew.Rows[0]["Title"].ToString();
                }
            }
            return View();
        }

        public ActionResult KioskDataShow(string input)
        {
            try
            {
                if (Session["EventId"] != null)
                {
                    BadgeDetailsScanning EC = new BadgeDetailsScanning();
                    EC.SearchText = input;
                    EC.EventId = Convert.ToInt32(Session["EventId"].ToString());
                    ds = masterHelper.FetchIdByBatch("DetailsByBadge", EC, Convert.ToInt16(Session["LoginId"]));
                    if (ds != null && ds.Tables[0].Rows.Count > 0)
                    {
                        EC.mBadgeList = CommonFunctions.ToList<BadgeDetailsScanning>(ds.Tables[0]);
                    }
                    ViewBag.KitTotal = ds.Tables[1].Rows[0]["TotalBadge"].ToString();
                    return PartialView("_KioskData", EC.mBadgeList);
                }
                else
                {
                    TempData["Msg"] = "Session Error.Please relogin";
                    return new HttpStatusCodeResult(401, "Session Expired");
                }
            }
            catch (Exception ex)
            {
                TempData["Msg"] = ex + " Oops some error occured. Redirected to login page";
                return new HttpStatusCodeResult(500, "Internal Error: " + ex.Message);
            }
        }

        #endregion
    }

}
