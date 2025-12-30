using iTextSharp.text.html;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using Document = iTextSharp.text.Document;
using Font = iTextSharp.text.Font;
using Image = iTextSharp.text.Image;
using EventManagement.Models;

namespace EventManagement.Controllers
{
    public class PrintController : Controller
    {
        // GET: Print
        #region Mom pdf

        public void PrintCodeBadgeSingle(int id = 0, int eventid = 0)
        {
            StringBuilder sb = new StringBuilder();
            Master_Helper PrintHelper = new Master_Helper();
            DataSet ds = PrintHelper.PrintBadgeCode("Get Date PrintSetup", id, "", "", eventid,"");
            sb.Append("<div id=divprint>");
            // Add content as absolutely positioned inside the badge
            if (ds.Tables[2].Rows.Count > 0)
            {
                for (int i = 0; ds.Tables[2].Rows.Count > i; i++)
                {
                    for (int j = 0; ds.Tables[2].Rows.Count > j; j++)
                    {
                        string mstyleloopcode = $"margin-top: {ds.Tables[2].Rows[j]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[2].Rows[j]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[2].Rows[j]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[2].Rows[j]["MarginRight"].ToString().ToLower()}px;";

                        if (ds.Tables[2].Rows[j]["columName"].ToString().ToLower() == "photo" && ds.Tables[2].Rows[i]["columName"].ToString().ToLower() == "photo")
                        {
                            sb.Append("<div style=\"" + mstyleloopcode + "\"><img src='../Content/BadgePhoto/Dummybadgephoto.jpg' height='75px' width='75px'></div>");
                        }
                    }
                    for (int k = 0; ds.Tables[2].Rows.Count > k; k++)
                    {
                        string mstyleloopcode = $"margin-top: {ds.Tables[2].Rows[k]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[2].Rows[k]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[2].Rows[k]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[2].Rows[k]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[2].Rows[k]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[2].Rows[k]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[2].Rows[k]["FontWeight"].ToString().ToLower()};";

                        if (ds.Tables[2].Rows[k]["columName"].ToString().ToLower() == "name" && ds.Tables[2].Rows[i]["columName"].ToString().ToLower() == "name")
                        {
                            sb.Append("<div style=\"" + mstyleloopcode + "\">Display Exhibitor Name</div>");
                        }
                    }
                    for (int l = 0; ds.Tables[2].Rows.Count > l; l++)
                    {
                        string mstyleloopcode = $"margin-top: {ds.Tables[2].Rows[l]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[2].Rows[l]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[2].Rows[l]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[2].Rows[l]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[2].Rows[l]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[2].Rows[l]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[2].Rows[l]["FontWeight"].ToString().ToLower()};";

                        if (ds.Tables[2].Rows[l]["columName"].ToString().ToLower() == "designation" && ds.Tables[2].Rows[i]["columName"].ToString().ToLower() == "designation")
                        {
                            sb.Append("<div style=\"" + mstyleloopcode + "\">Display Designation Name</div>");
                        }
                    }
                    for (int m = 0; ds.Tables[2].Rows.Count > m; m++)
                    {
                        string mstyleloopcode = $"margin-top: {ds.Tables[2].Rows[m]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[2].Rows[m]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[2].Rows[m]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[2].Rows[m]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[2].Rows[m]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[2].Rows[m]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[2].Rows[m]["FontWeight"].ToString().ToLower()};";

                        if (ds.Tables[2].Rows[m]["columName"].ToString().ToLower() == "company" && ds.Tables[2].Rows[i]["columName"].ToString().ToLower() == "company")
                        {
                            sb.Append("<div style=\"" + mstyleloopcode + "\">Display Company Name</div>");
                        }
                    }
                    for (int n = 0; ds.Tables[2].Rows.Count > n; n++)
                    {
                        string mstyleloopcode = $"margin-top: {ds.Tables[2].Rows[n]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[2].Rows[n]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[2].Rows[n]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[2].Rows[n]["MarginRight"].ToString().ToLower()}px;";

                        if (ds.Tables[2].Rows[n]["columName"].ToString().ToLower() == "barcode" && ds.Tables[2].Rows[i]["columName"].ToString().ToLower() == "barcode")
                        {
                            sb.Append("<div style=\"" + mstyleloopcode + "\"><img src='../Content/QRCodesScanner/DummyQR.png' height='75px' width='75px'></div>");
                        }
                    }
                    for (int o = 0; ds.Tables[2].Rows.Count > o; o++)
                    {
                        string mstyleloopcode = $"margin-top: {ds.Tables[2].Rows[o]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[2].Rows[o]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[2].Rows[o]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[2].Rows[o]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[2].Rows[o]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[2].Rows[o]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[2].Rows[o]["FontWeight"].ToString().ToLower()};";

                        if (ds.Tables[2].Rows[o]["columName"].ToString().ToLower() == "category" && ds.Tables[2].Rows[i]["columName"].ToString().ToLower() == "category")
                        {
                            sb.Append("<div style=\"" + mstyleloopcode + "\" id='mcat'>Display Category Name</div>");
                        }
                    }
                }
            }
            else
            {
                if (ds.Tables[1].Rows[0]["Photo"].ToString().ToLower() == "true")
                {
                    sb.Append("<div id='divimage'><img src='../Content/BadgePhoto/Dummybadgephoto.jpg' height='75px' width='75px'></div>");
                }
                if (ds.Tables[1].Rows[0]["RName"].ToString().ToLower() == "true")
                {
                    sb.Append("<div  id='divname'>Display Exhibitor Name</div>");
                }
                if (ds.Tables[1].Rows[0]["Designation"].ToString().ToLower() == "true")
                {
                    sb.Append("<div  id='divdesignation'>Designation Name</div>");
                }
                if (ds.Tables[1].Rows[0]["Company"].ToString().ToLower() == "true")
                {
                    sb.Append("<div  id='divcompany'>Company Name</div>");
                }
                if (ds.Tables[1].Rows[0]["BarCode"].ToString().ToLower() == "true")
                {
                    sb.Append("<div  id='divbarcode'><img src='../Content/QRCodesScanner/DummyQR.png' height='75px' width='75px'></div>");
                }
                if (ds.Tables[1].Rows[0]["Category"].ToString().ToLower() == "true")
                {
                    sb.Append("<div  id='divcartegory'>Category Name</div>");
                }
            }
            sb.Append("</div>");
            // This actually makes your HTML output to be downloaded as .xls file
            Response.Clear();
            Response.ClearContent();
            Response.ContentType = "application/pdf";
            Response.AddHeader("Content-Disposition", "attachment; filename=PrintBadge.pdf");
            Response.Charset = "utf-8";
            Response.ContentEncoding = System.Text.Encoding.GetEncoding("windows-1250");
            //sets font
            // Response.Write("<font style='font-size:10.0pt; font-family:Calibri;'>");
            Response.Write(sb);
            // Response.Write("</font>");
            Response.Flush();
            Response.End();
        }
        #endregion       
    }
}