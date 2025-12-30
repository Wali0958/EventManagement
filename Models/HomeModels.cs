using DocumentFormat.OpenXml.Drawing.Diagrams;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using static ClosedXML.Excel.XLPredefinedFormat;

namespace EventManagement.Models
{
    public class HomeModels
    {

    }
    public class GetPrintBadge
    {
        public int BadgeId { get; set; }
        public int EventId { get; set; }
        public string Category { get; set; }
        public string SpecialNo { get; set; }
        public string Name { get; set; }
        public string Designation { get; set; }
        public string Company { get; set; }
        public string Mobile { get; set; } = "0000000000";
        public string Email { get; set; } = "0000000000";
        public string Photo { get; set; } = null;
        public System.DateTime RecInsTime { get; set; }
        public string IsPrint { get; set; }
        public System.DateTime PrintOn { get; set; }
        public string KitCollect { get; set; }
        public System.DateTime CollectOn { get; set; }

        public string Country { get; set; }
        public string IDType { get; set; }
        public string IDNumber { get; set; }

        public bool IsReprint { get; set; }
        public System.DateTime ReprintOn { get; set; }
        public string Remarks { get; set; }
        public bool DataInsertby { get; set; }
        public bool IsActive { get; set; }
        public bool IsVisit { get; set; }
        public System.DateTime VisitOn { get; set; }
        public string CollectBy { get; set; }

        public string SearchText { get; set; }
        public List<EventCatNew> EventWiseCategory { get; set; }
        public List<company> BadgeCompany { get; set; }
        public List<GetPrintBadge> ListBadgeCompany { get; set; }
    }
    public class company
    {
        public int EventId { get; set; }
        public string Company { get; set; }
    }
    public class AddBadgeCode
    {
        public int EventId { get; set; }
        public bool Category { get; set; }
        public bool SpecialNo { get; set; }
        public bool RName { get; set; }
        public bool Designation { get; set; }
        public bool Company { get; set; }
        public bool Mobile { get; set; }
        public bool Email { get; set; }
        public bool Country { get; set; }
        public bool IdType { get; set; }
        public bool IdNumber { get; set; }
        public bool Photo { get; set; }
        public bool Remarks { get; set; }
        public bool IsActive { get; set; }

        public List<EventCatNew> EventCatAll { get; set; }
        public List<AddBadgeCode> masterBadgeSave { get; set; }
        public List<EventCountry> mastercountry { get; set; }
        public int BadgeId { get; set; }
    }
    public class BadgeDetailsScanning
    {
        public int BadgeId { get; set; }
        public int EventId { get; set; }
        public string EventName { get; set; }
        public string Category { get; set; }
        public string SpecialNo { get; set; }
        public string Name { get; set; }
        public string Designation { get; set; }
        public string Company { get; set; }
        public string Mobile { get; set; }
        public string Country { get; set; }
        public string IDType { get; set; }
        public string IDNumber { get; set; }
        public string Photo { get; set; }
        public string IsPrint { get; set; }
        public System.DateTime PrintOn { get; set; }
        public string IsReprint { get; set; }
        public System.DateTime ReprintOn { get; set; }
        public string Remarks { get; set; }
        public string DataInsertby { get; set; }
        public System.DateTime RecInsTime { get; set; }
        public string IsActive { get; set; }
        public string IsVisit { get; set; }
        public System.DateTime VisitOn { get; set; }
        public string KitCollect { get; set; }
        public System.DateTime CollectOn { get; set; }
        public string CollectBy { get; set; }
        public string IsKitAllow { get; set; }
        public string IsPaymentAllow { get; set; }
        public bool PaymentStatus { get; set; }
        public System.DateTime PaymentOn { get; set; }
        public string SearchText { get; set; }
        public List<BadgeDetailsScanning> mBadgeList { get; set; }
    }



    //chat gpt code for print badge 
    public class BadgePrintVM
    {
        public int BadgeId { get; set; }

        public string BgImage { get; set; }
        public string PhotoPath { get; set; }
        public string QRPath { get; set; }
        public string QRText { get; set; }

        public string Name { get; set; }
        public string Company { get; set; }
        public string Category { get; set; }
        public string Designation { get; set; }
        public string Mobile { get; set; }
        public string Email { get; set; }
        public string Country { get; set; }
        public string Remarks { get; set; }

        public List<string> FieldsToPrint { get; set; }

        // Dynamic layout (ds.Tables[2])
        public System.Data.DataTable LayoutTable { get; set; }

        public int BadgeWidth { get; set; }
        public int BadgeHeight { get; set; }
    }


}