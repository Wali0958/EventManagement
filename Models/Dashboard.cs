using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EventManagement.Models
{
    public class Dashboard
    {
        public int TotalEvent { get; set; }
        public int TotalUser { get; set; }
        public int TotalBadge { get; set; }
        public int TotalBadgeNo { get; set; }
        public int Id { get; set; }
        public int IsActive { get; set; }

        public List<mExcel> EventExcelUpload { get; set; }
    }
    
}