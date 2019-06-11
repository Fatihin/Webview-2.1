using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WebView.Models
{
    public class MileageRatesViewModel
    {


        public string Grade { get; set; }
        public string GradeDescription { get; set; }
        public int LowLev1 { get; set; }
        public int HighLev1 { get; set; }
        public int Rate1 { get; set; }

        //public int LowLev2 { get; set; }
        //public int HighLev2 { get; set; }
        //public int Rate2 { get; set; }
        //public int LowLev3 { get; set; }
        //public int HighLev3 { get; set; }
        //public int Rate3 { get; set; }

    }
}
