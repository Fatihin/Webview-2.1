using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.ComponentModel.DataAnnotations;

namespace WebView.Models
{
    public class ExchangeMaintenanceModel
    {
        [Required]
        [StringLength(4, ErrorMessage = "The {0} must be at least {2} characters long.", MinimumLength = 3)]
        [Display(Name = "Exchange Abbr")]
        public string EXC_ABB { get; set; }

        [Required]
        [Display(Name = "Exchange Name")]
        public string EXC_NAME { get; set; }

        [Required]
        [Display(Name = "PTT Id")]
        public string PTT_ID { get; set; }

        //[Required]
        [Display(Name = "Loc No")]
        public string LOC_NO { get; set; }

        [Required]
        [Display(Name = "State Id")]
        public string STATE_ID { get; set; }

        //[Required]
        [Display(Name = "Main No")]
        public int MAIN_NO{ get; set; }

        [Required]
        [Display(Name = "Area Classification")]
        public string CAT_CODE { get; set; }

        [Required]
        [Display(Name = "EXC Type")]
        public string EXC_TYPE { get; set; }

        //[Required]
        [Display(Name = "EXC Id")]
        public string EXC_ID { get; set; }

        [Required]
        [Display(Name = "Segment")]
        public string SEGMENT { get; set; }

        [Display(Name = "Address 1")]
        public string ADDR1 { get; set; }

        [Display(Name = "Address 2")]
        public string ADDR2 { get; set; }

        [Display(Name = "City 11")]
        public string CITY { get; set; }
        
    }
}
