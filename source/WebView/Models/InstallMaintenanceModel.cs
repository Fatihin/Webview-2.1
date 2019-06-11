using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.ComponentModel.DataAnnotations;

namespace WebView.Models
{
    public class InstallMaintenanceModel
    {
        public int codeid { get; set; }
    
        [Required]
        [Display(Name = "Install Code")]
        public string InstallCode { get; set; }

        [Display(Name = "Install Name")]
        public string InstallName { get; set; }

        [Display(Name = "SUP Normal Rate")]
        public int SUPNormRate { get; set; }

        [Display(Name = "SUP RWO Rate")]
        public int SUPRWORate { get; set; }

        [Display(Name = "IMPL Normal Rate")]
        public int IMPLNormalRate{ get; set; }

        [Display(Name = "IMPL RWO Rate")]
        public int IMPLRWORate { get; set; }

        [Display(Name = "Install Type")]
        public string InstallType { get; set; }

        [Display(Name = "Rank")]
        public string Rank { get; set; }        

    }
}
