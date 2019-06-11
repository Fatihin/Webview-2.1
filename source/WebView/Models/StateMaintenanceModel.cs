using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.ComponentModel.DataAnnotations;

namespace WebView.Models
{
    public class StateMaintenanceModel 
    {
        [Required]
        [Display(Name = "State Id")]
        public string StateId { get; set; }

        [Required]
        [Display(Name = "State Name")]
        public string StateName { get; set; }

        [Required]
        [Display(Name = "Region Id")]
        public string RegionId { get; set; }

    }
}
