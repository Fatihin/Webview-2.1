using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.ComponentModel.DataAnnotations;

namespace WebView.Models
{
    public class ANDModel
    {
        [Required]
        [Display(Name = "Can Code")]
        public string CanCode { get; set; }

        [Required]
        [Display(Name = "Can Name")]
        public string CanName { get; set; }

        [Required]
        [Display(Name = "Region Id")]
        public string RegionId { get; set; }

        [Display(Name = "Target ECP")]
        public string TargetEcp { get; set; }

    }
}