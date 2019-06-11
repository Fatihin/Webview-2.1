using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.ComponentModel.DataAnnotations;

namespace WebView.Models
{
    public class MaterialMastMaintenanceModel
    {
        public string MatId { get; set; }

        [Required]
        [StringLength(60)]
        [Display(Name = "Material Name")]
        public string MatName { get; set; }

        [StringLength(6)]
        [Display(Name = "Material UOM")]
        public string MatUOM { get; set; }

        [Display(Name = "Material Price")]
        public int MatPrice { get; set; }

        [StringLength(2)]
        [Display(Name = "Material Category")]
        public string MatCat { get; set; }

        [StringLength(1)]
        [Display(Name = "Record Type")]
        public string RecordType { get; set; }

        [StringLength(1)]
        [Display(Name = "Trans Type")]
        public string TransType { get; set; }
    }
}
