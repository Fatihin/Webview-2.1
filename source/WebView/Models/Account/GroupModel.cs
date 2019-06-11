using System;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace WebView.Models.Account
{
    public class GroupModel
    {
        [Required]
        [Display(Name = "Id")]
        public decimal Id { get; set; }
        
        [Required]
        [Display(Name = "Group Name")]
        public string GrpName { get; set; }

        [Required]
        [Display(Name = "Full Name")]
        public string FullName { get; set; }
    }
}