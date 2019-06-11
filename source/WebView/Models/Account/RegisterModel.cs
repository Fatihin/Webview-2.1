﻿using System;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web.Mvc;

namespace WebView.Models.Account
{
    public class RegisterModel
    {
        [Required]
        [Display(Name = "User name")]
        public string USERNAME { get; set; }

        [Required]
        [StringLength(100, ErrorMessage = "The {0} must be at least {2} characters long.", MinimumLength = 6)]
        [DataType(DataType.Password)]
        [Display(Name = "Password")]
        public string PASSWORD { get; set; }

        [Required]
        [DataType(DataType.Password)]
        [Display(Name = "Confirm password")]
        [CompareAttribute("PASSWORD", ErrorMessage = "The password and confirmation password do not match.")]
        public string CONFIRMPASSWORD { get; set; }

        [Required]
        [Display(Name = "Full name")]
        public string FULL_NAME { get; set; }

        [Required]
        [Display(Name = "Group ID")]
        public string GROUPID { get; set; }

        [Required]
        [Display(Name = "Region")]
        public string REGIONID { get; set; }

        [Required]
        [Display(Name = "Nationwide")]
        public string NATION { get; set; }

        [Display(Name = "Area")]
        public string AREA { get; set; }
        
        [Display(Name = "RW Access")]
        public string RW_ACCESS { get; set; }
        
        [Display(Name = "Autorization")]
        public string AUTORIZATION { get; set; }
        
        [Display(Name = "Network")]
        public string NETWORK { get; set; }
        
        [Required]
        [Display(Name = "Ptt ID")]
        public string PTT_STATE { get; set; }

        [Required]
        [Display(Name = "Handover")]
        public string HANDOVER { get; set; }
        
        [Display(Name = "Roles")]
        public string USER_ROLES { get; set; }
        
        [Display(Name = "Exchange")]
        public string EXC { get; set; }
        
        [Required]
        [DataType(DataType.EmailAddress)]
        [Display(Name = "Email address")]
        public string Email { get; set; }

        [Display(Name = "No Tel")]
        public string NO_TEL { get; set; }
    }
}