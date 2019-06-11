using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace WebView.Models
{
    public class HandoverModel
    {
        public class user
        {
            [Required]
            [Display(Name = "Id")]
            public decimal Id { get; set; }

            [Required]
            [Display(Name = "Name")]
            public string Name { get; set; }

            [Required]
            [Display(Name = "GroupId")]
            public string GroupId { get; set; }
        }

        public class ViewModel
        {
            public List<user> AvailableUsers { get; set; }
            public List<user> AvailableUsers2 { get; set; }
            public List<user> RequestedUsers { get; set; }
            public List<int> ListSend { get; set; }

            public int[] AvailableSelected { get; set; }
            public int[] RequestedSelected { get; set; }
            public string SavedRequested { get; set; }
        }
    }
}