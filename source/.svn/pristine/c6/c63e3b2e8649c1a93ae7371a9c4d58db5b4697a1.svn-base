using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace WebView.Models
{
    public class UserSendModel
    {
        public class User
        {
            [Required]
            [Display(Name = "Id")]
            public decimal Id { get; set; }

            [Required]
            [Display(Name = "Name")]
            public string Name { get; set; }

            [Required]
            [Display(Name = "Group Id")]
            public string GroupId { get; set; }
        }
        public class ViewModel
        {
            public List<User> AvailableProducts { get; set; }
            public List<User> RequestedProducts { get; set; }
            public List<int> ListSend { get; set; }

            public int[] AvailableSelected { get; set; }
            public int[] RequestedSelected { get; set; }
            public string SavedRequested { get; set; }
        }
            

            //public List<SelectListItem> AvailableProducts { get; set; }
            //public List<SelectListItem> RequestedProducts { get; set; }

            //public int[] AvailableSelected { get; set; }
            //public int[] RequestedSelected { get; set; }
            //public string SavedRequested { get; set; }

            //public int[] SelectedItemIds { get; set; }
            //public MultiSelectList Items { get; set; }
    }


}