using System;
using System.ComponentModel.DataAnnotations;

namespace WebView.Models
{
    public class DllVersion_NewJobModel
    {
        [Display(Name = "Title")]
        public string DLL_NAME { get; set; }

        [StringLength(250)]
        [Display(Name = "Description")]
        public string DLL_DESCRIPTION { get; set; }

        [Display(Name = "Title")]
        public string DLL_MASTER { get; set; }

        [StringLength(250)]
        [Display(Name = "Description")]
        public string DLL_VERSION { get; set; }


        DateTime userStartDate = DateTime.Now;//set Default Value here
        [DataType(DataType.Date)]
        public DateTime CREATE_DATE
        {
            get
            {
                return userStartDate;
            }
            set
            {
                userStartDate = value;
            }
        }
      
    }
}