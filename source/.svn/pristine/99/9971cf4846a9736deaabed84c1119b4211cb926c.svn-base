using System;
using System.ComponentModel.DataAnnotations;

namespace WebView.Models
{
    public class NewJobModel
    {
        public int JobId { get; set; }

        [Display(Name = "Job Name")]
        public string Name { get; set; }

        [StringLength(250)]
        [Display(Name = "Description")]
        public string Description { get; set; }

        [Display(Name = "Job State")]
        public string JobState { get; set; }

        [Required]
        [Display(Name = "EXC_ABB")]
        public string EXC_ABB { get; set; }

        [Required]
        [UIHint("SchemeType")]
        [Display(Name = "Scheme Type")]
        public string SchemeType { get; set; }

        DateTime userStartDate = DateTime.Now;//set Default Value here
        DateTime userEndDate = DateTime.Now;//set Default Value here
        [DataType(DataType.Date)]
        public DateTime PlanStartDate
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

        [DataType(DataType.Date)]
        public DateTime PlanEndDate
        {
            get
            {
                return userEndDate;
            }
            set
            {
                userEndDate = value;
            }
        }
        
        public bool chkISP { get; set; }
        public bool chckOSP { get; set; }

        [StringLength(250)]
        [Display(Name = "Description 2")]
        public string Description_2 { get; set; }
    }
}