using System;
using System.ComponentModel.DataAnnotations;

namespace WebView.Models
{
    public class JobModel
    {
        public int JobId { get; set; }

        [Required]
        [Display(Name = "Name")]
        public string Name { get; set; }

        [Display(Name = "Description")]
        public string Description { get; set; }

        [Display(Name = "Project Id")]
        public int ProjectId { get; set; }

        [Display(Name = "Scheme Id")]
        public int SchemeId { get; set; }

        [Display(Name = "ISP Design Id")]
        public int ISPdesign { get; set; }

        [Display(Name = "Job State")]
        public int JobState { get; set; }

        [Display(Name = "Job State")]
        public string JobStateNew { get; set; }

        [Display(Name = "Transition State")]
        public int TransitionState { get; set; }

        public bool FlagIsp { get; set; }
        public bool FlagOsp { get; set; }
    }
}