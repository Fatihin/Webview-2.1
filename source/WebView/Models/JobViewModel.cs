using System;

namespace WebView.Models
{
    public class JobViewModel
    {
        public decimal JobId { get; set; }
        public string JobName { get; set; } 
        public string Description { get; set; }
        public string ProjectName { get; set; }
        public string SchemeName { get; set; }
        public string IspName { get; set; }
        public string JobState { get; set; }
        public string TransitionState { get; set; }
    }
}