using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace LeadTracker.Models
{
    public class Data
    {
        [Display(Name = "Agent ID Name")]
        public string Agent_ID_Name { get; set; }
        [Display(Name = "Agent Name")]
        public string Agent_Name { get; set; }
        [Display(Name = "Client Name")]
        public string Client_Name { get; set; }
        [Display(Name = "Client Email")]
        public string Client_Email { get; set; }
        [Display(Name = "Proper Details")]
        public string Proper_Details { get; set; }
        [Display(Name = "Upfront")]
        public string Upfront { get; set; }
        [Display(Name = "Total Sale")]
        public string Total_Sales { get; set; }
        [Display(Name = "Remaining")]
        public string Remaining { get; set; }
        [Display(Name = "Product")]
        public string Product { get; set; }
    }
}