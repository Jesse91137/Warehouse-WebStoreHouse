using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using WebStoreHouse.Models;

namespace WebStoreHouse.ViewModels
{
    public class InvoicingEditModel
    {
        public string Sno { get; set; }
        public string Wono { get; set; }
        public string Wono_Cust { get; set; }
        public string AvalueDN { get; set; }
        public string Engsr { get; set; }
        public int Amount { get; set; }
        public string Position { get; set; }
        public DateTime Date { get; set; }
        public string Mark { get; set; }
        public string I_1 { get; set; }
        public int I_2 { get; set; }
        public string I_3 { get; set; }
        public int I_4 { get; set; }
    }
}