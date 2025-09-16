using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using WebStoreHouse.Models;

namespace WebStoreHouse.ViewModels
{
    public class MaintainViewModels
    {
        public int fId { get; set; }
        public string fUserId { get; set; }
        public string fName { get; set; }
        public string ROLE_DESC { get; set; }
        public string ROLE_ID { get; set; }
    }
}