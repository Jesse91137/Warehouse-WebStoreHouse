using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebStoreHouse.Models
{
    public class ConsignmentModel
    {
        private E_StoreHouseEntities db;
        public string Code { get; set; }
        public List<E_Compyany> E_Compyany { get; set; }
        public ConsignmentModel(E_StoreHouseEntities dbContext,string code,string contact)
        {
            db = dbContext;
            Code = code;
            //Contact = contact;
        }
         

        public List<E_Compyany> CompList()
        {
            List<E_Compyany> comp = db.E_Compyany.Where(c => c.code == Code).ToList();

            return comp;
        }
    }
}