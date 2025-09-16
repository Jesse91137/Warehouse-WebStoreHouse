using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebStoreHouse.Models
{
    public class QC_Model
    {        
        private E_StoreHouseEntities db;

        public QC_Model(E_StoreHouseEntities dbContext)
        {
            db = dbContext;
        }
        public class QCData
        {
            public int ProCount { get; set; }
            public string Product { get; set; }
            public string IClass { get; set; }
            public List<QCDataList> QCDataList { get; set; }
        }
        public class QCDataList
        {
            public string wono { get; set; }
            public string wono_cust { get; set; }
            public string engsr { get; set; }
            public string order_cust { get; set; }            
        }
        public List<QCData> GetData()
        {
            List<QCData> QCData = new List<QCData>();
            //var query = from qc in db.E_QC
            //        group qc by qc.product_serialno.Substring(0, 10) into g
            //        select new QCData
            //        {
            //            ProCount = g.Count(),                        
            //            Product = g.Key,                        
            //        };

            var query = from qc in db.E_QC
                        group qc by new { qc.iclass, Product = qc.product_serialno.Substring(0, 10) } into g
                        select new QCData
                        {
                            ProCount = g.Count(),
                            Product = g.Key.Product,
                            IClass = g.Key.iclass
                        };
            foreach (var item in query)
            {
                var _Z19 = (from _19 in db.E_ZRSD19
                            where _19.wono_cust == item.Product
                            select new QCDataList
                            {
                                wono = _19.wono,
                                wono_cust = _19.wono_cust,
                                engsr = _19.item,
                                order_cust = _19.order_cust
                            }).ToList();

                item.QCDataList = _Z19;                
                QCData.Add(item);
            }
            return QCData;
        }
    }
}