using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProductionOrderAddOn.Models
{
    public class ProductionKey
    {
        public string ProdNo { get; }
        public DateTime OrderDate { get; }
        public string RefProd { get; }

        public ProductionKey(string prodNo, DateTime orderDate, string refProd)
        {
            ProdNo = prodNo;
            OrderDate = orderDate.Date; // pakai hanya tanggal, buang jam
            RefProd = refProd; // pakai hanya tanggal, buang jam
        }

        public override bool Equals(object obj)
        {
            return obj is ProductionKey other &&
                   ProdNo == other.ProdNo &&
                   OrderDate == other.OrderDate &&
                   RefProd == other.RefProd
                   ;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(ProdNo, OrderDate);
        }
    }

}
