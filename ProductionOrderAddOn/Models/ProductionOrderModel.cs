using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProductionOrderAddOn.Models
{
    class ProductionOrderModel
    {
        public string ProdNo { get; set; }
        public string ProdDesc { get; set; }
        public DateTime OrderDate { get; set; }
        public double Qty { get; set; }
        public ProductionType ProdType { get; set; }
        public string RefProdEntry { get; set; }
        public string RefProdNum { get; set; }
        public string RefBomVer { get; set; }
    }

    public enum ProductionType
    {
        FG,WIP
    }
}
