using SAPbobsCOM;
using System.Collections.Generic;

namespace ProductionOrderAddOn.Helpers
{
    public static class SapQueryHelper
    {
        public static List<Dictionary<string, object>> ExecuteQuery(string sql, Company company)
        {
            var results = new List<Dictionary<string, object>>();

            Recordset rs = null;

            try
            {
                rs = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rs.DoQuery(sql);

                while (!rs.EoF)
                {
                    var row = new Dictionary<string, object>();

                    for (int i = 0; i < rs.Fields.Count; i++)
                    {
                        var field = rs.Fields.Item(i);
                        row[field.Name] = field.Value;
                    }

                    results.Add(row);
                    rs.MoveNext();
                }
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }

            return results;
        }

        public static T ExecuteScalar<T>(string sql, Company company)
        {
            Recordset rs = null;

            try
            {
                rs = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rs.DoQuery(sql);

                if (!rs.EoF)
                {
                    return (T)rs.Fields.Item(0).Value;
                }

                return default;
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }
    }

}
