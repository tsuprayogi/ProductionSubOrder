using SAPbobsCOM;
using System.Runtime.InteropServices;

namespace ProductionOrderAddOn.Services
{
    public static class DatabaseInitializer
    {

        public static void Init()
        {
            Company company = CompanyService.GetCompany();

            CreateProductionOrderUDFs(company);
        }


        // =====================================================
        // ADD FIELD (SAFE)
        // =====================================================
        static void AddField(Company company, string table, string field,
                             string desc, BoFieldTypes type, int size = 0)
        {
            if (FieldExists(company, table, field))
                return;

            UserFieldsMD udf = null;

            try
            {
                udf = (UserFieldsMD)
                    company.GetBusinessObject(BoObjectTypes.oUserFields);

                udf.TableName = table;
                udf.Name = field;
                udf.Description = desc;
                udf.Type = type;

                if (type == BoFieldTypes.db_Alpha)
                    udf.Size = size;

                Check(company, udf.Add());
            }
            finally
            {
                Release(udf);
            }
        }


        //Create BOM Version on OWOR
        static void CreateProductionOrderUDFs(Company company)
        {
            AddFieldWithValidValues(
                    company,
                    "OWOR",
                    "T2_PRODTYPE",
                    "Production Type",
                    BoFieldTypes.db_Alpha,
                    10,
                    new[]
                    {
                        ("FG",  "Finished Goods"),
                        ("WIP", "Work In Progress")
                    });


            AddField(company, "OWOR", "T2_Ref_Production",
                    "Ref Production Number", BoFieldTypes.db_Alpha, 12);

            AddField(company, "OWOR", "T2_Ref_Prod_DocNum",
                    "Ref Production Document Number", BoFieldTypes.db_Alpha, 12);

            AddField(company, "OWOR", "T2_Is_Import",
                    "Imported File", BoFieldTypes.db_Alpha, 1);

            AddFieldWithValidValues(
                    company,
                    "ITT1",
                    "T2_ITEM_GROUP",
                    "Item Group for Production",
                    BoFieldTypes.db_Alpha,
                    10,
                    new[]
                    {
                        ("1", "WIP"),
                        ("2",  "NON")
                    });
        }

        // =====================================================
        // CHECK TABLE
        // =====================================================
        static bool TableExists(Company company, string tableName)
        {
            Recordset rs = null;

            try
            {
                rs = (Recordset)
                    company.GetBusinessObject(BoObjectTypes.BoRecordset);

                rs.DoQuery($@"
                    SELECT 1
                    FROM OUTB
                    WHERE TableName = '{tableName.Replace("@", "")}'
                ");

                return !rs.EoF;
            }
            finally
            {
                Release(rs);
            }
        }

        // =====================================================
        // CHECK FIELD
        // =====================================================
        static bool FieldExists(Company company, string table, string field)
        {
            Recordset rs = null;

            try
            {
                rs = (Recordset)
                    company.GetBusinessObject(BoObjectTypes.BoRecordset);

                rs.DoQuery($@"
                    SELECT 1
                    FROM CUFD
                    WHERE TableID = '{table.Replace("@", "")}'
                      AND AliasID = '{field.Replace("U_", "")}'
                ");

                return !rs.EoF;
            }
            finally
            {
                Release(rs);
            }
        }

        // =====================================================
        // ERROR HANDLER
        // =====================================================
        static void Check(Company company, int ret)
        {
            if (ret != 0)
            {
                company.GetLastError(out int errCode, out string errMsg);
                throw new System.Exception($"{errCode} - {errMsg}");
            }
        }


        static void AddFieldWithValidValues(
            Company company,
            string table,
            string field,
            string desc,
            BoFieldTypes type,
            int size,
            (string Value, string Desc)[] validValues,
            string defaultValue = null
)
        {
            if (FieldExists(company, table, field))
                return;

            UserFieldsMD udf = null;

            try
            {
                udf = (UserFieldsMD)
                    company.GetBusinessObject(BoObjectTypes.oUserFields);

                udf.TableName = table;
                udf.Name = field;
                udf.Description = desc;
                udf.Type = type;

                if (type == BoFieldTypes.db_Alpha)
                    udf.Size = size;

                // ===== Add Valid Values =====
                foreach (var vv in validValues)
                {
                    udf.ValidValues.Value = vv.Value;
                    udf.ValidValues.Description = vv.Desc;
                    udf.ValidValues.Add();
                }

                if (!string.IsNullOrEmpty(defaultValue))
                    udf.DefaultValue = defaultValue;

                Check(company, udf.Add());
            }
            finally
            {
                Release(udf);
            }
        }


        // =====================================================
        // COM RELEASE HELPER (KUNCI 1120)
        // =====================================================
        static void Release(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                Marshal.ReleaseComObject(obj);
            }
        }

    }
}
