using ProductionOrderAddOn.Helpers;
using ProductionOrderAddOn.Models;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace ProductionOrderAddOn.Services
{
    internal static class ProductionOrderSapService
    {
        //private static readonly Company oCompany = CompanyService.GetCompany();

        public static List<int> CreateProductionOrdersRecursive(Company oCompany,string fileName, IEnumerable<ProductionOrderModel> models, HashSet<ProductionKey> visitedKeys = null)
        {
            if (models == null) throw new ArgumentNullException(nameof(models));
            if (visitedKeys == null) visitedKeys = new HashSet<ProductionKey>();

            var allDocEntries = new List<int>();

            // 🔍 Filter hanya model dengan kombinasi ProdNo + OrderDate yang belum diproses
            var batchModels = models
                .Where(m => visitedKeys.Add(new ProductionKey(m.ProdNo, m.OrderDate, m.RefProdEntry)))
                .ToList();

            if (batchModels.Count == 0) return allDocEntries;
            
            //bool startedTran = false;

            try
            {
                //if (!oCompany.InTransaction)
                //{
                //    oCompany.StartTransaction();
                //    startedTran = true;
                //}

                foreach (var m in batchModels)
                {
                    ProductionOrders po = null;
                    try
                    {
                        po = (ProductionOrders)oCompany.GetBusinessObject(BoObjectTypes.oProductionOrders);
                        po.ItemNo = m.ProdNo;
                        po.ProductionOrderType = BoProductionOrderTypeEnum.bopotStandard;
                        po.ProductionOrderStatus = BoProductionOrderStatusEnum.boposPlanned;
                        po.PlannedQuantity = m.Qty;
                        po.PostingDate = m.OrderDate.Date;
                        po.StartDate = m.OrderDate.Date;
                        po.DueDate = m.OrderDate.Date;
                        po.UserFields.Fields.Item("U_T2_PRODTYPE").Value = m.ProdType.ToString();
                        if (!string.IsNullOrEmpty(fileName))
                        {
                            po.Remarks = $"Imported from file {fileName}";
                            po.UserFields.Fields.Item("U_T2_Is_Import").Value = "Y";
                        }

                        if (!string.IsNullOrEmpty(m.RefProdEntry))
                            po.UserFields.Fields.Item("U_T2_Ref_Production").Value = m.RefProdEntry;
                        if (!string.IsNullOrEmpty(m.RefProdEntry))
                            po.UserFields.Fields.Item("U_T2_Ref_Prod_DocNum").Value = m.RefProdNum;
                        if (!string.IsNullOrEmpty(m.RefProdEntry))
                            po.UserFields.Fields.Item("U_BOMVER").Value = m.RefBomVer;

                        int rc = po.Add();
                        if (rc != 0)
                        {
                            oCompany.GetLastError(out int errCode, out string errMsg);
                            throw new Exception($"Failed to create {m.ProdNo} ({errCode}): {errMsg}");
                        }

                        int docEntry = int.Parse(oCompany.GetNewObjectKey());
                        allDocEntries.Add(docEntry);

                        UpdatePoStatus(oCompany,docEntry, BoProductionOrderStatusEnum.boposReleased);
                    }
                    finally
                    {
                        if (po != null) Marshal.ReleaseComObject(po);
                    }
                }

                //if (startedTran) oCompany.EndTransaction(BoWfTransOpt.wf_Commit);

                var wipModels = GetProductionOrders(oCompany,allDocEntries).ToList();

                if (wipModels.Count > 0)
                {
                    var subDocEntries = CreateProductionOrdersRecursive(oCompany,fileName,wipModels, visitedKeys);
                    allDocEntries.AddRange(subDocEntries);
                }

                return allDocEntries;
            }
            catch
            {
                //if (startedTran && oCompany.InTransaction)
                //    oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                throw;
            }
        }


        //Create ProdOrder Dari FORM1
        public static int CreateProductionOrder(
        Company oCompany,
        string fileName,
        ProductionOrderModel model)
        {
            if (model == null) throw new ArgumentNullException(nameof(model));

            ProductionOrders po = null;

            try
            {
                po = (ProductionOrders)oCompany.GetBusinessObject(BoObjectTypes.oProductionOrders);
                po.ItemNo = model.ProdNo;
                po.ProductionOrderType = BoProductionOrderTypeEnum.bopotStandard;
                po.ProductionOrderStatus = BoProductionOrderStatusEnum.boposPlanned;
                po.PlannedQuantity = model.Qty;
                po.PostingDate = model.OrderDate.Date;
                po.StartDate = model.OrderDate.Date;
                po.DueDate = model.OrderDate.Date;
                po.UserFields.Fields.Item("U_T2_PRODTYPE").Value = model.ProdType.ToString();

                if (!string.IsNullOrEmpty(fileName))
                {
                    po.Remarks = $"Imported from file {fileName}";
                    po.UserFields.Fields.Item("U_T2_Is_Import").Value = "Y";
                }

                if (!string.IsNullOrEmpty(model.RefProdEntry))
                {
                    po.UserFields.Fields.Item("U_T2_Ref_Production").Value = model.RefProdEntry;
                    po.UserFields.Fields.Item("U_T2_Ref_Prod_DocNum").Value = model.RefProdNum;
                }

                int rc = po.Add();
                if (rc != 0)
                {
                    oCompany.GetLastError(out int errCode, out string errMsg);
                    throw new Exception($"Failed to create PO {model.ProdNo} ({errCode}): {errMsg}");
                }

                string keyStr = oCompany.GetNewObjectKey();
                if (!int.TryParse(keyStr, out int docEntry))
                    throw new Exception("Failed to retrieve new production order DocEntry.");

                UpdatePoStatus(oCompany, docEntry, BoProductionOrderStatusEnum.boposReleased);

                return docEntry;
            }
            finally
            {
                if (po != null) Marshal.ReleaseComObject(po);
            }
        }


        public static void UpdatePoStatus(Company oCompany,int docEntry, BoProductionOrderStatusEnum target)
        {

            // Ambil PO
            var po = (ProductionOrders)oCompany.GetBusinessObject(
                         BoObjectTypes.oProductionOrders);

            if (!po.GetByKey(docEntry))
                throw new InvalidOperationException($"Production Order DocEntry {docEntry} not found.");

            // Jika sudah di status target, abaikan
            if (po.ProductionOrderStatus == target) return;

            po.ProductionOrderStatus = target;

            if (po.Update() != 0)
            {
                oCompany.GetLastError(out int errCode, out string errMsg);
                throw new InvalidOperationException(
                    $"Failed to update status Production Order ({errCode}): {errMsg}");
            }
        }
        
        public static List<ProductionOrderModel> GetProductionOrders(Company oCompany,IEnumerable<int> docEntries)
        {
            if (docEntries == null)
                throw new ArgumentNullException(nameof(docEntries));

            var ids = docEntries.Distinct().ToArray();
            if (ids.Length == 0)
                return new List<ProductionOrderModel>();

            string inClause = string.Join(", ", docEntries);

            string sql = $@"
                        SELECT
                            t0.DocEntry      AS RefProdEntry,
                            t0.DocNum      AS RefProdNum,
                            t2.Code        AS ProdNo,
                            t2.ItemName    AS ProdDesc,
                            t3.PlannedQty  AS Qty,
                            CAST(t0.PostDate AS DATE) AS OrderDate,
                            t0.U_BOMVER as RefBomVer
                        FROM OWOR  t0
                        INNER JOIN OITT t1 ON t0.ItemCode = t1.Code
                        INNER JOIN ITT1 t2 ON t1.Code     = t2.Father
                        INNER JOIN WOR1 t3 ON t3.DocEntry = t0.DocEntry
                                            AND t3.ItemCode = t2.Code
                        WHERE t0.DocEntry IN ({inClause})
                            AND ISNULL(t2.U_T2_ITEM_GROUP, '') = '1'
                        ORDER BY t0.PostDate DESC, t2.Code;";
            
            string sqlNonWIP = $@"
                        SELECT
                            t0.DocEntry      AS RefProdEntry,
                            t0.DocNum      AS RefProdNum,
                            t2.Code        AS ProdNo,
                            t2.ItemName    AS ProdDesc,
                            t3.PlannedQty  AS Qty,
                            CAST(t0.PostDate AS DATE) AS OrderDate,
                            t0.U_BOMVER as RefBomVer
                        FROM OWOR  t0
                        INNER JOIN OITT t1 ON t0.ItemCode = t1.Code
                        INNER JOIN ITT1 t2 ON t1.Code     = t2.Father
                        INNER JOIN WOR1 t3 ON t3.DocEntry = t0.DocEntry
                                            AND t3.ItemCode = t2.Code
                        INNER JOIN OITM t4 ON t4.ItemCode = t2.Code
                        WHERE t0.DocEntry IN ({inClause})
                            AND ISNULL(t2.U_T2_ITEM_GROUP, '') <> '1'
                            AND t4.TreeType = 'P'
                        ORDER BY t0.PostDate DESC, t2.Code;";

            try
            {
                var result = new List<ProductionOrderModel>();
                var data = SapQueryHelper.ExecuteQuery(sql, oCompany);
                var dataNonWIP = SapQueryHelper.ExecuteQuery(sqlNonWIP, oCompany);

                foreach (var row in data)
                {
                    var newOrder = new ProductionOrderModel
                    {
                        RefProdEntry = row["RefProdEntry"].ToString(),
                        RefProdNum = row["RefProdNum"].ToString(),
                        ProdNo = row["ProdNo"].ToString(),
                        ProdDesc = row["ProdDesc"].ToString(),
                        Qty = Convert.ToDouble(row["Qty"]),
                        OrderDate = Convert.ToDateTime(row["OrderDate"]),
                        ProdType = ProductionType.WIP,
                        RefBomVer = row["RefBomVer"].ToString(),
                    };
                    result.Add(newOrder);
                }

                foreach (var item in dataNonWIP)
                {
                    int tempDocEntry = int.Parse(item["RefProdEntry"].ToString());
                    string tempItem = item["ProdNo"].ToString();
                    double tempQty = Convert.ToDouble(item["Qty"]);
                    var resNonWip = GetRecursiveProductionOrders(oCompany, tempDocEntry, tempItem, tempQty);
                    if (resNonWip != null && resNonWip.Any())
                    {
                        result.AddRange(resNonWip);
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                throw new Exception("Error while retrieving Sub production orders: " + ex.Message, ex);
            }
        }

        public static List<ProductionOrderModel> GetRecursiveProductionOrders(Company oCompany, int docEntry, string itemCode, double qty)
        {
            string sqlNonWip = $@"
                            SELECT 
                                T1.VisOrder,
                                T1.Code AS ItemCode,
                                CAST(((T1.Quantity * {qty}) / T0.Qauntity) AS DECIMAL(19,6)) AS PlannedQty
                            FROM OITT T0
                            INNER JOIN ITT1 T1 ON T0.Code = T1.Father
                            INNER JOIN OITM T2 ON T2.ItemCode = T1.Code
                            WHERE T0.Code = '{itemCode}'
                              AND T2.TreeType = 'P'
                              AND ISNULL(T1.U_T2_ITEM_GROUP,'') <> '1' ";

            string sqlWip = $@"
                            SELECT 
                                T3.DocEntry AS RefProdEntry,
                                T3.DocNum AS RefProdNum,
                                T1.Code AS ProdNo,
                                T2.ItemName AS ProdDesc,
                                CAST(((T1.Quantity * {qty})/T0.Qauntity) AS DECIMAL(19,6))  AS Qty,
                                CAST(T3.PostDate AS DATE) AS OrderDate,
                                T3.U_BOMVER as RefBomVer
                            FROM OITT T0
                            INNER JOIN ITT1 T1 ON T0.Code = T1.Father
                            INNER JOIN OITM T2 ON T2.ItemCode = T1.Code
                            INNER JOIN OWOR T3 ON T3.DocEntry = {docEntry}
                            WHERE T0.Code = '{itemCode}'
                              AND T2.TreeType = 'P'
                              AND ISNULL(T1.U_T2_ITEM_GROUP,'') = '1' ";
            try
            {
                var result = new List<ProductionOrderModel>();
                var dataNonWip = SapQueryHelper.ExecuteQuery(sqlNonWip, oCompany);
                var dataWip = SapQueryHelper.ExecuteQuery(sqlWip, oCompany);

                foreach (var row in dataNonWip)
                {
                    string code = row["ItemCode"].ToString();
                    double pQty = double.Parse(row["PlannedQty"].ToString());
                    var res = GetRecursiveProductionOrders(oCompany, docEntry, code, pQty);
                    if (res != null && res.Any())
                    {
                        result.AddRange(res);
                    }
                }

                foreach (var row in dataWip)
                {
                    result.Add(new ProductionOrderModel
                    {
                        RefProdEntry = row["RefProdEntry"].ToString(),
                        RefProdNum = row["RefProdNum"].ToString(),
                        ProdNo = row["ProdNo"].ToString(),
                        ProdDesc = row["ProdDesc"].ToString(),
                        Qty = Convert.ToDouble(row["Qty"]),
                        OrderDate = Convert.ToDateTime(row["OrderDate"]),
                        ProdType = ProductionType.WIP,
                        RefBomVer = row["RefBomVer"].ToString(),
                    });
                }

                return result;
            }
            catch (Exception ex)
            {
                throw new Exception("Error while retrieving Sub production orders: " + ex.Message, ex);
            }
        }

        public static bool IsProdOrderExists(Company oCompany,ProductionOrderModel model)
        {
            if (model == null)
                throw new ArgumentNullException(nameof(model));

            // Sanitize ProdNo dan format tanggal dengan benar
            string prodNo = model.ProdNo.Replace("'", "''"); // untuk hindari SQL injection
            string dateStr = model.OrderDate.ToString("yyyy-MM-dd");

            string sql = $@"
                SELECT TOP 1 T0.DocEntry
                FROM OWOR T0
                WHERE T0.ItemCode = '{prodNo}'
                AND CAST(T0.PostDate AS DATE) = '{dateStr}'
                AND ISNULL(Status,'') <> 'C'   ";

            try
            {
                var result = SapQueryHelper.ExecuteQuery(sql, oCompany);
                return result.Count > 0;
            }
            catch (Exception ex)
            {
                throw new Exception("Error while validating production orders: " + ex.Message, ex);
            }
        }


        //Create Sub Dari FORM1
        public static Dictionary<int,string> GenerateSubOrder(Company oCompany,int docEntry, string fileName = "")
        {
            Dictionary<int,string> resultEntries = new Dictionary<int,string>();
            List<int> listEntries = new List<int>();

            try
            {
                var docEntries = new List<int> { docEntry };
                var wipModels = GetProductionOrders(oCompany,docEntries)?.ToList();

                if (wipModels != null && wipModels.Any())
                {
                    var subDocs = CreateProductionOrdersRecursive(oCompany,fileName, wipModels);
                    if (subDocs != null && subDocs.Any())
                    {
                        listEntries.AddRange(subDocs);
                    }
                }

                if (listEntries.Any())
                {
                    resultEntries = GetDocNum(oCompany,listEntries);
                }
                return resultEntries;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error in generate sub production order method : {ex.Message}", ex);
            }
        }

        public static List<int> GetSubCanceled(Company oCompany, int docEntry)
        {
            if (docEntry == 0)
                throw new ArgumentNullException(nameof(docEntry));

            SAPbobsCOM.ProductionOrders oProd = null;
            var result = new List<int>();
            try
            {
                oProd = (ProductionOrders)oCompany.GetBusinessObject(BoObjectTypes.oProductionOrders);
                var rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                if (!oProd.GetByKey(docEntry))
                    throw new InvalidOperationException($"Production Order DocEntry {docEntry} not found.");

                var refEntries = new List<int>();
                for (int i = 0; i < oProd.DocumentReferences.Count; i++)
                {
                    oProd.DocumentReferences.SetCurrentLine(i);
                    int refEntry = int.Parse(oProd.DocumentReferences.ReferencedDocEntry.ToString());
                    refEntries.Add(refEntry);
                }

                if (!refEntries.Any()) return result;

                var strRefEntries = string.Join(",", refEntries);
                string sql = $"SELECT DocNum AS ProdOrderNum FROM OWOR WHERE Status = 'C' AND DocEntry IN ({strRefEntries}) ORDER BY DocNum DESC ";

                rs.DoQuery(sql);
                while (!rs.EoF)
                {
                    result.Add((int)rs.Fields.Item("ProdOrderNum").Value);
                    rs.MoveNext();
                }

                return result;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static List<int> GetSubClosed(Company oCompany, int docEntry)
        {
            if (docEntry == 0)
                throw new ArgumentNullException(nameof(docEntry));

            SAPbobsCOM.ProductionOrders oProd = null;
            var result = new List<int>();
            try
            {
                oProd = (ProductionOrders)oCompany.GetBusinessObject(BoObjectTypes.oProductionOrders);
                var rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                if (!oProd.GetByKey(docEntry))
                    throw new InvalidOperationException($"Production Order DocEntry {docEntry} not found.");

                var refEntries = new List<int>();
                for (int i = 0; i < oProd.DocumentReferences.Count; i++)
                {
                    oProd.DocumentReferences.SetCurrentLine(i);
                    int refEntry = int.Parse(oProd.DocumentReferences.ReferencedDocEntry.ToString());
                    refEntries.Add(refEntry);
                }

                if (!refEntries.Any()) return result;

                var strRefEntries = string.Join(",", refEntries);
                string sql = $"SELECT DocNum AS ProdOrderNum FROM OWOR WHERE Status = 'L' AND DocEntry IN ({strRefEntries}) ORDER BY DocNum DESC ";

                rs.DoQuery(sql);
                while (!rs.EoF)
                {
                    result.Add((int)rs.Fields.Item("ProdOrderNum").Value);
                    rs.MoveNext();
                }

                return result;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static List<int> GetSubReceipted(Company oCompany, int docEntry)
        {
            if (docEntry == 0)
                throw new ArgumentNullException(nameof(docEntry));

            var result = new List<int>();

            var oProd = (ProductionOrders)oCompany.GetBusinessObject(BoObjectTypes.oProductionOrders);
            if (!oProd.GetByKey(docEntry))
                throw new InvalidOperationException($"Production Order DocEntry {docEntry} not found.");

            // Collect reference doc entries
            var refEntries = new List<int>();
            for (int i = 0; i < oProd.DocumentReferences.Count; i++)
            {
                oProd.DocumentReferences.SetCurrentLine(i);
                if (int.TryParse(oProd.DocumentReferences.ReferencedDocEntry.ToString(), out int refEntry))
                    refEntries.Add(refEntry);
            }

            if (!refEntries.Any())
                return result;

            var strRefEntries = string.Join(",", refEntries);

            string sql = $@"
        SELECT T0.DocNum AS ProdOrderNum
        FROM OWOR T0
        INNER JOIN IGN1 T1 ON T1.BaseType = 202 AND T1.BaseEntry = T0.DocEntry
        INNER JOIN OIGN T2 ON T2.DocEntry = T1.DocEntry
        WHERE T0.DocEntry IN ({strRefEntries})
        ORDER BY T0.DocNum DESC
    ";

            var rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rs.DoQuery(sql);

            while (!rs.EoF)
            {
                result.Add(Convert.ToInt32(rs.Fields.Item("ProdOrderNum").Value));
                rs.MoveNext();
            }

            return result;
        }

        public static bool UpdateSubOrder(Company oCompany, int docEntry)
        {
            SAPbobsCOM.ProductionOrders oProd = null;
            if (docEntry == 0)
                throw new ArgumentNullException(nameof(docEntry));
            
            try
            {
                oProd = (ProductionOrders)oCompany.GetBusinessObject(BoObjectTypes.oProductionOrders);
                var rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                if (!oProd.GetByKey(docEntry))
                    throw new InvalidOperationException($"Production Order DocEntry {docEntry} not found.");

                var refEntries = new List<int>();
                for (int i = 0; i < oProd.DocumentReferences.Count; i++)
                {
                    oProd.DocumentReferences.SetCurrentLine(i);
                    int refEntry = int.Parse(oProd.DocumentReferences.ReferencedDocEntry.ToString());
                    refEntries.Add(refEntry);
                    double newPlanQty = 0;

                    string sql = $"SELECT T1.PlannedQty FROM OWOR T0 JOIN WOR1 T1 ON T1.DocEntry = T0.U_T2_Ref_Production AND T1.ItemCode = T0.ItemCode WHERE T0.DocEntry = {oProd.DocumentReferences.ReferencedDocEntry} ";

                    rs.DoQuery(sql);
                    if (!rs.EoF)
                    {
                        newPlanQty = (double)rs.Fields.Item("PlannedQty").Value;
                        UpdatePoQty(oCompany,refEntry,newPlanQty);
                    }

                }

                //if (!refEntries.Any()) return false;

                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (oProd != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oProd);
            }
        }

        public static void UpdatePoQty(Company oCompany, int docEntry, double qty)
        {
            SAPbobsCOM.ProductionOrders po = null;
            try
            {
                // Ambil PO
                po = (ProductionOrders)oCompany.GetBusinessObject(
                             BoObjectTypes.oProductionOrders);

                if (!po.GetByKey(docEntry))
                    throw new InvalidOperationException($"Production Order DocEntry {docEntry} not found.");

                po.PlannedQuantity = qty;

                if (po.Update() != 0)
                {
                    oCompany.GetLastError(out int errCode, out string errMsg);
                    throw new InvalidOperationException(
                        $"Failed to update planned qty Production Order ({errCode}): {errMsg}");
                }
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                if (po != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(po);
            }
        }

        public static Dictionary<int,string> GetDocNum(Company oCompany,List<int> docEntries)
        {
            Dictionary<int,string> result = new Dictionary<int,string>();
            string inClause = string.Join(", ", docEntries);

            string sql = $@"
                    SELECT 
                        W.DocEntry AS DocEntry,
                        W.DocNum AS DocNum
                    FROM 
                        OWOR W
                        INNER JOIN NNM1 N ON W.Series = N.Series
                    WHERE 
                        W.DocEntry IN ({inClause})
                    ORDER BY 
                        W.DocNum
                    ";
            try
            {
                var data = SapQueryHelper.ExecuteQuery(sql, oCompany);

                foreach (var row in data)
                {
                    var entry = int.Parse(row["DocEntry"].ToString());
                    var num = row["DocNum"].ToString();
                    result.Add(entry,num);
                }

                return result;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static void UpdateRemarks(Company oCompany,int docEntry, string remarks)
        {
            SAPbobsCOM.ProductionOrders po = null;
            try
            {
                // Ambil PO
                po = (ProductionOrders)oCompany.GetBusinessObject(
                             BoObjectTypes.oProductionOrders);

                if (!po.GetByKey(docEntry))
                    throw new InvalidOperationException($"Production Order DocEntry {docEntry} not found.");
                po.Remarks = (string.IsNullOrEmpty(po.Remarks)) ? remarks : $"{po.Remarks}{Environment.NewLine}{remarks}";

               

                if (po.Update() != 0)
                {
                    oCompany.GetLastError(out int errCode, out string errMsg);
                    throw new InvalidOperationException(
                        $"Failed to update remarks Production Order ({errCode}): {errMsg}");
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (po != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(po);
            }
        }

        public static bool LinkWipToFG(Company oCompany, int fgDocEntry, int wipDocEntry)
        {
            SAPbobsCOM.ProductionOrders oWO = null;
            string sql = "SELECT T2.RefDocEntr AS TransId FROM OWOR T0 INNER JOIN WOR5 T2 ON T2.DocEntry = T0.DocEntry WHERE T0.DocEntry = '" + fgDocEntry + "'";
            try
            {
                Recordset rs = null;
                int recCount = 0;
                rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery(sql);

                if (!rs.EoF)
                {
                    recCount = rs.RecordCount;
                }
                
                oWO = (SAPbobsCOM.ProductionOrders)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders);

                if (!oWO.GetByKey(fgDocEntry))
                    throw new Exception($"Goods Receipt PO with DocEntry {fgDocEntry} not found.");

                // Add Goods Receipt as referenced document
                oWO.DocumentReferences.Add();
                oWO.DocumentReferences.SetCurrentLine(recCount);
                oWO.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_ProductionOrder;
                oWO.DocumentReferences.ReferencedDocEntry = wipDocEntry;
                oWO.DocumentReferences.IssueDate = DateTime.Today;
                oWO.DocumentReferences.Remark = "Auto-linked WIP Production Order";

                // Update GRPO
                int updateResult = oWO.Update();

                if (updateResult != 0)
                {
                    oCompany.GetLastError(out int errCode, out string errMsg);
                    throw new Exception($"Failed to link Sub Production Order to FG Production Order. Error ({errCode}): {errMsg}");
                }
                return true;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (oWO != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oWO);
            }
        }

        public static void CancelSubOrder(Company oCompany, int docEntry)
        {
            SAPbobsCOM.ProductionOrders po = null;
            try
            {
                po = (ProductionOrders)oCompany.GetBusinessObject(
                         BoObjectTypes.oProductionOrders);

                if (!po.GetByKey(docEntry))
                    throw new InvalidOperationException($"Production Order DocEntry {docEntry} not found.");

                for (int i = 0; i < po.DocumentReferences.Count; i++)
                {
                    po.DocumentReferences.SetCurrentLine(i);
                    if (po.DocumentReferences.ReferencedObjectType == ReferencedObjectTypeEnum.rot_ProductionOrder)
                    {
                        CancelProductionOrder(oCompany,po.DocumentReferences.ReferencedDocEntry);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (po != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(po);
            }
        }

        public static bool CancelProductionOrder(SAPbobsCOM.Company oCompany, int docEntry)
        {
            SAPbobsCOM.ProductionOrders oProd = null;
            try
            {
                oProd = (SAPbobsCOM.ProductionOrders)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders);

                if (!oProd.GetByKey(docEntry))
                    throw new Exception($"Production Order {docEntry} not found.");

                // Cancel the order
                int ret = oProd.Cancel();
                if (ret != 0)
                {
                    oCompany.GetLastError(out int errCode, out string errMsg);
                    throw new Exception($"Failed to cancel Production Order {docEntry}. Error ({errCode}): {errMsg}");
                }

                return true;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (oProd != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oProd);
            }
        }

    }
}
