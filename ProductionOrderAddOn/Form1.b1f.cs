using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using ProductionOrderAddOn.Helpers;
using ProductionOrderAddOn.Models;
using ProductionOrderAddOn.Services;
using SAPbobsCOM;
using SAPbouiCOM.Framework;

namespace ProductionOrderAddOn
{
    [FormAttribute("ProductionOrderAddOn.ImportFile", "Form1.b1f")]
    class ImportForm : UserFormBase
    {
        private SAPbouiCOM.EditText TxtFrom;
        private SAPbouiCOM.EditText TxtTo;
        private SAPbouiCOM.EditText TxtPath;
        private SAPbouiCOM.StaticText LblDateFrom;
        private SAPbouiCOM.StaticText LblDateTo;
        private SAPbouiCOM.Button BtnImport;
        private List<ProductionOrderModel> listData;
        const string DT_NAME = "DT_IMPORT";
        SAPbouiCOM.DataTable dt;
        private SAPbouiCOM.Grid GridData;
        private SAPbouiCOM.StaticText LblPath;
        private SAPbouiCOM.Button BtnBrowse;
        private string FilePath;
        private SAPbouiCOM.Button BtnRefresh;
        private SAPbouiCOM.Button BtnReset;
        private SAPbouiCOM.Button BtnProdOrderList;

        //public ImportForm()
        //{
        //}

        public override void OnInitializeComponent()
        {
            this.LblPath = ((SAPbouiCOM.StaticText)(this.GetItem("LblPath").Specific));
            this.TxtPath = ((SAPbouiCOM.EditText)(this.GetItem("TxtPath").Specific));
            this.TxtFrom = ((SAPbouiCOM.EditText)(this.GetItem("TxtFrom").Specific));
            this.TxtTo = ((SAPbouiCOM.EditText)(this.GetItem("TxtTo").Specific));
            this.TxtPath.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.TxtPath_KeyDownAfter);
            this.LblDateFrom = ((SAPbouiCOM.StaticText)(this.GetItem("LblFrom").Specific));
            this.LblDateTo = ((SAPbouiCOM.StaticText)(this.GetItem("LblTo").Specific));
            this.BtnImport = ((SAPbouiCOM.Button)(this.GetItem("BtnImport").Specific));
            this.BtnImport.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.BtnImport_ClickBefore);
            this.GridData = ((SAPbouiCOM.Grid)(this.GetItem("GridData").Specific));
            this.BtnBrowse = ((SAPbouiCOM.Button)(this.GetItem("BtnBrowse").Specific));
            this.BtnBrowse.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.BtnBrowse_ClickBefore);
            this.TxtFrom.ValidateAfter += new SAPbouiCOM._IEditTextEvents_ValidateAfterEventHandler(this.DateValidateAfter);
            this.TxtTo.ValidateAfter += new SAPbouiCOM._IEditTextEvents_ValidateAfterEventHandler(this.DateValidateAfter);
            this.BtnRefresh = ((SAPbouiCOM.Button)(this.GetItem("BtnRefresh").Specific));
            this.BtnRefresh.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.BtnRefresh_ClickBefore);
            this.BtnReset = ((SAPbouiCOM.Button)(this.GetItem("BtnReset").Specific));
            this.BtnReset.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.BtnReset_ClickBefore);
            this.BtnProdOrderList = ((SAPbouiCOM.Button)(this.GetItem("BtnImpList").Specific));
            this.BtnProdOrderList.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.BtnProdOrderList_ClickBefore);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);

        }

        private void OnCustomInitialize()
        {
            
        }

        #region Events
        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            
        }

        private void TxtPath_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (!string.IsNullOrEmpty(TxtPath.Value))
            {
                FilePath = TxtPath.Value;
            }
        }

        private void BtnImport_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (!this.ImportFromExcelProdOrder(pVal.FormUID))
                {
                    ClearDataModel();
                }
                else
                {
                    this.SetDataGrid(pVal.FormUID);
                    if(this.ImportToSAP(pVal.FormUID))
                        this.Reset();
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            
        }

        private void BtnBrowse_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            HandleBrowse(pVal.FormUID);
        }

        private void DateValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            DateRangeValidation(pVal);
            //OnChangeDate(pVal);
        }

        private void BtnRefresh_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (!string.IsNullOrEmpty(FilePath))
                {
                    if (!this.ImportFromExcelProdOrder(pVal.FormUID))
                        ClearDataModel();
                    else
                        this.SetDataGrid(pVal.FormUID);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void BtnReset_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            Reset();
        }

        private void BtnProdOrderList_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            var poForm = new ProdOrderListForm();
            poForm.Show();

        }

        #endregion

        #region Functions
        private void HandleBrowse(string formUID)
        {
            try
            {
                if (this.GetPathFile())
                {
                    if (!this.ImportFromExcelProdOrder(formUID))
                    {
                        ClearDataModel();
                    }
                    else
                    {
                        this.SetDataGrid(formUID);

                        Application.SBO_Application.StatusBar.SetText(
                            "Data imported successfully.",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error import: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }
        private void DateRangeValidation(SAPbouiCOM.SBOItemEventArg pVal)
        {
            DateTime? from = ParseDate_yyyyMMdd(TxtFrom.Value);
            DateTime? to = ParseDate_yyyyMMdd(TxtTo.Value);

            // 5) Jika KEDUA‑DUANYA terisi dan From > To → error
            if (from.HasValue && to.HasValue && from.Value.Date > to.Value.Date)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Date From cannot be greater than Date To",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                // Kembalikan field yang barusan diubah supaya tetap valid
                if (pVal.ItemUID == "TxtFrom")
                    TxtFrom.Value = to.Value.ToString("yyyyMMdd");
                else                                            // TxtTo yang diubah
                    TxtTo.Value = from.Value.ToString("yyyyMMdd");

                return;
            }
        }

        private bool GetPathFile()
        {
            bool res = false;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
            try
            {
                FormHelper.StartLoading(oForm, "Get path file...", 0, false);

                Thread t = new Thread(() =>
                {
                    using (var dummyForm = new System.Windows.Forms.Form
                    {
                        TopMost = true,
                        ShowInTaskbar = false,
                        WindowState = System.Windows.Forms.FormWindowState.Minimized
                    })
                    using (var dialog = new System.Windows.Forms.OpenFileDialog
                    {
                        Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
                        Title = "Select Excel file"
                    })
                    {
                        dummyForm.Show();
                        dummyForm.Hide();

                        if (dialog.ShowDialog(dummyForm) == System.Windows.Forms.DialogResult.OK)
                        {
                            FilePath = dialog.FileName.Trim('"').Trim();
                            res = true;
                        }
                        else
                        {
                            res = false;
                        }
                    }
                });

                t.SetApartmentState(ApartmentState.STA);
                t.Start();
                t.Join();

                TxtPath.Value = FilePath ?? "";
                return res;
            }
            catch
            {
                throw;
            }
            finally
            {
                FormHelper.FinishLoading(oForm);
            }
        }

        private bool ImportFromExcelProdOrder(string formUID)
        {
            bool res = false;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(formUID);
            try
            {
                FormHelper.StartLoading(oForm, "Importing data ...", 0, false);

                string fromStr = TxtFrom.Value;
                string toStr = TxtTo.Value;

                DateTime? fromDate = ParseDate_yyyyMMdd(fromStr);

                DateTime? toDate = ParseDate_yyyyMMdd(toStr);

                if (string.IsNullOrEmpty(FilePath)) throw new Exception("Please select file to import");

                // Panggil import service
                this.listData = ExcelImportService.ImportProductionOrders(FilePath, fromDate, toDate);
                
                if (!this.listData.Any())
                {
                    ClearDataModel();
                    throw new Exception("Data not found");
                }
                res = true;
                return res;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                FormHelper.FinishLoading(oForm);
            }
        }

        private bool ImportToSAP(string formUID)
        {
            bool success = false;
            Company oCompany = null;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(formUID);
            try
            {
                oCompany = Services.CompanyService.GetCompany();
                oCompany.StartTransaction();
                if (string.IsNullOrEmpty(this.FilePath))
                    throw new Exception("Please select a file to import.");

                int result = Application.SBO_Application.MessageBox(
                    "Are you sure you want to Import to SAP?",
                    1, "Yes", "No", "");

                if (result != 1)
                {
                    if (oCompany.InTransaction)
                    {
                        oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                    }
                    return false;
                }
                FormHelper.StartLoading(oForm, "Importing data to SAP", 0, false);

                if (listData == null || listData.Count == 0)
                    if (!this.ImportFromExcelProdOrder(formUID))
                        ClearDataModel();


                if (listData == null || listData.Count == 0)
                    throw new Exception("No data found in the selected file.");

                foreach (var item in listData)
                {
                    if (ProductionOrderSapService.IsProdOrderExists(oCompany,item))
                    {
                        throw new Exception($"A Production Order for item '{item.ProdNo}' already exists on {item.OrderDate:dddd, dd MMMM yyyy}.");

                    }
                }

                string fileName = System.IO.Path.GetFileName(FilePath);
                
                List<int> fgDocEntries = new List<int>();

                foreach (var item in listData)
                {
                    if (item.Qty > 0)
                    {
                        var resultEntry = ProductionOrderSapService.CreateProductionOrder(oCompany, fileName, item);
                        var listDoc = ProductionOrderSapService.GenerateSubOrder(oCompany, resultEntry, fileName);
                        foreach (var dictionary in listDoc)
                        {
                            int wipEntry = dictionary.Key;
                            ProductionOrderSapService.LinkWipToFG(oCompany, resultEntry, wipEntry);
                        }

                        if (listDoc != null && listDoc.Any())
                        {
                            string remarks = "WIP Production Orders: " + string.Join(" | ", listDoc.Values);
                            ProductionOrderSapService.UpdateRemarks(oCompany, resultEntry, remarks);

                        }

                        fgDocEntries.Add(resultEntry);
                    }
                }

                if (fgDocEntries == null || fgDocEntries.Count == 0)
                    throw new Exception("No production orders were created in SAP.");

                if (oCompany.InTransaction)
                {
                    oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                }
                success = true;
                return success;
            }
            catch (Exception ex)
            {
                if (oCompany.InTransaction)
                {
                    oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
                throw new Exception("Error during import: " + ex.Message);

                // Optional: log ke file atau tampilkan pesan lebih detail jika diperlukan
            }
            finally
            {
                FormHelper.FinishLoading(oForm);
                oCompany = null;
                if (success)
                {
                    System.Threading.ThreadPool.QueueUserWorkItem(_ =>
                    {
                        System.Threading.Thread.Sleep(500); // Delay agar SAP selesai menampilkan pesan
                        Application.SBO_Application.StatusBar.SetText(
                        "All records have been successfully imported into SAP.",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    });
                }
            }
        }

        private void Reset()
        {
            ClearDataModel();
            this.TxtFrom.Value = string.Empty;
            this.TxtTo.Value = string.Empty;
            this.FilePath = string.Empty;
            this.TxtPath.Value = string.Empty;
        }

        private void ClearDataModel()
        {
            if (listData != null) this.listData.Clear();
            if (dt != null)
            {
                this.dt.Clear();
                this.GridData.DataTable.Clear();
            }
        }

        private void BuildOrResetDataTable(SAPbouiCOM.IForm oForm)
        {
            try
            {
                dt = oForm.DataSources.DataTables.Item(DT_NAME);   // ada? ambil
            }
            catch (System.Runtime.InteropServices.COMException)   // belum ada
            {
                dt = oForm.DataSources.DataTables.Add(DT_NAME);    // → buat
            }

            dt.Clear();

            dt.Columns.Add("No.", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
            dt.Columns.Add("Description", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
            dt.Columns.Add("Qty", SAPbouiCOM.BoFieldsType.ft_Quantity);
            dt.Columns.Add("Order Date", SAPbouiCOM.BoFieldsType.ft_Date);

            foreach (var x in listData)
            {
                int row = dt.Rows.Count;   // ambil indeks berikutnya
                dt.Rows.Add();             // tambahkan baris kosong
                dt.SetValue("No.", row, x.ProdNo);
                dt.SetValue("Description", row, x.ProdDesc);
                dt.SetValue("Order Date", row, x.OrderDate);
                dt.SetValue("Qty", row, x.Qty);
            }

            int white = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

            GridData.DataTable = dt;
            GridData.AutoResizeColumns();
            for (int i = 0; i < GridData.Columns.Count; i++)
            {
                var col = GridData.Columns.Item(i);

                // Untuk SAP B1 ≥ 9.2: properti ada di TitleObject
                col.TitleObject.Sortable = true;
                col.Editable = false;
                // Jika Anda memakai versi lama dan TitleObject.Sortable belum ada,
                // gunakan:  col.Sortable = true;
            }
            for (int i = 0; i < GridData.Rows.Count; i++)
            {
                GridData.CommonSetting.SetRowBackColor(i + 1, white);
            }
        }

        private void SetDataGrid(string formUID)
        {
            var oForm = Application.SBO_Application.Forms.Item(formUID);
            try
            {
                FormHelper.StartLoading(oForm, "Refreshing data table ...", 0, false);
                if (listData != null)
                {
                    BuildOrResetDataTable(oForm); // isi dt & bind ke GridData
                    GridData.AutoResizeColumns();

                    SetRowNumber();

                    // pastikan nomor tetap ketika user sort
                    if (!_sortHandlerAdded)
                    {
                        GridData.GridSortAfter += (s, e) => SetRowNumber();
                        _sortHandlerAdded = true;
                    }
                }
            }
            finally
            {
                FormHelper.FinishLoading(oForm);
            }
        }

        bool _sortHandlerAdded = false;

        private void SetRowNumber()
        {
            var grid = this.GridData;
            grid.RowHeaders.TitleObject.Caption = "#";    // judul kolom
            grid.RowHeaders.Width = 30;                   // lebar (pixel) — sesuaikan

            int rowCount = grid.DataTable.Rows.Count;
            for (int i = 0; i < rowCount; i++)
            {
                // RowHeaders indeks‑nya sama dengan indeks baris DataTable
                grid.RowHeaders.SetText(i, (i + 1).ToString());
            }
        }

        DateTime? ParseDate_yyyyMMdd(string raw)
        {
            return DateTime.TryParseExact(
                       raw?.Trim(),             // string sumber
                       "yyyyMMdd",              // format persis
                       CultureInfo.InvariantCulture,
                       DateTimeStyles.None,
                       out DateTime dt)
                   ? (DateTime?)dt
                   : null;
        }

        #endregion

    }
}