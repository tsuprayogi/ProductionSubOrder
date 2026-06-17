using ProductionOrderAddOn.Helpers;
using SAPbobsCOM;
using SAPbouiCOM.Framework;
using System;

namespace ProductionOrderAddOn
{
    public class ProdOrderListForm
    {
        private SAPbouiCOM.Form _form;
        private SAPbouiCOM.EditText _fromDate;
        private SAPbouiCOM.EditText _toDate;
        private SAPbouiCOM.EditText _searchText;
        private SAPbouiCOM.Button _btnFilter;
        private SAPbouiCOM.Grid _grid;
        private SAPbouiCOM.Matrix _matrix;

        public void Show()
        {
            try
            {
                // If form already open, just select it
                try
                {
                    var existingForm = Application.SBO_Application.Forms.Item("WorkOrderList");
                    existingForm.Select();
                    return;
                }
                catch { }

                // Create form
                var formParams = (SAPbouiCOM.FormCreationParams)
                    Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                formParams.UniqueID = "WorkOrderList";
                formParams.FormType = "PO_LIST";
                formParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;

                _form = Application.SBO_Application.Forms.AddEx(formParams);
                _form.Title = "Production Order List";
                _form.Width = 580;
                _form.Height = 410;

                // ---------------- Data Sources ----------------
                _form.DataSources.UserDataSources.Add("fromDateDS", SAPbouiCOM.BoDataType.dt_DATE);
                _form.DataSources.UserDataSources.Add("toDateDS", SAPbouiCOM.BoDataType.dt_DATE);
                _form.DataSources.UserDataSources.Add("SearchDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);

                // DataTable for Matrix
                _form.DataSources.DataTables.Add("WOData");
                var dt = _form.DataSources.DataTables.Item("WOData");
                dt.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_Integer);
                dt.Columns.Add("DocNum", SAPbouiCOM.BoFieldsType.ft_Integer);
                dt.Columns.Add("DueDate", SAPbouiCOM.BoFieldsType.ft_Date);
                dt.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("PlannedQty", SAPbouiCOM.BoFieldsType.ft_Quantity);
                dt.Columns.Add("Status", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
                dt.Columns.Add("ProdType", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);

                // ---------------- From/To Date ----------------
                var fromDateItem = _form.Items.Add("FromDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                fromDateItem.Top = 10; fromDateItem.Left = 100; fromDateItem.Width = 100;
                _fromDate = (SAPbouiCOM.EditText)fromDateItem.Specific;
                _fromDate.DataBind.SetBound(true, "", "fromDateDS");

                var lblFrom = _form.Items.Add("lblFrom", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                lblFrom.Top = 10; lblFrom.Left = 10; lblFrom.Width = 80;
                ((SAPbouiCOM.StaticText)lblFrom.Specific).Caption = "From Date";
                lblFrom.LinkTo = "FromDate";

                var toDateItem = _form.Items.Add("ToDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                toDateItem.Top = 10; toDateItem.Left = 300; toDateItem.Width = 100;
                _toDate = (SAPbouiCOM.EditText)toDateItem.Specific;
                _toDate.DataBind.SetBound(true, "", "toDateDS");

                var lblTo = _form.Items.Add("lblTo", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                lblTo.Top = 10; lblTo.Left = 220; lblTo.Width = 80;
                ((SAPbouiCOM.StaticText)lblTo.Specific).Caption = "To Date";
                lblTo.LinkTo = "ToDate";

                // ---------------- Search Box ----------------
                var searchItem = _form.Items.Add("SearchTxt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                searchItem.Left = 100; searchItem.Top = 30; searchItem.Width = 100;
                _searchText = (SAPbouiCOM.EditText)searchItem.Specific;
                _searchText.DataBind.SetBound(true, "", "SearchDS");

                var lblItem = _form.Items.Add("lblSearch", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                lblItem.Left = 10; lblItem.Top = 30; lblItem.Width = 70;
                ((SAPbouiCOM.StaticText)lblItem.Specific).Caption = "Search";
                lblItem.LinkTo = "SearchTxt";

                // ---------------- Filter Button ----------------
                var btnItem = _form.Items.Add("FilterBtn", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                btnItem.Top = 10; btnItem.Left = 420; btnItem.Width = 80;
                _btnFilter = (SAPbouiCOM.Button)btnItem.Specific;
                _btnFilter.Caption = "Filter";

                // ---------------- Matrix ----------------
                var matItem = _form.Items.Add("Mat1", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                matItem.Top = 50; matItem.Left = 10; matItem.Width = 560; matItem.Height = 320;
                _matrix = (SAPbouiCOM.Matrix)matItem.Specific;

                // DocEntry (hidden)
                SAPbouiCOM.Column colDocEntry = _matrix.Columns.Add("DocEntry", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                colDocEntry.TitleObject.Caption = "DocEntry";
                colDocEntry.Visible = false;
                colDocEntry.DataBind.Bind("WOData", "DocEntry");

                // DocNum
                SAPbouiCOM.Column colDocNum = _matrix.Columns.Add("DocNum", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                colDocNum.TitleObject.Caption = "Doc. No";
                colDocNum.Width = 80;    // set fixed width
                colDocNum.DataBind.Bind("WOData", "DocNum");
                SAPbouiCOM.LinkedButton oLink = (SAPbouiCOM.LinkedButton)colDocNum.ExtendedObject;
                oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_ProductionOrder;

                // DueDate
                SAPbouiCOM.Column colDueDate = _matrix.Columns.Add("DueDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                colDueDate.TitleObject.Caption = "Order Date";
                colDueDate.Width = 100;
                colDueDate.DataBind.Bind("WOData", "DueDate");

                // ItemCode
                SAPbouiCOM.Column colItem = _matrix.Columns.Add("ItemCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                colItem.TitleObject.Caption = "Item Code";
                colItem.Width = 120;
                colItem.DataBind.Bind("WOData", "ItemCode");

                // PlannedQty
                SAPbouiCOM.Column colQty = _matrix.Columns.Add("PlannedQty", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                colQty.TitleObject.Caption = "Planned Qty";
                colQty.Width = 80;
                colQty.DataBind.Bind("WOData", "PlannedQty");

                // Status
                SAPbouiCOM.Column colStatus = _matrix.Columns.Add("Status", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                colStatus.TitleObject.Caption = "Status";
                colStatus.Width = 60;
                colStatus.DataBind.Bind("WOData", "Status");

                // ProdType
                SAPbouiCOM.Column colProdType = _matrix.Columns.Add("ProdType", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                colProdType.TitleObject.Caption = "Prod. Type";
                colProdType.Width = 100;
                colProdType.DataBind.Bind("WOData", "ProdType");

                // Subscribe to events only once
                Application.SBO_Application.ItemEvent += OnItemEvent;

                // Show form
                _form.Visible = true;
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText("Error loading WorkOrderList: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }


       
        private void OnItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.FormUID == _form.UniqueID && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && !pVal.BeforeAction)
            {
                if (pVal.ItemUID == "FilterBtn")
                {
                    LoadProductionOrders();
                }
            }

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && !pVal.BeforeAction)
            {
                // Check Enter key
                if (pVal.CharPressed == 13)
                {
                    if (pVal.ItemUID == "FromDate" || pVal.ItemUID == "ToDate" || pVal.ItemUID == "SearchTxt")
                    {
                        LoadProductionOrders();
                    }
                }
            }

            if (pVal.FormUID == _form.UniqueID
                && pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                && pVal.BeforeAction) // only after press
            {
                if (pVal.ItemUID == "Mat1" && pVal.ColUID == "DocNum" && pVal.Row >= 0)
                {
                    SAPbouiCOM.EditText oCell = (SAPbouiCOM.EditText)_matrix.GetCellSpecific("DocEntry", pVal.Row);
                    string docEntry = oCell.Value;
                    Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_ProductionOrder, "", docEntry);
                    BubbleEvent = false; // cancel default link behavior
                }
            }

        }

        private void LoadProductionOrders()
        {
            try
            {
                //_form.Freeze(true); // Stop screen updates
                //SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(
                //    "Loading Imported Production Orders...",
                //    SAPbouiCOM.BoMessageTime.bmt_Short,
                //    SAPbouiCOM.BoStatusBarMessageType.smt_Warning
                //);
                FormHelper.StartLoading(_form, "Loading Imported Production Orders...", 0, false);

                string fromDate = _fromDate.Value.Trim();
                string toDate = _toDate.Value.Trim();
                string searchText = _searchText.Value.Trim(); // Your search input

                int currentYear = DateTime.Now.Year;
                string today = DateTime.Now.ToString("yyyyMMdd");
                string select = "SELECT DocEntry, DocNum [Doc. No], DueDate [Order Date], ItemCode [Item Code], PlannedQty [Planned Qty], Status, ISNULL(U_T2_PRODTYPE, '') [Prod. Type] FROM OWOR ";
                string where = "WHERE ISNULL(U_T2_Is_Import,'') = 'Y' ";

                // Add search filter if not empty
                if (!string.IsNullOrWhiteSpace(searchText))
                {
                    string safeSearch = searchText.Replace("'", "''");
                    where += $" AND (CAST(DocNum AS NVARCHAR) LIKE '%{safeSearch}%' OR ItemCode LIKE '%{safeSearch}%') ";
                }

                string query;

                if (string.IsNullOrWhiteSpace(fromDate) && string.IsNullOrWhiteSpace(toDate))
                {
                    query = $@"
        {select}
        {where} AND YEAR(DueDate) = {currentYear}
        ORDER BY DueDate";
                }
                else if (!string.IsNullOrWhiteSpace(fromDate) && string.IsNullOrWhiteSpace(toDate))
                {
                    query = $@"
        {select}
        {where} AND DueDate BETWEEN '{fromDate}' AND '{today}'
        ORDER BY DueDate";
                }
                else if (string.IsNullOrWhiteSpace(fromDate) && !string.IsNullOrWhiteSpace(toDate))
                {
                    string startOfYear = new DateTime(currentYear, 1, 1).ToString("yyyyMMdd");
                    query = $@"
        {select}
        {where} AND DueDate BETWEEN '{startOfYear}' AND '{toDate}'
        ORDER BY DueDate";
                }
                else
                {
                    query = $@"
        {select}
        {where} AND DueDate BETWEEN '{fromDate}' AND '{toDate}'
        ORDER BY DueDate";
                }

                var dt = _form.DataSources.DataTables.Item("WOData");
                dt.ExecuteQuery(query);

                _matrix.Clear();               // clear previous rows
                _matrix.LoadFromDataSource();  // fill matrix with new dt rows
                _matrix.AutoResizeColumns();
                int white = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                for (int i = 0; i < _matrix.Columns.Count; i++)
                {
                    _matrix.Columns.Item(i).TitleObject.Sortable = true;
                    _matrix.Columns.Item(i).Editable = false;
                }
                for (int i = 0; i < _matrix.RowCount; i++)
                {
                    _matrix.CommonSetting.SetRowBackColor(i + 1, white);
                }

                //var dt = _form.DataSources.DataTables.Item("WOData");
                //dt.ExecuteQuery(query);
                ////_grid.DataTable = dt;

                //// Enable sort on all columns
                //for (int i = 0; i < _grid.Columns.Count; i++)
                //{
                //    _grid.Columns.Item(i).TitleObject.Sortable = true;
                //    _grid.Columns.Item(i).Editable = false;
                //}

                //FormatGrid();
                //_grid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                //_grid.AutoResizeColumns();
                //_grid.CollapseLevel = 0;

                //// SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(
                ////    "Imported Production Orders loaded successfully",
                ////    SAPbouiCOM.BoMessageTime.bmt_Short,
                ////    SAPbouiCOM.BoStatusBarMessageType.smt_Success
                ////);
                FormHelper.SetTextValueLoading(_form, 0, "Imported Production Orders loaded successfully.");
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(
                    $"Error: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Medium,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
            finally
            {
                //_form.Freeze(false);
                FormHelper.FinishLoading(_form);
            }
        }

        private void FormatGrid()
        {
            //_grid.Columns.Item("DocEntry").Visible = false;
            // Set link to open existing Production Order
            //var docEntryCol = (SAPbouiCOM.EditTextColumn)_grid.Columns.Item("Doc. No");
            var docEntryCol = (SAPbouiCOM.EditTextColumn)_grid.Columns.Item("DocEntry");
            docEntryCol.LinkedObjectType = "202"; // 202 = Production Order
        }
    }
}
