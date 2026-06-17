using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ProductionOrderAddOn.Helpers;
using ProductionOrderAddOn.Services;
using SAPbobsCOM;
using SAPbouiCOM.Framework;

namespace ProductionOrderAddOn
{
    class Menu
    {
        private bool _triggeredBySubButton = false;

        private string _oldStatus = "";
        private double _oldQty = 0;

        public void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            oMenus = Application.SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = Application.SBO_Application.Menus.Item("43520"); // moudles'

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            oCreationPackage.UniqueID = "ProductionOrderAddOn";
            oCreationPackage.String = "Production Order Add On";
            oCreationPackage.Enabled = true;
            oCreationPackage.Position = -1;

            Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;

            Application.SBO_Application.RightClickEvent += SBO_Application_RightClickEvent;
            Application.SBO_Application.MenuEvent += SBO_Application_MenuEvent;


            oMenus = oMenuItem.SubMenus;

            try
            {
                //  If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception )
            {

            }

            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("ProductionOrderAddOn");
                oMenus = oMenuItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "ProductionOrderAddOn.ImportFile";
                oCreationPackage.String = "Import File Production";
                oMenus.AddEx(oCreationPackage);

              
            }
            catch (Exception )
            { //  Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        // =======================================================
        // RIGHT CLICK MENU
        // =======================================================

        private void SBO_Application_RightClickEvent(
            ref SAPbouiCOM.ContextMenuInfo eventInfo,
            out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (eventInfo.BeforeAction)
                {
                    SAPbouiCOM.Form oForm =
                        Application.SBO_Application.Forms.ActiveForm;

                    if (oForm != null && oForm.TypeEx == "65211")
                    {
                        AddCancelSubMenu();
                    }
                }
                else
                {
                    RemoveCancelSubMenu();
                }
            }
            catch { }
        }

        private void AddCancelSubMenu()
        {
            SAPbouiCOM.Form oForm =
        Application.SBO_Application.Forms.ActiveForm;

            SAPbouiCOM.DBDataSource ds =
                oForm.DataSources.DBDataSources.Item("OWOR");

            string status = ds.GetValue("Status", 0).Trim();
            string prodType = ds.GetValue("U_T2_PRODTYPE", 0).Trim();

            if (status != "R" || prodType != "FG")
                return;

            ////
            SAPbouiCOM.Menus oMenus =
                Application.SBO_Application.Menus.Item("1280").SubMenus;

            if (!oMenus.Exists("MY_PO_CANCEL"))
            {
                SAPbouiCOM.MenuCreationParams oParam =
                    (SAPbouiCOM.MenuCreationParams)
                    Application.SBO_Application.CreateObject(
                        SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

                oParam.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oParam.UniqueID = "MY_PO_CANCEL";
                oParam.String = "Cancel Sub ProdOrder";

                oMenus.AddEx(oParam);
            }
        }

        private void RemoveCancelSubMenu()
        {
            SAPbouiCOM.Menus oMenus =
                Application.SBO_Application.Menus.Item("1280").SubMenus;

            if (oMenus.Exists("MY_PO_CANCEL"))
                oMenus.RemoveEx("MY_PO_CANCEL");
        }




        private int ExtractDocEntry(string xml)
        {
            if (string.IsNullOrWhiteSpace(xml))
                throw new Exception("ObjectKey is empty, cannot extract DocEntry.");

            var doc = new System.Xml.XmlDocument();
            doc.LoadXml(xml);

            // Production Order uses AbsoluteEntry
            var node = doc.SelectSingleNode("//AbsoluteEntry");
            if (node == null || string.IsNullOrWhiteSpace(node.InnerText))
                throw new Exception("AbsoluteEntry node not found in ObjectKey XML.");

            return int.Parse(node.InnerText);
        }

        private bool ItemExists(SAPbouiCOM.Form oForm, string itemUid)
        {
            try
            {
                var item = oForm.Items.Item(itemUid);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void AddSubProdOrderButton(SAPbouiCOM.Form oForm)
        {
            try
            {
                // Hindari double add
                if (ItemExists(oForm, "btnSubPO"))
                    return;

                oForm.Freeze(true);

                // Ambil referensi button OK (ItemUID = "1")
                SAPbouiCOM.Item refItem = oForm.Items.Item("2");

                SAPbouiCOM.Item btnItem = oForm.Items.Add("btnSubPO", SAPbouiCOM.BoFormItemTypes.it_BUTTON);

                // Posisi sejajar dengan OK button
                btnItem.Left = refItem.Left + refItem.Width + 5;
                btnItem.Top = refItem.Top;
                btnItem.Width = 120;
                btnItem.Height = refItem.Height;

                SAPbouiCOM.Button btn = (SAPbouiCOM.Button)btnItem.Specific;
                btn.Caption = "Sub ProdOrder";

                // Optional: disable di Add Mode
                btnItem.Enabled = oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE;
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private List<int> GetProdTree(int rootDocEntry)
        {
            List<int> result = new List<int>();

            Company oCompany = Services.CompanyService.GetCompany();
            SAPbobsCOM.Recordset rs =
                (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string sql = $@"
                WITH ProdTree AS
                (
                    SELECT DocEntry, 0 AS [Level]
                    FROM OWOR
                    WHERE DocEntry = {rootDocEntry}

                    UNION ALL

                    SELECT C.DocEntry, P.[Level] + 1
                    FROM OWOR C
                    INNER JOIN ProdTree P
                        ON C.U_T2_Ref_Prod_DocNum = P.DocEntry
                )
                SELECT DocEntry
                FROM ProdTree
                ORDER BY [Level] DESC";

            rs.DoQuery(sql);

            while (!rs.EoF)
            {
                result.Add(Convert.ToInt32(rs.Fields.Item("DocEntry").Value));
                rs.MoveNext();
            }

            return result;
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "ProductionOrderAddOn.ImportFile")
                {
                    ImportForm activeForm = new ImportForm();
                    activeForm.Show();
                }
                if (pVal.MenuUID == "1284" && pVal.BeforeAction)
                {
                    // Check if the active form is Production Order
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                    if (oForm.TypeEx == "65211") // 65211 = Production Order form type
                    {
                        SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item("OWOR");
                        string docEntryStr = ds.GetValue("DocEntry", 0).Trim();
                        if (int.TryParse(docEntryStr, out int docEntry))
                        {
                            string prodType = ds.GetValue("U_T2_PRODTYPE", 0).Trim();
                            if (prodType == "FG")
                            {
                                Company oCompany = Services.CompanyService.GetCompany();
                                List<string> messages = new List<string>();
                                try
                                {
                                    var canceledDocs = ProductionOrderSapService.GetSubCanceled(oCompany, docEntry);
                                    if (canceledDocs.Any())
                                    {
                                        string label = canceledDocs.Count > 1 ? "Production Orders" : "Production Order";
                                        string verb = canceledDocs.Count > 1 ? "are" : "is";
                                        messages.Add($"{label} ({string.Join(", ", canceledDocs)}) {verb} canceled.");
                                    }

                                    var closedDocs = ProductionOrderSapService.GetSubClosed(oCompany, docEntry);
                                    if (closedDocs.Any())
                                    {
                                        string label = closedDocs.Count > 1 ? "Production Orders" : "Production Order";
                                        string verb = closedDocs.Count > 1 ? "are" : "is";
                                        messages.Add($"{label} ({string.Join(", ", closedDocs)}) {verb} closed.");
                                    }

                                    var receiptedDocs = ProductionOrderSapService.GetSubReceipted(oCompany, docEntry);
                                    if (receiptedDocs.Any())
                                    {
                                        string label = receiptedDocs.Count > 1 ? "Production Orders" : "Production Order";
                                        string verb = receiptedDocs.Count > 1 ? "already have" : "already has";
                                        messages.Add($"{label} ({string.Join(", ", receiptedDocs)}) {verb} Receipt from Production.");
                                    }

                                    if (messages.Any())
                                    {
                                        string finalMsg = string.Join("\n", messages);
                                        Application.SBO_Application.MessageBox(finalMsg);
                                        BubbleEvent = false; // 🚫 cancel here
                                        return;
                                    }
                                }
                                catch (Exception e)
                                {
                                    Application.SBO_Application.MessageBox(e.Message);
                                }

                            }
                        }
                    }
                }
                if (pVal.MenuUID == "1284" && !pVal.BeforeAction)
                {
                    // Check if the active form is Production Order
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                    if (oForm.TypeEx == "65211") // 65211 = Production Order form type
                    {
                        SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item("OWOR");
                        string status = ds.GetValue("Status", 0).Trim();
                        string prodType = ds.GetValue("U_T2_PRODTYPE", 0).Trim();
                        if (status == "C" && prodType == "FG")
                        {
                            int docEntry = 0;
                            if (int.TryParse(ds.GetValue("DocEntry", 0).Trim(), out int parsedEntry)) docEntry = parsedEntry;
                            CancelSubOrder(oForm,docEntry);
                        }
                    }
                }

                if (!pVal.BeforeAction && pVal.MenuUID == "MY_PO_CANCEL")
                {
                    SAPbouiCOM.Form oForm = null;

                    try
                    {
                        oForm = Application.SBO_Application.Forms.ActiveForm;
                    }
                    catch { }

                    if (oForm == null)
                        return;

                    if (oForm.TypeEx != "65211")
                        return;
                    
                    // call logic cancel

                    int docEntry = int.Parse(
                        ((SAPbouiCOM.EditText)oForm.Items.Item("18").Specific).Value.Trim());                        

                        CancelProdTree(docEntry);
                       
                    
                }

            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }


        private void CancelProdTree(int rootDocEntry)
        {
            Company oCompany = Services.CompanyService.GetCompany();
            List<int> docs = GetProdTree(rootDocEntry);

            foreach (int docEntry in docs)
            {
                SAPbobsCOM.ProductionOrders po = (SAPbobsCOM.ProductionOrders)oCompany
                    .GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders);

                if (po.GetByKey(docEntry))
                {
                    if (po.ProductionOrderStatus != BoProductionOrderStatusEnum.boposCancelled)
                    {
                        // --- STEP 1: FORCE SYNC UDF (Optional) ---
                        // We touch a UDF to ensure the object is "ready", 
                        // but we don't use the 'bomVer' variable here.
                        string currentVer = po.UserFields.Fields.Item("U_BOMVER").Value.ToString();
                        po.UserFields.Fields.Item("U_BOMVER").Value = currentVer;

                        // Update first to clear any pending line changes
                        po.Update();

                        // --- STEP 2: RELOAD & CANCEL ---
                        // Re-fetch the object to ensure a clean state for status change
                        po.GetByKey(docEntry);
                        po.ProductionOrderStatus = BoProductionOrderStatusEnum.boposCancelled;

                        int ret = po.Update();

                        if (ret != 0)
                        {
                            string err;
                            oCompany.GetLastError(out ret, out err);
                            Application.SBO_Application.MessageBox($"Sync UDF Document :  {docEntry}: {err}");
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(po);
                            return;
                        }
                    }
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(po);
            }

            Application.SBO_Application.MessageBox("Cancel document success ✅");
        }
       

        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            // Production Order Form
            if (pVal.FormTypeEx == "65211")
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && !pVal.BeforeAction)
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                    AddSubProdOrderButton(oForm);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && pVal.BeforeAction == false)
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);                 

                    // Check if form is in Add Mode (new document)
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        _oldStatus = "";
                        _oldQty = 0;
                        FormHelper.FinishLoading(oForm);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction && pVal.ItemUID == "1")
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);

                    // ✅ Run only in UPDATE MODE
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        BubbleEvent = BeforeUpdateHandler(FormUID);
                    }
                }

                if (pVal.ItemUID == "10" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                    SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item("OWOR");
                    var status = ds.GetValue("Status", 0).Trim();
                    var prodType = ds.GetValue("U_T2_PRODTYPE", 0).Trim();
                    if (prodType == "FG")
                    {
                        if (pVal.BeforeAction)
                        {
                            // Save current (old) value before change
                            _oldStatus = status;
                        }
                        else
                        {
                            // After change → validate new value
                            string newValue = status;

                            if (newValue == "P" && _oldStatus == "R")
                            {
                                SAPbouiCOM.ComboBox combo = (SAPbouiCOM.ComboBox)oForm.Items.Item("10").Specific;
                                combo.Select(_oldStatus, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                Application.SBO_Application.SetStatusBarMessage("Can't change status of Finish Good from Released to Planned.", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                            }
                        }
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ActionSuccess && pVal.ItemUID == "1")
                {
                    GenerateHandler(FormUID);
                    _triggeredBySubButton = false;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    && !pVal.BeforeAction
                    && pVal.ItemUID == "btnSubPO")
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);

                    if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        Application.SBO_Application.StatusBar.SetText(
                            "Document must be in Update mode.",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        return;
                    }

                    // Tandai bahwa OK dipicu dari button
                    _triggeredBySubButton = true;

                    // 🔥 Trigger OK SAP
                    oForm.Items.Item("1").Click();
                }

            }
        }

        private void RefreshFormProdOrder(SAPbouiCOM.Form oForm, int docEntry, string msg)
        {
            try
            {
                string sDocEntry = docEntry.ToString();
                oForm.Close(); // Close current form

                // Timer delay 1 second before reopening
                System.Timers.Timer reopenTimer = new System.Timers.Timer(1000);
                reopenTimer.AutoReset = false;
                reopenTimer.Elapsed += (sender, e) =>
                {
                    try
                    {
                        Application.SBO_Application.OpenForm(
                            SAPbouiCOM.BoFormObjectEnum.fo_ProductionOrder,
                            "",
                            sDocEntry);

                        Application.SBO_Application.StatusBar.SetText(
                            msg,
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                    catch (Exception exOpen)
                    {
                        Application.SBO_Application.StatusBar.SetText(
                            "Failed to reopen Production Order form: " + exOpen.Message,
                            SAPbouiCOM.BoMessageTime.bmt_Long,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                    finally
                    {
                        reopenTimer.Stop();
                        reopenTimer.Dispose();
                    }
                };
                reopenTimer.Start();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error during refresh: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Long,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private bool BeforeUpdateHandler(string FormUID)
        {
            bool result = true; // ✅ Default allow
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            try
            {
                FormHelper.StartLoading(oForm, "Validating Production Order...", 0, false);

                Company oCompany = Services.CompanyService.GetCompany();
                SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item("OWOR");

                string docEntryStr = ds.GetValue("DocEntry", 0).Trim();
                if (int.TryParse(docEntryStr, out int docEntry) && docEntry > 0)
                {
                    // 🔹 Load old values from DB
                    var oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRS.DoQuery($"SELECT Status, PlannedQty FROM OWOR WHERE DocEntry = {docEntry}");

                    if (!oRS.EoF)
                    {
                        _oldStatus = oRS.Fields.Item("Status").Value.ToString();
                        _oldQty = (double)oRS.Fields.Item("PlannedQty").Value;
                    }

                    // 🔹 Get new values from form buffer
                    string newStatus = ds.GetValue("Status", 0).Trim();
                    var prodType = ds.GetValue("U_T2_PRODTYPE", 0).Trim();
                    string isImported = ds.GetValue("U_T2_Is_Import", 0).Trim();
                    double plannedQty = 0;
                    double.TryParse(ds.GetValue("PlannedQty", 0).Trim(), out plannedQty);

                    // 🔹 Checks
                    if (_oldStatus != newStatus && newStatus == "R" && prodType == "FG" && isImported == "N")
                    {
                        int resultDiag = Application.SBO_Application.MessageBox(
                            "This action will generate sub production orders. Do you want to continue?",
                            1, "Yes", "No"
                        );
                        result = (resultDiag == 1);
                    }
                    else if (newStatus == "R" && prodType == "FG" && _oldQty != plannedQty)
                    {
                        int resultDiag = Application.SBO_Application.MessageBox(
                            "This action will affect the related sub production orders. Do you want to continue?",
                            1, "Yes", "No"
                        );
                        //result = (resultDiag == 1);
                        List<string> messages = new List<string>();
                        if (resultDiag == 1)
                        {
                            var canceledDocs = ProductionOrderSapService.GetSubCanceled(oCompany, docEntry);
                            if (canceledDocs.Any())
                            {
                                string label = canceledDocs.Count > 1 ? "Production Orders" : "Production Order";
                                string verb = canceledDocs.Count > 1 ? "are" : "is";
                                messages.Add($"{label} ({string.Join(", ", canceledDocs)}) {verb} canceled.");
                            }

                            var closedDocs = ProductionOrderSapService.GetSubClosed(oCompany, docEntry);
                            if (closedDocs.Any())
                            {
                                string label = closedDocs.Count > 1 ? "Production Orders" : "Production Order";
                                string verb = closedDocs.Count > 1 ? "are" : "is";
                                messages.Add($"{label} ({string.Join(", ", closedDocs)}) {verb} closed.");
                            }

                            if (messages.Any())
                            {
                                string finalMsg = string.Join("\n", messages);
                                throw new Exception(finalMsg);
                            }

                            result = true;
                        }
                        else
                        {
                            result = false;
                        }
                    }
                }
                else
                {
                    // 🔹 New document: reset values but allow add
                    _oldStatus = string.Empty;
                    _oldQty = 0;
                    result = true;
                }

                return result;
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.Message);
                //Application.SBO_Application.StatusBar.SetText(ex.Message,
                //    SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false; // Fail safe: block on exception
            }
            finally
            {
                FormHelper.FinishLoading(oForm);
            }
        }

        private void GenerateHandler(string FormUID)
        {
            int docEntry = 0;
            bool isGenerate = false;
            Company oCompany = null;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            try
            {
                oCompany = Services.CompanyService.GetCompany();
                oCompany.StartTransaction();
                SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item("OWOR");

                string docEntryStr = ds.GetValue("DocEntry", 0).Trim();
                if (int.TryParse(docEntryStr, out docEntry))
                {
                    var oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRS.DoQuery($"SELECT Status, ISNULL(U_T2_PRODTYPE,'') AS ProdType, ISNULL(PlannedQty, 0) AS PlannedQty, ISNULL(U_T2_Is_Import,'N') AS IsImported FROM OWOR WHERE DocEntry = {docEntry}");
                    string newStatus = oRS.Fields.Item("Status").Value.ToString();
                    string prodType = oRS.Fields.Item("ProdType").Value.ToString();
                    string isImported = oRS.Fields.Item("IsImported").Value.ToString();
                    double plannedQty = (double)oRS.Fields.Item("PlannedQty").Value;

                    // Cek perubahan
                    if (_oldStatus != newStatus && newStatus == "R" && prodType == "FG" && isImported == "N")
                    {
                        isGenerate = true;
                        // Tambahkan proses kamu di sini
                        FormHelper.StartLoading(oForm, "Generating sub-orders...", 0, false);

                        // Generate suborder
                        FormHelper.SetTextValueLoading(oForm, 0, "Generating Sub Production Orders...");
                        var listDoc = ProductionOrderSapService.GenerateSubOrder(oCompany, docEntry);
                        foreach (var item in listDoc)
                        {
                            int wipEntry = item.Key;
                            ProductionOrderSapService.LinkWipToFG(oCompany, docEntry, wipEntry);
                        }

                        if (listDoc != null && listDoc.Any())
                        {
                            string remarks = "Sub Production Orders: " + string.Join(" | ", listDoc.Values);
                            ProductionOrderSapService.UpdateRemarks(oCompany, docEntry, remarks);

                        }

                        RefreshFormProdOrder(oForm, docEntry, "Sub production orders successfully genareted.");
                    }
                    else if (newStatus == "R" && prodType == "FG" && _oldQty != plannedQty)
                    {
                        isGenerate = false;
                        // Tambahkan proses kamu di sini
                        FormHelper.StartLoading(oForm, "Updating sub-orders...", 0, false);

                        // Generate suborder
                        FormHelper.SetTextValueLoading(oForm, 0, "Updating WIP Production Orders...");

                        if (!ProductionOrderSapService.UpdateSubOrder(oCompany, docEntry))
                            throw new Exception("There is no documents updated.");

                        RefreshFormProdOrder(oForm, docEntry, "Sub production orders successfully updated.");
                    }
                }
                if (oCompany.InTransaction)
                {
                    oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                }

            }
            catch (Exception ex)
            {
                if (oCompany.InTransaction)
                {
                    oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                }

                Application.SBO_Application.MessageBox(ex.Message, 1, "OK");

                if (isGenerate)
                {
                    ResetStatus(oCompany, FormUID);
                }
                else
                {
                    ResetQty(oCompany, FormUID);
                }
            }
            finally
            {
                FormHelper.FinishLoading(oForm);
            }
        }

        private void ResetStatus(Company oCompany, string FormUID)
        {
            int docEntry = 0;
            try
            {
                if (!oCompany.InTransaction)
                {
                    oCompany.StartTransaction();
                }
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item("OWOR");

                string docEntryStr = ds.GetValue("DocEntry", 0).Trim();
                if (int.TryParse(docEntryStr, out docEntry))
                {
                    ProductionOrderSapService.UpdatePoStatus(oCompany, docEntry, BoProductionOrderStatusEnum.boposPlanned);
                    RefreshFormProdOrder(oForm, docEntry, $"Status reverted to planned.");
                }
                if (oCompany.InTransaction)
                {
                    oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                }
            }
            catch (Exception)
            {
                if (oCompany.InTransaction)
                {
                    oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
                throw;
            }
        }

        private void ResetQty(Company oCompany, string FormUID)
        {
            int docEntry = 0;
            try
            {
                if (!oCompany.InTransaction)
                {
                    oCompany.StartTransaction();
                }
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item("OWOR");

                string docEntryStr = ds.GetValue("DocEntry", 0).Trim();
                if (int.TryParse(docEntryStr, out docEntry))
                {
                    ProductionOrderSapService.UpdatePoQty(oCompany, docEntry, _oldQty);
                    RefreshFormProdOrder(oForm, docEntry, $"Planned Quantity reverted to previous value: {_oldQty}");
                }
                if (oCompany.InTransaction)
                {
                    oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                }
            }
            catch (Exception)
            {
                if (oCompany.InTransaction)
                {
                    oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
                throw;
            }
        }

        private void CancelSubOrder(SAPbouiCOM.Form oForm, int docEntry)
        {
            Company oCompany = null;
            try
            {
                FormHelper.StartLoading(oForm, "Cancelling sub-orders...", 0, false);
                FormHelper.SetTextValueLoading(oForm, 0, "Cancelling Sub Production Orders...");

                oCompany = Services.CompanyService.GetCompany();
                oCompany.StartTransaction();

                var rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery($"SELECT Status FROM OWOR WHERE DocEntry = {docEntry}");

                if (rs.Fields.Item("Status").Value.ToString() == "C")
                {
                    ProductionOrderSapService.CancelSubOrder(oCompany, docEntry);
                    System.Threading.ThreadPool.QueueUserWorkItem(_ =>
                    {
                        System.Threading.Thread.Sleep(1000);
                        Application.SBO_Application.StatusBar.SetText(
                            $"Sub Production Orders were cancelled.",
                            SAPbouiCOM.BoMessageTime.bmt_Long,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    });
                }

                if (oCompany.InTransaction)
                    oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
            }
            catch (Exception)
            {
                if (oCompany.InTransaction)
                    oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                throw;
            }
            finally
            {
                FormHelper.FinishLoading(oForm);
            }
        }
    }
}
