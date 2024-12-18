using SAPbobsCOM;
using SAPbouiCOM;
using System;

using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Xml;



namespace MassKaydan
{
    static class Program
    {
        private static SAPbouiCOM.Application SBO_Application;
        public static SAPbobsCOM.Company oCompany;
        public static SAPbobsCOM.Documents oInv;
        public static SAPbouiCOM.Form oForm;
        public static string selectedFilePath;
        private static SAPbouiCOM.ProgressBar oProgBar;
        private static int lineCount = 0;
        private static bool mybool = false;

        [STAThread]
        static void Main()
        {

            if (ConnectUI())
            {
                AddMenuItems();
            }
            SBO_Application.AppEvent += new _IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
            SBO_Application.ItemEvent += new _IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            SBO_Application.MenuEvent += new _IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
            SBO_Application.ProgressBarEvent += new _IApplicationEvents_ProgressBarEventEventHandler(SBO_Application_ProgressBarEvent);
            //SBO_Application.FormDataEvent += new _IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);
            System.Windows.Forms.Application.Run();
        }

        private static void SBO_Application_ProgressBarEvent(ref ProgressBarEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        static bool ConnectUI(string connectionString = "")
        {
            bool returnValue = false;
            SboGuiApi SboGuiApi = null;
            string sConnectionString;

            SboGuiApi = new SboGuiApi();
            sConnectionString = Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
            SboGuiApi.Connect(sConnectionString);
            SBO_Application = SboGuiApi.GetApplication();

            try
            {
                ConnectwithSharedMemory();
                returnValue = true;
            }
            catch (Exception exception)
            {
                var message = string.Format(CultureInfo.InvariantCulture, "{0} Initialization - Error accessing SBO: {1}", "DB_TestConnection", exception.Message);
                SBO_Application.SetStatusBarMessage("Initialisation... " + exception.Message, BoMessageTime.bmt_Short, false);
                returnValue = true;
            }
            return returnValue;
        }

        static void SBO_Application_AppEvent(BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case BoAppEventTypes.aet_CompanyChanged:
                    break;
                case BoAppEventTypes.aet_FontChanged:
                    break;
                case BoAppEventTypes.aet_LanguageChanged:
                    break;
                case BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }

        static void SBO_Application_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if ((pVal.MenuUID == "IMP") & (pVal.BeforeAction == true))
                {
                    LoadXMLSRF("udo.srf", SBO_Application, oCompany);
                    SAPbouiCOM.Form oForm;
                    oForm = SBO_Application.Forms.GetForm("U_CIMP", 1);
                    oForm.DataBrowser.BrowseBy = "Doc";
                    oForm.Visible = true;
                    //SetEditableProperties();
                }

                if (pVal.BeforeAction == true)
                {
                    if (pVal.MenuUID == "1290" || pVal.MenuUID == "1291" || pVal.MenuUID == "1289" || pVal.MenuUID == "1288")
                    {
                        /*SAPbouiCOM.Form oForm;
                        oForm = SBO_Application.Forms.GetForm("U_CIMP", 1);
                        EditText ID = (EditText)oForm.Items.Item("Item_6").Specific;
                        var DocEntry = ID.Value;*/
                        /*string Q = $@"SELECT 
                                    CASE 
                                        WHEN COUNT(*) > 0 THEN 'TRUE'
                                        ELSE 'FALSE'
                                    END AS Result
                                FROM [@IMPORT_L]
                                WHERE  (coalesce(U_InvDoc,0) != 0 OR coalesce(U_EncDoc,0) != 0)";
                        Recordset orec2;
                        orec2 = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        orec2.DoQuery(Q);
                        var res = orec2.Fields.Item(0).Value;
                        if (res == "TRUE")
                        {
                            UserObjectsMD oUserObjectsMD = null;
                            oUserObjectsMD = (UserObjectsMD)oCompany.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
                            if (!oUserObjectsMD.GetByKey("IMP0001"))
                            {
                                oUserObjectsMD.CanCancel = BoYesNoEnum.tNO;
                                oUserObjectsMD.CanClose = BoYesNoEnum.tNO;
                                oUserObjectsMD.CanDelete = BoYesNoEnum.tNO;
                            }
                            oUserObjectsMD.Update();
                        }
                        else
                        {
                            UserObjectsMD oUserObjectsMD = null;
                            oUserObjectsMD = (UserObjectsMD)oCompany.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
                            if (!oUserObjectsMD.GetByKey("IMP0001"))
                            {
                                oUserObjectsMD.CanCancel = BoYesNoEnum.tYES;
                                oUserObjectsMD.CanClose = BoYesNoEnum.tYES;
                                oUserObjectsMD.CanDelete = BoYesNoEnum.tYES;
                            }
                            oUserObjectsMD.Update();
                        }*/
                        /*if (DocEntry == "En attente")
                        {
                            Initial(false);
                        }
                        else if (DocEntry == "Validé")
                        {
                            Initial(true);
                        }
                        else if (DocEntry == "Exécuté")
                        {
                            Initial(true);
                        }
                        else
                        {
                            Initial(false);
                        }*/

                    }

                }
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        static void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                //Boite de bialogue
                if (pVal.ItemUID == "Charg" && pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false)
                {
                    SAPbouiCOM.Form oForm;
                    oForm = SBO_Application.Forms.GetForm("U_FormCH", 1);
                    Thread staThread = new Thread(() =>
                    {
                        OpenFileDialog openFileDialog = new OpenFileDialog
                        {
                            Filter = "CSV files (*.csv)|*.csv",
                            Title = "Sélectionner un fichier CSV"
                        };

                        if (openFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            selectedFilePath = openFileDialog.FileName;

                            EditText oEditText = (EditText)oForm.Items.Item("Item_5").Specific;
                            oEditText.Value = selectedFilePath;
                        }
                    });

                    staThread.SetApartmentState(ApartmentState.STA);
                    staThread.Start();
                    staThread.Join();
                }

                //Importer fichier chargé
                if (pVal.ItemUID == "valid" && pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false)
                {
                    if (string.IsNullOrEmpty(selectedFilePath))
                    {
                        SBO_Application.StatusBar.SetText("Fichier introuvable. Selectionner un fichier .CSV", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error);
                        return;
                    }

                    SAPbouiCOM.Form of;
                    of = SBO_Application.Forms.GetForm("U_FormCH", 1);

                    if (of.Mode == BoFormMode.fm_OK_MODE || of.Mode == BoFormMode.fm_UPDATE_MODE)
                    {
                        SAPbouiCOM.Form of2;
                        of2 = SBO_Application.Forms.GetForm("U_CIMP", 1);
                        of2.Mode = BoFormMode.fm_ADD_MODE;
                    }
                   
                    of.Freeze(true);
                    SBO_Application.StatusBar.SetText("Importation en cours.. ", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning);

                    Recordset oRecordset = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                    using (StreamReader reader = new StreamReader(selectedFilePath))
                    {
                        while (reader.ReadLine() != null)
                        {
                            lineCount++;
                        }
                    }
                  
                    using (StreamReader reader = new StreamReader(selectedFilePath))
                    {
                        string line;
                        int Progress = 0;
                        //oProgBar = SBO_Application.StatusBar.CreateProgressBar("Importation..", lineCount, false);

                        if ((line = reader.ReadLine()) != null)
                        {
                            // La première ligne est ignorée ici
                        }
                        of.Freeze(false);
                        of.Close();
                        of = SBO_Application.Forms.GetForm("U_CIMP", 1);
                        of.Freeze(true);
                        Matrix oMatrix = (Matrix)of.Items.Item("Mtx").Specific;

                        Columns oColumns = oMatrix.Columns;
                        int rowIndex = 0;
                        while ((line = reader.ReadLine()) != null)
                        {
                            string[] values = line.Split(';');

                            string U_DateDoc = values[0];
                            string U_CardCode = values[2];
                            string U_NameP = values[1];                           
                            string U_NumLot = values[3];
                            string U_CodePro = values[4];
                            string U_Mont = values[5];
                            string U_NatPay = values[6];
                            string U_CmpteGle= values[7];
                            string U_Bank = values[8];
                            string U_Cheq = values[9];
                            string U_Remark = values[10];
                            string U_Lib = values[11];                                                      

                            oMatrix.AddRow();

                            Column oColumnDateDoc = oColumns.Item("DateDoc");
                            Column oColumnCBP = oColumns.Item("CBP");
                            Column oColumnNameP = oColumns.Item("NameP");
                            Column oColumnLot = oColumns.Item("Lot");                           
                            Column oColumnCdePr = oColumns.Item("CdePr");
                            Column oColumnMont = oColumns.Item("Mont");
                            Column oColumnPay = oColumns.Item("Pay");
                            Column oColumnCmpteGle = oColumns.Item("Cmpte");
                            Column oColumnBank = oColumns.Item("Bank");
                            Column oColumnCheq = oColumns.Item("Cheq");
                            Column oColumnRemk = oColumns.Item("Remk");
                            Column oColumnLib = oColumns.Item("Lib");
                            Column oColumnAcc = oColumns.Item("Acc");                                                       
                            Column oColumnEnc = oColumns.Item("Enc");
                            Column oColumnLog = oColumns.Item("Log");


                            ((EditText)oColumnDateDoc.Cells.Item(rowIndex + 1).Specific).Value = U_DateDoc;
                            ((EditText)oColumnCBP.Cells.Item(rowIndex + 1).Specific).Value = U_CardCode;
                            ((EditText)oColumnNameP.Cells.Item(rowIndex + 1).Specific).Value = U_NameP;
                            ((EditText)oColumnLot.Cells.Item(rowIndex + 1).Specific).Value = U_NumLot;
                            ((EditText)oColumnCdePr.Cells.Item(rowIndex + 1).Specific).Value = U_CodePro;
                            ((EditText)oColumnMont.Cells.Item(rowIndex + 1).Specific).Value = U_Mont;
                            ((EditText)oColumnPay.Cells.Item(rowIndex + 1).Specific).Value = U_NatPay;
                            ((EditText)oColumnCmpteGle.Cells.Item(rowIndex + 1).Specific).Value = U_CmpteGle;
                            ((EditText)oColumnBank.Cells.Item(rowIndex + 1).Specific).Value = U_Bank;
                            ((EditText)oColumnCheq.Cells.Item(rowIndex + 1).Specific).Value = U_Cheq;
                            ((EditText)oColumnLib.Cells.Item(rowIndex + 1).Specific).Value = U_Lib;
                           
                            rowIndex++;
                        }
                    }
                    SBO_Application.StatusBar.SetText("Importation terminée.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                    of.Freeze(false);
                    //SetEditableProperties();
                }

                //Ouvrir le formulaire
                if (pVal.ItemUID == "Load" && pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false)
                {
                    SAPbouiCOM.Form of;
                    /*of = SBO_Application.Forms.GetForm("U_CIMP", 1);
                    of.Mode = BoFormMode.fm_ADD_MODE;*/
                    
                    LoadXMLSRF("imp.srf", SBO_Application, oCompany);
                   
                    of = SBO_Application.Forms.GetForm("U_FormCH", 1);
                    of.Left = 600;
                    of.Top = 300;
                    of.Visible = true;
                    //OpenAndLockForm("U_FormCH");
                }

                //Fermer la forme de selection
                if (pVal.ItemUID == "Can" && pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false)
                {
                    SAPbouiCOM.Form of;
                    of = SBO_Application.Forms.GetForm("U_FormCH", 1);
                    of.Close();
                }

                //Creation facture/Encaissement
                if (pVal.ItemUID == "exec" && pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false)
                {
                    
                    SAPbouiCOM.Form of;
                    of = SBO_Application.Forms.GetForm("U_CIMP", 1);
                    if (of.Mode != BoFormMode.fm_OK_MODE)
                    {
                        SBO_Application.StatusBar.SetText("L'importation n'est pas sauvegardée. ", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error);
                        return;
                    }
                    of.Freeze(true);
                    SBO_Application.StatusBar.SetText("Traitement en cours.. ", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning);

                    EditText doc = (EditText)of.Items.Item("Doc").Specific;
                    var ID = doc.Value;
                    string q = $@"SELECT * FROM [dbo].[@IMPORT_L] T0 WHERE T0.[DocEntry] = {int.Parse(ID)} AND coalesce(T0.U_EncDoc,0) = 0";
                    Recordset orec;
                    Recordset orec2;
                    orec2 = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    orec = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    orec.DoQuery(q);
                    //oProgBar = SBO_Application.StatusBar.CreateProgressBar("Execution..", orec.RecordCount, false);
                    int Progress = 0;
                    int total = orec.RecordCount;
                    while (!orec.EoF)
                    {
                        string CardCode = (string)orec.Fields.Item("U_CardCode").Value;
                        string NaturePay = (string)orec.Fields.Item("U_NatPay").Value;
                        string U_Remark = (string)orec.Fields.Item("U_Remk").Value;
                        string U_Lib = (string)orec.Fields.Item("U_Lib").Value;

                        double montant = Convert.ToDouble(orec.Fields.Item("U_Total").Value);
                        DateTime docDate = Convert.ToDateTime(orec.Fields.Item("U_DateDoc").Value);
                        
                        string DocEntryUDO = orec.Fields.Item("DocEntry").Value.ToString();
                        string lineId = orec.Fields.Item("LineId").Value.ToString(); 
                        string projet = orec.Fields.Item("U_CodePro").Value.ToString();

                        string temp_string = "";
                        Documents oDP;
                        oDP = (Documents)oCompany.GetBusinessObject(BoObjectTypes.oDownPayments);
                        oDP.CardCode = CardCode;
                        oDP.DocDate = docDate;
                        oDP.DocDueDate = docDate;
                        //oDP.DownPaymentPercentage = Convert.ToDouble(numPercent.Value);
                        oDP.DownPaymentType = DownPaymentTypeEnum.dptRequest;
                        oDP.Lines.ItemCode = "A00001";
                        oDP.Lines.UnitPrice = montant;
                        oDP.Lines.Add();
                        int ret = oDP.Add();
                        if (ret != 0)
                        {
                            SBO_Application.StatusBar.SetText("Erreur: "+ oCompany.GetLastErrorDescription(), BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error);
                        }
                        else
                        {
                            
                            oCompany.GetNewObjectCode(out temp_string);
                            string updateQuery = $"UPDATE [@IMPORT_L] SET U_Acc = {int.Parse(temp_string)} WHERE DocEntry = {DocEntryUDO} and LineId = {lineId}";
                            orec2.DoQuery(updateQuery);
                        }
                        if (!string.IsNullOrEmpty(temp_string))
                        {
                            Payments oPayment = (Payments)oCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments);

                            oDP.GetByKey(int.Parse(temp_string));

                            oPayment.CardCode = CardCode;
                            oPayment.DocDate = docDate;
                            oPayment.Remarks = U_Remark;
                            oPayment.JournalRemarks = U_Lib;
                            oPayment.Invoices.DocEntry = int.Parse(temp_string);
                            oPayment.Invoices.InvoiceType = BoRcptInvTypes.it_DownPayment;
                            oPayment.CashSum = montant;
                            oPayment.CashAccount = "511701";
                            //oPayment.ProjectCode = projet;
                            /*if (NaturePay.StartsWith("Vir"))
                            {
                                oPayment.TransferAccount = "52110001";
                                oPayment.TransferSum = montant;
                                oPayment.TransferDate = docDate;

                            }
                            else if (NaturePay.StartsWith("Es"))
                            {
                                oPayment.CashSum = montant;
                                oPayment.CashAccount = "57110000";
                            }
                            else
                            {
                                oPayment.CheckAccount = "51410000";
                                oPayment.Checks.CheckSum = montant;
                                oPayment.Checks.BankCode = "CI008";
                            }*/

                            if (oPayment.Add() != 0)
                            {
                                string updateQuery = $"UPDATE [@IMPORT_L] SET U_Log = '{oCompany.GetLastErrorDescription()}'  WHERE DocEntry = {DocEntryUDO} and LineId = {lineId}";
                                orec2.DoQuery(updateQuery);
                            }
                            else
                            {
                                string DocEntrEnc = oCompany.GetNewObjectKey();
                                string updateQuery = $"UPDATE [@IMPORT_L] SET U_EncDoc = {int.Parse(DocEntrEnc)}, U_Log = '' WHERE DocEntry = {DocEntryUDO} and LineId = {lineId}";
                                orec2.DoQuery(updateQuery);
                            }
                        }
                        
                        orec.MoveNext();
                        Progress += 1;
                    }

                    Matrix mt = (Matrix)of.Items.Item("Mtx").Specific;
                    mt.LoadFromDataSource();
                    of.Refresh();
                    of.Visible = true;

                    SBO_Application.ActivateMenuItem("1304");

                    SBO_Application.StatusBar.SetText("Operation terminée.", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Success);
                    of.Freeze(false);
                }
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox("Exception: " + ex.Message);
            }
            finally
            {
                if (oProgBar != null)
                {
                    //oProgBar.Stop();
                    //Marshal.ReleaseComObject(oProgBar);
                    //oProgBar = null;
                }
            }
        }

        static void InsertIntoUDO(int docentry, int LineId, string U_ItemCode, string U_Quantity, string U_Price, string U_CardCode, string U_DocDate, string U_Total, Recordset oRecordset)
        {
            try
            {
                string query = $@"INSERT INTO [@IMPORT_L] (DocEntry, LineId, U_ItemCode, U_Quantity, U_Price, U_CardCode, U_DocDate, U_Total) 
                VALUES ({docentry}, {LineId}, '{U_ItemCode}', {U_Quantity}, {U_Price}, '{U_CardCode}', '{U_DocDate}', {U_Total})";
                oRecordset.DoQuery(query);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox("Erreur lors de l'insertion dans l'UDO : " + ex.Message);
            }
        }

        static void ConnectwithSharedMemory()
        {
            oCompany = new SAPbobsCOM.Company();
            oCompany = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();
            SBO_Application.SetStatusBarMessage("Initialisation... " + oCompany.CompanyName, BoMessageTime.bmt_Short, false);
            SBO_Application.StatusBar.SetText("SOAS connected sucessfully to " + oCompany.CompanyName, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
        }

        static void AddMenuItems()
        {
            RunDBScript();
            Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            oMenus = SBO_Application.Menus;

            MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((MenuCreationParams)(SBO_Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams)));
            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = SBO_Application.Menus.Item("2048");
                oMenus = oMenuItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "IMP";
                oCreationPackage.String = "Importation Encaissements";
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception)
            { //  Menu already exists
                SBO_Application.SetStatusBarMessage("Menu Already Exists", BoMessageTime.bmt_Short, true);
            }
        }

        static void SetEditableProperties(string oRef = "")
        {
            SAPbouiCOM.Form oForm;
            oForm = SBO_Application.Forms.GetForm("U_CIMP", 1);
            oForm.Freeze(true);

            Matrix oMatrix = (Matrix)oForm.Items.Item("Mtx").Specific;

            oMatrix.LoadFromDataSource();
            oForm.DataBrowser.BrowseBy = "Doc";

            oForm.State = BoFormStateEnum.fs_Maximized;
            oForm.Visible = true;
            SBO_Application.ActivateMenuItem("1291");
            /*EditText oEdi = (EditText)oForm.Items.Item("Item_6").Specific;
            string stat = oEdi.Value;

            if (stat == "En attente")
            {
                Initial(false);
            }
            else if (stat == "Validé")
            {
                Initial(true);
            }
            else if (stat == "Exécuté")
            {
                Initial(true);
            }
            else
            {
                Initial(false);
            }*/

            oForm.Freeze(false);
        }

        static void LoadXMLSRF(string oPath, SAPbouiCOM.Application SBO_Application, SAPbobsCOM.Company oCompany)
        {
            XmlDocument oXml = new XmlDocument();
            string ls_Xml = "";
            Assembly AppAssembly = Assembly.GetEntryAssembly();
            Stream SRFFile = AppAssembly.GetManifestResourceStream("MassKaydan.SRF." + oPath);
            oXml.Load(SRFFile);
            ls_Xml = oXml.InnerXml.ToString();
            SBO_Application.LoadBatchActions(ref ls_Xml);           
        }

        static private bool CreateUDTObject(SAPbobsCOM.Company oCompany, string TableName, string Description, BoUTBTableType TableType)
        {
            UserTablesMD oUserTablesMD = null;
            int iRetCode = 0;
            string sErrMsg = null;

            try
            {
                oUserTablesMD = (UserTablesMD)(oCompany.GetBusinessObject(BoObjectTypes.oUserTables));
                if (!(oUserTablesMD.GetByKey(TableName)))
                {
                    oUserTablesMD.TableName = TableName;
                    oUserTablesMD.TableDescription = Description;
                    oUserTablesMD.TableType = TableType;
                    iRetCode = oUserTablesMD.Add();
                    if (iRetCode != 0)
                    {
                        oCompany.GetLastError(out iRetCode, out sErrMsg);
                        return false;
                    }
                    else
                    {
                        return true;
                    }

                }
                else
                {
                    return false;
                }

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                Marshal.ReleaseComObject(oUserTablesMD);
            }
        }

        static private bool CreateUDFObject(SAPbobsCOM.Company oCompany, string TableName, string UDFName, string UDFDescription, BoFieldTypes UDFType, BoFldSubTypes UDFSubtype, int UDFLength, string DefaultValue, BoYesNoEnum Mandatory)
        {
            UserTablesMD oUserTablesMD = null;
            int iRetCod = 0;
            string sErrMsg = "";
            UserFieldsMD FieldsMD = null;

            try
            {
                FieldsMD = (UserFieldsMD)(oCompany.GetBusinessObject(BoObjectTypes.oUserFields));
                if (!IsUDFExists(oCompany, TableName, UDFName))
                {
                    FieldsMD.TableName = TableName;
                    FieldsMD.Name = UDFName;
                    FieldsMD.Description = UDFDescription;
                    FieldsMD.Type = UDFType;
                    FieldsMD.SubType = UDFSubtype;

                    if (UDFType != BoFieldTypes.db_Float)
                    {
                        FieldsMD.EditSize = UDFLength;
                    }
                    else
                    {
                        FieldsMD.EditSize = 0;

                    }
                    if (DefaultValue != null && DefaultValue != "")
                    {
                        FieldsMD.DefaultValue = DefaultValue;
                    }
                    if (Mandatory == BoYesNoEnum.tYES || Mandatory == BoYesNoEnum.tNO)
                    {
                        FieldsMD.Mandatory = Mandatory;
                    }
                    iRetCod = FieldsMD.Add();
                    if (iRetCod != 0)
                    {
                        oCompany.GetLastError(out iRetCod, out sErrMsg);
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
                else
                {
                    return true;
                }

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                Marshal.ReleaseComObject(FieldsMD);
            }

        }

        static private bool CreateUDOObject(SAPbouiCOM.Application Application, SAPbobsCOM.Company Company, string UDOCode, string UDOName, BoUDOObjType UDOType, string FindColumnDesc, string HeaderTableName, string[] childTables, BoYesNoEnum LogOption)
        {
            UserObjectsMD oUserObjectsMD = null;

            try
            {
                int iRetCod = 0; string sErrMsg = string.Empty;
                oUserObjectsMD = (UserObjectsMD)(Company.GetBusinessObject(BoObjectTypes.oUserObjectsMD));
                if (!oUserObjectsMD.GetByKey(UDOCode))
                {
                    oUserObjectsMD.CanCancel = BoYesNoEnum.tNO;
                    oUserObjectsMD.CanClose = BoYesNoEnum.tNO;
                    oUserObjectsMD.CanCreateDefaultForm = BoYesNoEnum.tNO;
                    oUserObjectsMD.CanDelete = BoYesNoEnum.tNO;
                    oUserObjectsMD.CanFind = BoYesNoEnum.tYES;
                    oUserObjectsMD.CanLog = LogOption;
                    oUserObjectsMD.CanYearTransfer = BoYesNoEnum.tNO;
                    oUserObjectsMD.ManageSeries = BoYesNoEnum.tNO;
                    oUserObjectsMD.FormColumns.FormColumnAlias = FindColumnDesc;
                    oUserObjectsMD.Code = UDOCode;
                    oUserObjectsMD.Name = UDOName;
                    oUserObjectsMD.ObjectType = UDOType;
                    oUserObjectsMD.TableName = HeaderTableName;

                    if (LogOption == BoYesNoEnum.tYES)
                    {
                        oUserObjectsMD.LogTableName = "A" + HeaderTableName;
                    }
                    if (!string.IsNullOrEmpty(FindColumnDesc))
                    {
                        oUserObjectsMD.FindColumns.ColumnAlias = FindColumnDesc;
                    }
                    for (int i = 0; i <= childTables.Length - 1; i++)
                    {
                        oUserObjectsMD.ChildTables.TableName = childTables[i].ToString();
                        oUserObjectsMD.ChildTables.Add();
                    }
                    iRetCod = oUserObjectsMD.Add();
                    if (iRetCod != 0)
                    {
                        Company.GetLastError(out iRetCod, out sErrMsg);
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
                else
                {
                    return true;
                }

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                Marshal.ReleaseComObject(oUserObjectsMD);
            }

        }

        static private bool IsUDFExists(SAPbobsCOM.Company Company, string TableName, string UDFName)
        {
            Recordset RecordSet = null;
            try
            {
                string SqlQuery = null;
                RecordSet = (Recordset)Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                SqlQuery = "SELECT FieldID FROM [CUFD] WHERE TableID='" + TableName + "' AND AliasID ='" + UDFName + "'";
                RecordSet.DoQuery(SqlQuery);
                if (RecordSet.RecordCount > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                Marshal.ReleaseComObject(RecordSet);

            }
        }

        static private void Initial(SAPbouiCOM.Form of, bool etat)
        {
            Matrix oMatrix = (Matrix)of.Items.Item("Mtx").Specific;
            if (etat)
            {
                oMatrix.Columns.Item(0).Editable = false;
                oMatrix.Columns.Item(1).Editable = false;
                oMatrix.Columns.Item(2).Editable = false;
                oMatrix.Columns.Item(3).Editable = false;
                oMatrix.Columns.Item(4).Editable = false;
                oMatrix.Columns.Item(5).Editable = false;
                oMatrix.Columns.Item(6).Editable = false;
            }
            else
            {
                oMatrix.Columns.Item(0).Editable = true;
                oMatrix.Columns.Item(1).Editable = true;
                oMatrix.Columns.Item(2).Editable = true;
                oMatrix.Columns.Item(3).Editable = true;
                oMatrix.Columns.Item(4).Editable = true;
                oMatrix.Columns.Item(5).Editable = true;
                oMatrix.Columns.Item(6).Editable = true;
            }

        }

        static private bool RunDBScript()
        {
            try
            {
                if (CreateUDTObject(oCompany, "IMPORT_H", "Customized Import Header", BoUTBTableType.bott_Document))
                {
                    SBO_Application.StatusBar.SetText("Customized Import Header added sucessfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                if (CreateUDFObject(oCompany, "IMPORT_H", "Notes", "Remarque", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, null, BoYesNoEnum.tNO))
                {
                    SBO_Application.StatusBar.SetText("CardCode added sucessfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                if (CreateUDTObject(oCompany, "IMPORT_L", "Customized Import Line", BoUTBTableType.bott_DocumentLines))
                {
                    SBO_Application.StatusBar.SetText("Customized Import Line added sucessfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                if (CreateUDFObject(oCompany, "IMPORT_L", "DateDoc", "Date Doc.", BoFieldTypes.db_Date, BoFldSubTypes.st_None, 10, null, BoYesNoEnum.tNO))
                {
                    SBO_Application.StatusBar.SetText("Date Document sucessfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                if (CreateUDFObject(oCompany, "IMPORT_L", "CardCode", "Code Partenaire", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, null, BoYesNoEnum.tYES))
                {
                    SBO_Application.StatusBar.SetText("Code Partenaire added sucessfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                if (CreateUDFObject(oCompany, "IMPORT_L", "CardName", "Nom du Partenaire", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, null, BoYesNoEnum.tNO))
                {
                    SBO_Application.StatusBar.SetText("Nom du Partenaire added sucessfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                if (CreateUDFObject(oCompany, "IMPORT_L", "NumLot", "Numero Lot", BoFieldTypes.db_Numeric, BoFldSubTypes.st_Quantity, 9, null, BoYesNoEnum.tYES))
                {
                    SBO_Application.StatusBar.SetText("Num Lot added sucessfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                if (CreateUDFObject(oCompany, "IMPORT_L", "CodePro", "Code Projet", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 50, null, BoYesNoEnum.tNO))
                {
                    SBO_Application.StatusBar.SetText("Code Projet added sucessfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                if (CreateUDFObject(oCompany, "IMPORT_L", "Total", "Montant Enc.", BoFieldTypes.db_Numeric, BoFldSubTypes.st_Quantity, 9, null, BoYesNoEnum.tYES))
                {
                    SBO_Application.StatusBar.SetText("Montant Total added sucessfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                if (CreateUDFObject(oCompany, "IMPORT_L", "NatPay", "Mode de Paiement", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, null, BoYesNoEnum.tYES))
                {
                    SBO_Application.StatusBar.SetText("Mode de Paiement added sucessfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                if (CreateUDFObject(oCompany, "IMPORT_L", "CteGle", "Compte Général", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, null, BoYesNoEnum.tYES))
                {
                    SBO_Application.StatusBar.SetText("Compte Général added sucessfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                if (CreateUDFObject(oCompany, "IMPORT_L", "Bank", "Banque", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, null, BoYesNoEnum.tYES))
                {
                    SBO_Application.StatusBar.SetText("Banque added sucessfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                if (CreateUDFObject(oCompany, "IMPORT_L", "Cheq", "N° Cheque", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, null, BoYesNoEnum.tNO))
                {
                    SBO_Application.StatusBar.SetText("N° Cheque added sucessfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                if (CreateUDFObject(oCompany, "IMPORT_L", "Remk", "Remarque", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, null, BoYesNoEnum.tNO))
                {
                    SBO_Application.StatusBar.SetText("Remarque added sucessfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                if (CreateUDFObject(oCompany, "IMPORT_L", "Lib", "Libellé", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, null, BoYesNoEnum.tNO))
                {
                    SBO_Application.StatusBar.SetText("Libellé added sucessfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                if (CreateUDFObject(oCompany, "IMPORT_L", "Acc", "N° Accompte", BoFieldTypes.db_Numeric, BoFldSubTypes.st_None, 9, null, BoYesNoEnum.tNO))
                {
                    SBO_Application.StatusBar.SetText("N° Accompte added sucessfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }                
                if (CreateUDFObject(oCompany, "IMPORT_L", "EncDoc", "No. Encaissement", BoFieldTypes.db_Numeric, BoFldSubTypes.st_None, 9, null, BoYesNoEnum.tNO))
                {
                    SBO_Application.StatusBar.SetText("Initial Qty added sucessfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                if (CreateUDFObject(oCompany, "IMPORT_L", "Log", "Log/Erreur", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 120, null, BoYesNoEnum.tNO))
                {
                    SBO_Application.StatusBar.SetText("Log/Erreur added sucessfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                string[] childTables = new string[1];
                childTables[0] = "IMPORT_L";
                if (CreateUDOObject(SBO_Application, oCompany, "IMP0001", "Encaissements Importés", BoUDOObjType.boud_Document, "DocEntry", "IMPORT_H", childTables, SAPbobsCOM.BoYesNoEnum.tYES))
                {
                    SBO_Application.StatusBar.SetText("IMP0001 Object created successfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

    }
}
