using ADDONBASE.Extensions;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
namespace ADDONBASE
{
    public partial class _UDOFormBase : UDOFormBase
    {
        public _UDOFormBase()
        {
            try
            {
                _Initializer.SBO_Application.ItemEvent -= _SBO_Application_ItemEvent1;
            }
            catch { }
            _Initializer.SBO_Application.ItemEvent += _SBO_Application_ItemEvent1;

        }
        protected void ExtractQuery(string query, string queryName)
        {
            var outputPath = Path.Combine(Path.GetTempPath(), queryName);

            System.IO.File.WriteAllText(outputPath, query);
        }
        #region properties
        protected DBDataSource MasterDS
        {
            get
            {
                return CurrentForm.DataSources.DBDataSources.Item(0);
            }
        }
        protected DBDataSource DetailDS
        {
            get
            {
                return CurrentForm.DataSources.DBDataSources.Item(1);
            }
        }
        protected DBDataSource DetailDS2
        {
            get
            {
                return CurrentForm.DataSources.DBDataSources.Item(2);
            }
        }
        #endregion
        #region MyCode
        protected SAPbobsCOM.Recordset GetRecordSet(String Query, params object[] args)
        {

            var recset = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            recset.DoQuery(string.Format(Query, args));
            recset.MoveFirst();
            return recset;
        }
        protected bool ForcePLDEnabled = false;

        protected string getObjectKeyFromXML(String XML)
        {
            return Extensions.Extensions.getObjectKeyFromXML(XML);
        }
        int Last_index = 0;
        private Dictionary<string, string> _LstAutoAddRow = new Dictionary<string, string>();
        protected void Add_AUTO_Matrix(string MatrixID, string ColUID)
        {
            try
            {
                _LstAutoAddRow.Add(MatrixID, ColUID);
                if (GetItem(MatrixID).Specific is SAPbouiCOM.Matrix)
                {
                    var mat = GetItem(MatrixID).Specific as SAPbouiCOM.Matrix;
                    mat.AutoResizeColumns();
                    //mat.LostFocusAfter += (object sboObject, SBOItemEventArg pVal) =>
                    //{
                    //    _OnValidate(pVal);
                    //};
                    mat.ValidateAfter += mat_ValidateAfter;
                }
                AddMatrix(MatrixID);
            }
            catch (Exception ex)
            {
                ex.AppendInLogFile();
            }
        }
        protected virtual void AutoMatrixValidateAfter(SBOItemEventArg pVal) { }
        void mat_ValidateAfter(object sboObject, SBOItemEventArg pVal)
        {

            var mat = (SAPbouiCOM.Matrix)sboObject;
            SAPbouiCOM.EventFilters currentfilter = Application.GetFilter();
            Application.SetFilter(null);
            //mat.ValidateAfter -= mat_ValidateAfter;
            _OnValidate(pVal);
            AutoMatrixValidateAfter(pVal);
            //mat.ValidateAfter += mat_ValidateAfter;
            Application.SetFilter(currentfilter);
        }



        private bool _OnValidate(SAPbouiCOM.SBOItemEventArg pVal)
        {
            bool yes = true;
            try
            {
                if (_LstAutoAddRow.ContainsKey(pVal.ItemUID))
                {

                    var ColUID = _LstAutoAddRow[pVal.ItemUID];
                    string MatrixId = string.Empty;
                    if (!string.IsNullOrEmpty(ColUID))
                    {
                        MatrixId = pVal.ItemUID;
                    }
                    if (pVal.ItemUID == MatrixId && pVal.ColUID == ColUID)
                    {
                        var matrix = GetItem(pVal.ItemUID).Specific as SAPbouiCOM.Matrix;

                        if (matrix.RowCount == pVal.Row)
                        {
                            var tablname = matrix.Columns.Item(1).DataBind.TableName;
                            matrix.FlushToDataSource();
                            var ds = CurrentForm.DataSources.DBDataSources.Item(tablname);
                            var alias = matrix.Columns.Item(ColUID).DataBind.Alias;
                            var value = ds.GetValue(alias, ds.Size - 1).Trim();
                            var b = false;
                            if (ds.Fields.Item(alias).Type == BoFieldsType.ft_Text || ds.Fields.Item(alias).Type == BoFieldsType.ft_Date || ds.Fields.Item(alias).Type == BoFieldsType.ft_AlphaNumeric)
                            {
                                if (!string.IsNullOrEmpty(value))
                                {
                                    ds.InsertRecord(ds.Size);
                                    b = true;
                                }
                            }
                            else
                            {
                                var _value = 0.0;
                                try
                                {
                                    _value = Convert.ToDouble(value);
                                }
                                catch { }
                                if (_value != 0.0)
                                {
                                    ds.InsertRecord(ds.Size);
                                    b = true;
                                }
                            }

                            matrix.LoadFromDataSourceEx(false);
                            if (b)
                                OnMatrixRowAdded(MatrixId, ds.Size);
                        }
                    }
                    yes = true;
                }
            }
            catch (Exception ex)
            {
                ex.AppendInLogFile();
            }
            return yes;
        }


        protected void Add_AUTO_Matrix(string MatrixID)
        {
            try
            {
                _LstAutoAddRow.Add(MatrixID, "");
                if (GetItem(MatrixID).Specific is SAPbouiCOM.Matrix)
                {
                    var mat = GetItem(MatrixID).Specific as SAPbouiCOM.Matrix;
                    mat.GotFocusAfter += (object sboObject, SBOItemEventArg pVal) =>
                      {
                          try
                          {

                              if (pVal.Row == mat.RowCount)
                              {
                                  var tablname = mat.Columns.Item(1).DataBind.TableName;
                                  mat.FlushToDataSource();
                                  var ds = CurrentForm.DataSources.DBDataSources.Item(tablname);
                                  ds.InsertRecord(ds.Size);
                                  mat.LoadFromDataSourceEx();
                                  mat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();

                              }
                              OnMatrixRowAdded(MatrixID, pVal.Row + 1);

                          }
                          catch (Exception ex)
                          {

                              ex.AppendInLogFile();
                          }
                          finally
                          {
                          }
                      };
                    //        mat.ValidateAfter += mat_ValidateAfter;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }




        //private List<string> ManagedMatrix = new List<string>();
        /// <summary>
        /// User this function to handle Lost focus of control from matrix when you click on Choose From List
        /// </summary>
        /// <param name="Matrix"></param>
        protected void AddMatrix(string Matrix)
        {
            if (GetItem(Matrix).Specific is SAPbouiCOM.Matrix)
            {
                var mat = GetItem(Matrix).Specific as SAPbouiCOM.Matrix;
                mat.ChooseFromListBefore += (object sboObject, SBOItemEventArg pVal, out bool BubbleEvent) =>
                {
                    if (!string.IsNullOrEmpty(pVal.ColUID))
                    {
                        mat.FlushToDataSource();
                        var ColumnIndex = mat.GetCellFocus().ColumnIndex;
                        matrixid = pVal.ItemUID;
                        IsOnMatrix = 1;
                        IsRow = pVal.Row;
                        IsCol = ColumnIndex;
                    }
                    BubbleEvent = MatrixChoosFromListBefore(sboObject, pVal);

                };
                mat.ChooseFromListAfter +=
                     (object sboObject, SBOItemEventArg pVal) =>
                     {
                         ISBOChooseFromListEventArg cflitemevent;
                         string key = "";
                         string ColUID = pVal.ColUID;
                         int row = pVal.Row;
                         try
                         {
                             cflitemevent = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                             key = cflitemevent.SelectedObjects.GetValue(0, 0).ToString().Trim();
                         }
                         catch { }
                         finally
                         {
                         }

                         //if (IsOnMatrix == 1)
                         //{
                         Task tsk = new Task(() =>
                         {


                             // mat.FlushToDataSource();
                             if (IsRow == 0) IsRow = 1;
                             var t = true;
                             while (t)
                             {
                                 try
                                 {
                                     t = CurrentForm.UniqueID != Application.Forms.ActiveForm.UniqueID;
                                 }
                                 catch (Exception ex)
                                 {
                                     // ex.AppendInLogFile();
                                 }
                             }

                             mat.Columns.Item(IsCol).Cells.Item(IsRow).Click();
                             //mat.SetCellFocus(IsRow, IsCol);
                             IsOnMatrix = 0;
                             MatrixChoosFromListAfter(sboObject, row, key, ColUID);
                             try { MatrixChoosFromListAfter(sboObject, row, key, ColUID, pVal); }
                             catch { }

                         });
                         tsk.Start();

                         // }
                     };
            }
        }

        protected virtual void MatrixChoosFromListAfter(object sboObject, int row, string key, string ColUID)
        {
        }
        protected virtual void MatrixChoosFromListAfter(object sboObject, int row, string key, string ColUID, SBOItemEventArg pval)
        {
        }


        protected virtual bool MatrixChoosFromListBefore(object sboObject, SBOItemEventArg pVal)
        {
            return true;
        }


        string matrixid = ""; Int32 IsOnMatrix = 0, IsRow = 0, IsCol = 0;



        protected bool ForceViewModeonClose = false;
        private List<string> _ItemToClose = new List<string>();
        protected void ApplyDisableOnClose(string ItemID)
        {
            if (!_ItemToClose.Contains(ItemID))
            {
                _ItemToClose.Add(ItemID);
            }


        }
        void EnforceViewModeOnCLose()
        {
            try
            {
                if (_ItemToClose.Count > 0 && MasterDS.GetValue("Status", 0).Trim().ToLower() == "c")
                {
                    CurrentForm.Freeze(true);
                    foreach (var item in _ItemToClose)
                    {
                        GetItem(item).Enabled = false;
                    }

                    CurrentForm.Freeze(false);
                }
                else
                {

                    CurrentForm.Freeze(true);
                    foreach (var item in _ItemToClose)
                    {
                        GetItem(item).Enabled = true;
                    }

                    CurrentForm.Freeze(false);
                }
            }
            catch (Exception ex) { ex.AppendInLogFile(); }
            if (ForceViewModeonClose)
            {
                if (MasterDS.GetValue("Status", 0).Trim() == "C" && CurrentForm.Mode != BoFormMode.fm_VIEW_MODE)
                {

                    CurrentForm.Mode = BoFormMode.fm_VIEW_MODE;
                }
                if (MasterDS.GetValue("Status", 0).Trim() == "O" && CurrentForm.Mode == BoFormMode.fm_VIEW_MODE)
                {
                    CurrentForm.Mode = BoFormMode.fm_OK_MODE;
                }
            }
        }







        #endregion

        protected SAPbobsCOM.Company Company
        {
            get
            {
                return _Initializer.Company;
            }
        }
        protected SAPbouiCOM.Application Application
        {
            get
            {
                return _Initializer.SBO_Application;
            }
        }

        private void _SBO_Application_ItemEvent1(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            try
            {
                if (CurrentForm.UniqueID == FormUID)
                    _SBO_Application_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                else BubbleEvent = true;
            }
            catch { BubbleEvent = true; }
            try
            {
                if (BubbleEvent)
                {
                    SBO_Application_ItemEvent_ForAllForms(FormUID, ref pVal, out BubbleEvent);
                }
            }
            catch { }
        }

        protected virtual void _SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }
        protected virtual void SBO_Application_ItemEvent_ForAllForms(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }
        protected SAPbouiCOM.IForm CurrentForm
        {
            get
            {
                return this.UIAPIRawForm;
            }
        }
        public override void OnInitializeComponent()
        {
            // CurrentForm.Mode = BoFormMode.fm_ADD_MODE;

            base.OnInitializeComponent();
        }
        protected override void OnFormVisibleAfter(SBOItemEventArg pVal)
        {
            if (CurrentForm.Visible)
            {
                CurrentForm.Freeze(true);

                if (Application.Menus.Item("1297").Enabled)
                    Application.Menus.Item("1297").Activate();

                if (ForcePLDEnabled)
                    try
                    {
                        PLDAdder.PLDAdder pldadder = new PLDAdder.PLDAdder();
                        CurrentForm.ReportType = pldadder.getReportTypeCode(CurrentForm.Title, CurrentForm.BusinessObject.Type, Initializer.ADDON_NAME, CurrentForm.BusinessObject.Type);
                    }
                    catch { }
                Application.MenuEvent -= Application_MenuEvent;

                Application.MenuEvent += Application_MenuEvent;

                CurrentForm.Freeze(false);
            }
            base.OnFormVisibleAfter(pVal);
        }

        protected override void OnFormCloseAfter(SBOItemEventArg pVal)
        {
            Application.MenuEvent -= Application_MenuEvent;
            base.OnFormCloseAfter(pVal);
        }

        void Application_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            try
            {
                if (Application.Forms.ActiveForm.UniqueID == CurrentForm.UniqueID && pVal.BeforeAction)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1286":
                            {
                                _Initializer.IsMenuResultClear = _Initializer.IsMenuResultClear && BeforeCloseMenuClicked();
                            }
                            break;
                    }
                    _Initializer.IsMenuResultClear = _Initializer.IsMenuResultClear && BeforeMenuClicked(pVal.MenuUID);

                }
                else
                if (Application.Forms.ActiveForm.UniqueID == CurrentForm.UniqueID && !pVal.BeforeAction)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1282":
                            {
                                AddMenuClickAfter();
                            }
                            break;
                        case "1287":
                            {
                                DuplicateMenuClickAfter();
                            }
                            break;
                        case "1281":
                            {
                                FindMenuClickAfter();
                            }
                            break;
                    }
                }
                BubbleEvent = _Initializer.IsMenuResultClear;
            }
            catch (Exception ex)
            { ex.AppendInLogFile(); BubbleEvent = true; }
            _Initializer.IsMenuResultClear = true;
        }

        protected virtual void AddMenuClickAfter()
        {
        }


        protected virtual void DuplicateMenuClickAfter()
        {


        }
        protected virtual void FindMenuClickAfter()
        {

        }

        protected virtual bool BeforeMenuClicked(string p)
        {
            return true;
        }

        protected virtual bool BeforeCloseMenuClicked() { return true; }
        static bool isSaved = true;
        protected override void OnFormDataAddBefore(ref BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            foreach (var row in _LstAutoAddRow)
            {
                var matrix = GetItem(row.Key).Specific as SAPbouiCOM.Matrix;
                var tablename = matrix.Columns.Item(1).DataBind.TableName;
                matrix.FlushToDataSource();
                var ds = CurrentForm.DataSources.DBDataSources.Item(tablename);
                var alias = matrix.Columns.Item(row.Value).DataBind.Alias;

                ds.ClearAt(alias, "", "==");
                matrix.LoadFromDataSourceEx(false);
            }
            BubbleEvent = isSaved && IsMendatoryValidated();
            // base.OnFormDataAddBefore(ref pVal, out BubbleEvent);
        }

        protected override void OnFormDataUpdateBefore(ref BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = isSaved && IsMendatoryValidated();
            if (BubbleEvent)
                foreach (var row in _LstAutoAddRow)
                {
                    var matrix = GetItem(row.Key).Specific as SAPbouiCOM.Matrix;
                    var tablename = matrix.Columns.Item(1).DataBind.TableName;
                    matrix.FlushToDataSource();
                    var ds = CurrentForm.DataSources.DBDataSources.Item(tablename);
                    var alias = matrix.Columns.Item(row.Value).DataBind.Alias;

                    ds.ClearAt(alias, "", "==");
                    matrix.LoadFromDataSourceEx(false);
                }

            // base.OnFormDataUpdateBefore(ref pVal, out BubbleEvent);
        }


        protected virtual void OnMatrixRowAdded(string MatrixID, int rowid)
        {


        }
        protected override void OnFormDataLoadAfter(ref BusinessObjectInfo pVal)
        {
            EnforceViewModeOnCLose();

            base.OnFormDataLoadAfter(ref pVal);
        }
        protected override void OnFormDataAddAfter(ref BusinessObjectInfo pVal)
        {
            try
            {
                var key = getObjectKeyFromXML(pVal.ObjectKey);
                OnFormDataAdded(key);
            }
            catch (Exception ex)
            { ex.AppendInLogFile(); }
            base.OnFormDataAddAfter(ref pVal);
        }
        protected virtual void OnFormDataAdded(String ObjectKey)
        {
        }
        private Dictionary<string, dynamic> MendatoryFields = new Dictionary<string, dynamic>();
        private Dictionary<string, dynamic> MendatoryColumns = new Dictionary<string, dynamic>();

        protected void SetMendatoryField(string itemUID, string Message, bool ShowOnMessageBox = false)
        {
            MendatoryFields.Add(itemUID, new { Message, ShowOnMessageBox });
        }
        protected void SetMendatoryColumns(string GridID, string ColumnID, string Message, bool ShowOnMessageBox = false)
        {
            MendatoryColumns.Add(GridID + "." + ColumnID, new { Message, ShowOnMessageBox });
        }
        private bool IsMendatoryValidated()
        {

            bool yes = true;
            try
            {
                foreach (var item in MendatoryFields)
                {
                    yes = !string.IsNullOrEmpty((GetItem(item.Key).Specific as dynamic).Value);
                    if (!yes)
                    {
                        if (item.Value.ShowOnMessageBox)
                            Application.MessageBox(item.Value.Message);
                        else
                            Application.SetStatusBarMessage(item.Value.Message, BoMessageTime.bmt_Medium, true);
                        break;
                    }
                }
                if (yes)
                    foreach (var item in MendatoryColumns)
                    {
                        var items = item.Key.Split('.');
                        var MatrixName = items[0];
                        var ColumnsName = items[1];
                        if (GetItem(MatrixName).Specific is SAPbouiCOM.Matrix)
                        {
                            var cells = (GetItem(MatrixName).Specific as SAPbouiCOM.Matrix).Columns.Item(ColumnsName).Cells;
                            for (int i = 1; i <= cells.Count - 1; i++)
                            {
                                yes = !string.IsNullOrEmpty((cells.Item(i).Specific as dynamic).Value);
                                if (!yes)
                                {
                                    if (item.Value.ShowOnMessageBox)
                                        Application.MessageBox(item.Value.Message);
                                    else
                                        Application.SetStatusBarMessage(item.Value.Message, BoMessageTime.bmt_Medium, true);
                                    break;
                                }
                            }
                        }
                        if (!yes) break;
                    }
            }
            catch { }

            return yes;
        }

        public string GetQuery(string key, params object[] args)
        {


            var query = string.Format(GetQuery(key), args);

            Logger.Logger.Log(query);
            var LogPath = System.IO.Path.GetTempPath();
            Logger.Logger.CreateLog(LogPath, key + ".txt");
            Logger.Logger.ClearLog();
            return query;
        }
        public string GetQuery(string key)
        {
            string dbType = "SQL";
            switch (this.Company.DbServerType)
            {
                case SAPbobsCOM.BoDataServerTypes.dst_DB_2:
                    break;
                case SAPbobsCOM.BoDataServerTypes.dst_HANADB:
                    dbType = "HANA";
                    break;
                case SAPbobsCOM.BoDataServerTypes.dst_MAXDB:
                    break;
                case SAPbobsCOM.BoDataServerTypes.dst_MSSQL:
                case SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005:
                case SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008:
                case SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012:
                case SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014:
                    dbType = "SQL";
                    break;
                case SAPbobsCOM.BoDataServerTypes.dst_SYBASE:
                    break;
                default:
                    break;
            }
            var xmlPathBuilder = new StringBuilder("/Queries/Query[@name=\"{0}\"]/");
            if (!string.IsNullOrEmpty(dbType))
                if (dbType == "SQL")
                    xmlPathBuilder.Append(DatabaseTypes.SQL).ToString();
                else if (dbType == "HANA")
                    xmlPathBuilder.Append(DatabaseTypes.HANA).ToString();
                else
                    xmlPathBuilder.Append(DatabaseTypes.ORACLE).ToString();

            return GetXmlNodeValue(System.IO.Directory.GetCurrentDirectory() + "\\Queries\\Queries.xml", string.Format(xmlPathBuilder.ToString(), key));
        }
        string GetXmlNodeValue(string file, string xPath)
        {
            var doc = new XmlDocument();
            doc.Load(file);
            var xmlPath = string.Empty;
            var node = doc.DocumentElement.SelectSingleNode(xPath);
            return node.InnerText;
        }
    }
}
