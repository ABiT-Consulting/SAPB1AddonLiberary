using ADDONBASE.Extensions;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace ADDONBASE
{

    public abstract class BaseHandler : IDisposable
    {

        #region AutoMatrixRows
        private Dictionary<string, string> _LstAutoAddRow = new Dictionary<string, string>();
        protected void Add_AUTO_Matrix(string MatrixID, string ColUID)
        {
            try
            {
                _LstAutoAddRow.Add(MatrixID, ColUID);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        private bool _OnValidate(SAPbouiCOM.ItemEvent pVal)
        {
            bool yes = true;
            if (_LstAutoAddRow.ContainsKey(pVal.ItemUID))
            {
                try
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
                                catch (Exception ex) { ex.AppendInLogFile(); }
                                if (_value != 0.0)
                                {
                                    ds.InsertRecord(ds.Size);
                                    b = true;
                                }
                            }

                            matrix.LoadFromDataSourceEx();
                            if (b)
                                OnMatrixRowAdded(MatrixId, ds.Size);
                        }
                    }
                    yes = true;
                }
                catch (Exception ex) { }
            }
            return yes;
        }

        protected virtual void OnMatrixRowAdded(string MatrixID, int rowid)
        {

        }
        private bool _BeforeFormDataUpdated()
        {

            foreach (var row in _LstAutoAddRow)
            {
                var matrix = GetItem(row.Key).Specific as SAPbouiCOM.Matrix;
                var tablename = matrix.Columns.Item(1).DataBind.TableName;
                matrix.FlushToDataSource();
                var ds = CurrentForm.DataSources.DBDataSources.Item(tablename);
                var alias = matrix.Columns.Item(row.Value).DataBind.Alias;

                ds.ClearAt(alias, "", "==");
                matrix.LoadFromDataSourceEx();
            }

            return true;
        }
        private bool _BeforeFormDataAdded()
        {
            foreach (var row in _LstAutoAddRow)
            {
                var matrix = GetItem(row.Key).Specific as SAPbouiCOM.Matrix;
                var tablename = matrix.Columns.Item(1).DataBind.TableName;
                matrix.FlushToDataSource();
                var ds = CurrentForm.DataSources.DBDataSources.Item(tablename);
                var alias = matrix.Columns.Item(row.Value).DataBind.Alias;

                ds.ClearAt(alias, "", "==");
                matrix.LoadFromDataSourceEx();
            }

            return true;
        }


        #endregion
        public BaseHandler(String FormType)
        {
            //  this.Application.ItemEvent += Application_ItemEvent;


            this.FormType = FormType;
            if (timer != null)
            {
                timer.Dispose();
                timer = null;
            }
            //timer = new System.Timers.Timer(1000);
            //timer.Elapsed += BaseFormtimer_Elapsed;
            if (Timer != null)
            {
                Timer.Dispose();
                Timer = null;

            }
            Timer = new System.Timers.Timer(500);

        }
        public BaseHandler()
        {

            //this.Application.ItemEvent += BaseFormApplication_ItemEvent;
            //this.Application.LayoutKeyEvent += Application_LayoutKeyEvent;
            //if (timer != null)
            //{
            //    timer.Dispose();
            //    timer = null;
            //}
            //timer = new System.Timers.Timer(1000);
            //timer.Elapsed += BaseFormtimer_Elapsed;
        }
        protected bool ForcePLDEnabled = false;
        private string ___FORMUID;

        #region Attachment Part
        private bool ISATTACHMENT = false;

        private string MTX_AT = "";
        private string BRS_AT = "";
        private string DSP_AT = "";
        private string DEL_AT = "";
        string ATT_DBDataSourceName = "";
        public string ShowFileDialogBox()
        {
            string path = "";

            try
            {





                Thread t = new Thread((ThreadStart)delegate
                {




                    System.Windows.Forms.OpenFileDialog objDialog = new System.Windows.Forms.OpenFileDialog();

                    //FileDialog objDialog = new OpenFileDialog(); 

                    //  objDialog.Description = "Please select a Folder";

                    //    objDialog.InitialDirectory = Company.AttachMentPath;

                    //  objDialog.SelectedPath = @"C:\"; // or any other path you want to start...



                    // objDialog.ShowNewFolderButton = true;

                    //create a Form for ownership (SAP-Window won't work)

                    System.Windows.Forms.Form g = new System.Windows.Forms.Form();

                    g.Width = 200;

                    g.Height = 200;

                    g.Activate();

                    g.BringToFront();

                    g.Visible = false;

                    g.TopMost = true;

                    g.Focus();

                    System.Windows.Forms.DialogResult objResult = objDialog.ShowDialog(g);



                    Thread.Sleep(100);

                    if (objResult == System.Windows.Forms.DialogResult.OK)
                    {

                        path = objDialog.FileName;

                    }

                }

                )

                {

                    IsBackground = false,

                    Priority = ThreadPriority.Highest

                };

                t.SetApartmentState(ApartmentState.STA);

                t.Start();

                while (!t.IsAlive) ;

                Thread.Sleep(1);

                t.Join();

            }

            catch (Exception ex)
            {

                //Exception handling...

                //Funktionen.WriteError(ex);            

            }
            return path;
        }
        public string GetFolderPath()
        {

            string path = "";
            if (!Directory.Exists(Company.AttachMentPath))
                Application.SetStatusBarMessage("Attachment Path Does not exist");
            else
                try
                {





                    Thread t = new Thread((ThreadStart)delegate
                    {




                        System.Windows.Forms.OpenFileDialog objDialog = new System.Windows.Forms.OpenFileDialog();

                        //FileDialog objDialog = new OpenFileDialog(); 

                        //  objDialog.Description = "Please select a Folder";

                        //   objDialog.InitialDirectory = Company.AttachMentPath;

                        //  objDialog.SelectedPath = @"C:\"; // or any other path you want to start...



                        // objDialog.ShowNewFolderButton = true;

                        //create a Form for ownership (SAP-Window won't work)

                        System.Windows.Forms.Form g = new System.Windows.Forms.Form();

                        g.Width = 200;

                        g.Height = 200;

                        g.Activate();

                        g.BringToFront();

                        g.Visible = false;

                        g.TopMost = true;

                        g.Focus();

                        System.Windows.Forms.DialogResult objResult = objDialog.ShowDialog(g);



                        Thread.Sleep(100);

                        if (objResult == System.Windows.Forms.DialogResult.OK)
                        {

                            path = objDialog.FileName;

                        }

                    }

                    )

                    {

                        IsBackground = false,

                        Priority = ThreadPriority.Highest

                    };

                    t.SetApartmentState(ApartmentState.STA);

                    t.Start();

                    while (!t.IsAlive) ;

                    Thread.Sleep(1);

                    t.Join();

                }

                catch (Exception ex)
                {

                    //Exception handling...

                    //Funktionen.WriteError(ex);            

                }





            return path;





        }
        private SAPbouiCOM.Matrix AttachmentMatrix
        {
            get
            {
                return CurrentForm.Items.Item(MTX_AT).Specific as SAPbouiCOM.Matrix;

            }
        }
        private SAPbouiCOM.Button BTN_BRS_AT
        {
            get
            {
                return CurrentForm.Items.Item(BRS_AT).Specific as SAPbouiCOM.Button;

            }
        }
        private SAPbouiCOM.Button BTN_DSP_AT
        {
            get
            {
                return CurrentForm.Items.Item(DSP_AT).Specific as SAPbouiCOM.Button;
            }
        }
        private SAPbouiCOM.Button BTN_DEL_AT
        {
            get
            {
                return CurrentForm.Items.Item(DEL_AT).Specific as SAPbouiCOM.Button;
            }
        }
        private DBDataSource DBDSATTCHMENT
        {
            get
            {
                return CurrentForm.DataSources.DBDataSources.Item(ATT_DBDataSourceName);
            }
        }
        const string U_Srcpth = "U_Srcpth";
        const string U_Trgetpth = "U_Trgetpth";
        const string U_FileName = "U_FileName";
        private void OnBrowsButtonCLick_ATT(SAPbouiCOM.ItemEvent pVal)
        {
            if (ISATTACHMENT && pVal.Action_Success && !pVal.Before_Action)
                try
                {

                    if (pVal.ItemUID == BRS_AT)
                    {
                        CurrentForm.Freeze(true);
                        AttachmentMatrix.FlushToDataSource();

                        try
                        {
                            var path = GetFolderPath();
                            if (!string.IsNullOrEmpty(path))
                            {
                                if (DBDSATTCHMENT.Size == 0) DBDSATTCHMENT.InsertRecord(DBDSATTCHMENT.Size);
                                if (!string.IsNullOrEmpty(DBDSATTCHMENT.GetValue(U_Srcpth, DBDSATTCHMENT.Size - 1).Trim()))
                                    DBDSATTCHMENT.InsertRecord(DBDSATTCHMENT.Size);
                                var filename = Path.GetFileName(path);
                                var dateString = Company.GetCompanyDate().ToString("yyyyMMdd");
                                var ext = Path.GetExtension(filename);
                                DBDSATTCHMENT.SetValue(U_Srcpth, DBDSATTCHMENT.Size - 1, path);
                                DBDSATTCHMENT.SetValue(U_Trgetpth, DBDSATTCHMENT.Size - 1, Company.AttachMentPath);
                                DBDSATTCHMENT.SetValue(U_FileName, DBDSATTCHMENT.Size - 1, FormType.ToString() + "_" + Company.GetCompanyDate().ToString("yyyyMMdd") + "_" + Company.UserSignature.ToString() + "_" + Company.GetCompanyTime().Replace(":", "") + "_" + DBDSATTCHMENT.Size.ToString() + ext);
                                DBDSATTCHMENT.SetValue("U_Date", DBDSATTCHMENT.Size - 1, dateString);
                                if (this.CurrentForm.Mode != BoFormMode.fm_ADD_MODE && this.CurrentForm.Mode != BoFormMode.fm_FIND_MODE)
                                    this.CurrentForm.Mode = BoFormMode.fm_UPDATE_MODE;
                            }
                            // DBDSATTCHMENT.InsertRecord(DBDSATTCHMENT.Size);
                        }
                        catch (Exception ex)
                        {
                            if (System.Diagnostics.Debugger.IsAttached)
                                ex.printAtMessageBox();
                            ex.StackTrace.PrintString();
                        }

                        AttachmentMatrix.LoadFromDataSourceEx();

                        CurrentForm.Freeze(false);
                    }
                    else
                        if (pVal.ItemUID == DSP_AT)
                    {
                        var ind = AttachmentMatrix.GetSelectedRow();
                        if (ind != -1)
                            System.Diagnostics.Process.Start(DBDSATTCHMENT.GetValue(U_Srcpth, ind));
                    }
                    else if (pVal.ItemUID == DEL_AT)
                    {
                        var ind = AttachmentMatrix.GetSelectedRow();
                        if (ind != -1)
                        {
                            AttachmentMatrix.FlushToDataSource();

                            if (this.CurrentForm.Mode != BoFormMode.fm_ADD_MODE && this.CurrentForm.Mode != BoFormMode.fm_FIND_MODE)
                                this.CurrentForm.Mode = BoFormMode.fm_UPDATE_MODE;
                            if (Path.Combine(DBDSATTCHMENT.GetValue(U_Trgetpth, ind).Trim(), DBDSATTCHMENT.GetValue(U_FileName, ind).Trim()) == DBDSATTCHMENT.GetValue(U_Srcpth, ind).Trim())
                            {
                                if (File.Exists(DBDSATTCHMENT.GetValue(U_Srcpth, 0).Trim()))
                                    File.Delete(DBDSATTCHMENT.GetValue(U_Srcpth, 0).Trim());
                            }
                            DBDSATTCHMENT.RemoveRecord(ind);
                            AttachmentMatrix.LoadFromDataSourceEx();
                        }
                    }
                }
                catch (Exception ex) { ex.AppendInLogFile(); }
        }
        private void Attachment_SAVE()
        {
            if (ISATTACHMENT)
            {

                AttachmentMatrix.FlushToDataSource();
                try
                {
                    var size = DBDSATTCHMENT.Size;
                    for (int i = 0; i < size; i++)
                    {
                        var srcPath = DBDSATTCHMENT.GetValue(U_Srcpth, i).Trim();
                        var targetPath = Path.Combine(DBDSATTCHMENT.GetValue(U_Trgetpth, i).Trim(), DBDSATTCHMENT.GetValue(U_FileName, i).Trim());

                        if (srcPath != targetPath)
                        {
                            if (Directory.Exists(DBDSATTCHMENT.GetValue(U_Trgetpth, i).Trim()))
                            {
                                Directory.CreateDirectory(DBDSATTCHMENT.GetValue(U_Trgetpth, i).Trim());
                            }
                            if (File.Exists(srcPath))
                            {

                                File.Copy(srcPath, targetPath, true);
                                DBDSATTCHMENT.SetValue(U_Srcpth, i, targetPath);
                            }



                        }
                    }
                }
                catch (Exception ex) { ex.AppendInLogFile(); }
                AttachmentMatrix.LoadFromDataSourceEx();
            }
        }
        protected void SetAttachmentFolder(string Matrix, string BrowseButton, string DisplayButton, string DeleteButton, string DBDataSourceName)
        {
            if (CurrentForm == null)
            {
                if (System.Diagnostics.Debugger.IsAttached)
                {
                    Application.SetStatusBarMessage("This Function should be called in FormLoad()");
                    return;
                }
            }
            this.MTX_AT = Matrix;
            this.BRS_AT = BrowseButton;
            this.DSP_AT = DisplayButton;
            this.DEL_AT = DeleteButton;
            this.ATT_DBDataSourceName = DBDataSourceName;

            if (AttachmentMatrix == null || BTN_BRS_AT == null || BTN_DSP_AT == null || BTN_DEL_AT == null || DBDSATTCHMENT == null)
            {
                if (AttachmentMatrix == null) Application.SetStatusBarMessage(string.Format("{0} does not exist", this.MTX_AT));
                if (BTN_BRS_AT == null) Application.SetStatusBarMessage(string.Format("{0} does not exist", this.BRS_AT));
                if (BTN_DSP_AT == null) Application.SetStatusBarMessage(string.Format("{0} does not exist", this.DSP_AT));
                if (BTN_DEL_AT == null) Application.SetStatusBarMessage(string.Format("{0} does not exist", this.DEL_AT));
                if (DBDSATTCHMENT == null) Application.SetStatusBarMessage(string.Format("{0} does not exist", this.ATT_DBDataSourceName));
            }
            else
            {
                ISATTACHMENT = true;
                BTN_BRS_AT.Item.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, BoModeVisualBehavior.mvb_True);
                BTN_DSP_AT.Item.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, BoModeVisualBehavior.mvb_True);
                BTN_DEL_AT.Item.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, BoModeVisualBehavior.mvb_True);
                this.AttachmentMatrix.Columns.Item(U_Srcpth).Visible = false;
            }
        }

        #endregion



        BoFormMode _LastMode;
        Int32 IsOnMatrix = 0, IsRow = 0, IsCol = 0;
        System.Timers.Timer Timer;
        private string _FormUID
        {
            get
            {
                return this._FORMUID;
            }
        }
        private string _itemUID;
        private String _FormType = "";
        string matrixid = "";
        static bool isSaved = false;
        private System.Timers.Timer timer;
        private Dictionary<string, List<dynamic>> _CFLArray = new Dictionary<string, List<dynamic>>();

        private List<string> _ItemToClose = new List<string>();
        private Dictionary<string, string> MendatoryFields = new Dictionary<string, string>();
        private static List<String> FormTypes = new List<string>();

        protected SAPbouiCOM.Application Application
        { get { return Initializer._Application; } }
        protected SAPbobsCOM.Company Company
        { get { return Initializer._Company; } }
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
        protected DBDataSource GetDBDataSource(object o)
        {
            return CurrentForm.DataSources.DBDataSources.Item(o);

        }

        internal string _FORMUID
        {
            get { return ___FORMUID; }
            set
            {
                ___FORMUID = value;
                if (FormType == Application.Forms.Item(value).TypeEx)
                {
                    //_FormUID = (String.IsNullOrEmpty(_FORMUID)) ? Application.Forms.ActiveForm.UniqueID : _FORMUID ;


                    //this.Application.ItemEvent += BaseFormApplication_ItemEvent;
                    //  this.Application.RightClickEvent += Application_RightClickEvent;
                    //this.Application.LayoutKeyEvent += Application_LayoutKeyEvent;
                    //   this.Application.MenuEvent += BaseFormApplication_MenuEvent;
                    try
                    {
                        OnFormLoaded(value);
                    }
                    catch (Exception ex) { }
                }
            }
        }
        protected SAPbouiCOM.Form CurrentForm
        {
            get
            {

                if (___FORMUID == null) return null;
                // if (___FORMUID != Application.Forms.ActiveForm.UniqueID) return null;
                return Application.Forms.Item(___FORMUID);


            }
        }
        protected SAPbouiCOM.Item CurrentItem
        {
            get
            {
                return CurrentForm.Items.Item(_itemUID);


            }
        }
        protected SAPbouiCOM.EditText CurrentEditText
        {
            get
            {
                return CurrentForm.Items.Item(_itemUID).Specific as SAPbouiCOM.EditText;
            }
        }
        protected bool EnableAddLineMenu
        {
            set { CurrentForm.EnableMenu("1292", value); }
        }
        protected bool EnableRemoveLineMenu
        {
            set { CurrentForm.EnableMenu("1293", value); }
        }
        protected string FormType
        {

            get { return _FormType; }
            set
            {
                Add(value);
                _FormType = value;
            }
        }

        //private bool OnEveryFormActivated()
        //{
        //    this.Application.FormDataEvent += BaseFormApplication_FormDataEvent;

        //   // this.Application.MenuEvent += Application_MenuEvent; 
        //    return true;
        //}
        //private bool OnEveryFormDeActivated()
        //{
        //    this.Application.FormDataEvent -= BaseFormApplication_FormDataEvent;

        //  //  this.Application.MenuEvent -= Application_MenuEvent; 
        //    return true;
        //}

        private bool _OnCFLLoad(SAPbouiCOM.ChooseFromListEvent chooseFromListEvent)
        {
            try
            {
                if (!string.IsNullOrEmpty(chooseFromListEvent.ColUID))
                {
                    var matrix = GetItem(chooseFromListEvent.ItemUID).Specific as SAPbouiCOM.Matrix;
                    var ColumnIndex = matrix.GetCellFocus().ColumnIndex;
                    matrixid = chooseFromListEvent.ItemUID;
                    IsOnMatrix = 1;
                    IsRow = chooseFromListEvent.Row;
                    IsCol = ColumnIndex;
                }
            }
            catch (Exception ex) { ex.AppendInLogFile(); }

            return true;
        }
        private bool _OnFormActivated(SAPbouiCOM.ItemEvent pVal)
        {
            if (IsOnMatrix == 1)
            {
                (GetItem(matrixid).Specific as SAPbouiCOM.Matrix).SetCellFocus(IsRow, IsCol);
                IsOnMatrix = 0;
            }
            return true;
        }
        private bool _OnLostFocus(SAPbouiCOM.ItemEvent pVal)
        {
            if (IsOnMatrix == 1)
            {
                (GetItem(matrixid).Specific as SAPbouiCOM.Matrix).SetCellFocus(IsRow, IsCol);
                IsOnMatrix = 0;
            }
            return true;
        }

        protected void SetMendatoryField(string itemUID, string Message)
        {
            MendatoryFields.Add(itemUID, Message);
        }
        protected SAPbouiCOM.Item GetItem(object uid)
        {
            return CurrentForm.Items.Item(uid);
        }
        protected SAPbouiCOM.EditText GetEditText(object uid)
        {
            return (SAPbouiCOM.EditText)CurrentForm.Items.Item(uid).Specific;
        }

        //public BaseHandler()
        //{
        //    this.Application.ItemEvent += Application_ItemEvent;
        //    this.Application.FormDataEvent += Application_FormDataEvent;
        //    this.Application.MenuEvent += Application_MenuEvent;

        //}


        //void Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        //{
        //    bool succes = true;
        //    if (FormUID == ___FORMUID && pVal.EventType == BoEventTypes.et_FORM_ACTIVATE && !pVal.BeforeAction && pVal.Action_Success)
        //    {

        //        this.Application.ItemEvent += BaseFormApplication_ItemEvent;
        //        this.Application.RightClickEvent += Application_RightClickEvent;
        //        this.Application.LayoutKeyEvent += Application_LayoutKeyEvent;
        //        this.Application.MenuEvent += BaseFormApplication_MenuEvent;
        //    }
        //    else
        //        if (FormUID == ___FORMUID && pVal.EventType == BoEventTypes.et_FORM_DEACTIVATE && !pVal.BeforeAction && pVal .Action_Success )
        //        {

        //            this.Application.ItemEvent -= BaseFormApplication_ItemEvent;
        //            this.Application.RightClickEvent -= Application_RightClickEvent;
        //            this.Application.LayoutKeyEvent -= Application_LayoutKeyEvent;
        //            this.Application.MenuEvent -= BaseFormApplication_MenuEvent;
        //        }
        //    BubbleEvent =  succes;
        //}


        protected virtual bool BeforeRightClick(ContextMenuInfo eventInfo)
        {
            return true;
        }

        protected virtual bool onRightClick(ContextMenuInfo eventInfo)
        {
            return true;
        }



        private bool _ForceDefaultButtonToOK = true;
        protected bool ForceDefaultButtonToOK
        {
            set
            {
                _ForceDefaultButtonToOK = value;
            }
            get
            {
                return _ForceDefaultButtonToOK;
            }
        }
        protected bool ForceViewModeonClose
        {
            get;
            set;
        }
        protected string getObjectKeyFromXML(String XML)
        {
            return Extensions.Extensions.getObjectKeyFromXML(XML);
        }
        void clickAddMenue(string FormUID)
        {
            //while (  CurrentForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)
            //   o.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

        }
        protected void AddRowToMatrix(object MatrixID)
        {
            var Matrix = (GetItem(MatrixID).Specific as SAPbouiCOM.Matrix);

            try
            {
                if (Matrix.RowCount == 0)
                {
                    Matrix.AddRow();
                    Matrix.FlushToDataSource();
                }
            }
            catch (Exception ex) { ex.AppendInLogFile(); }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(Matrix);
        }
        internal void OnFormLoaded(string FormUID)
        {
            var o = Application.Forms.Item(FormUID);
            if (o.Mode == BoFormMode.fm_VIEW_MODE && !o.IsSystem)
            {
                FormUID.PrintString();
                o.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                Task tsk = new Task(() =>
                {

                    while (!Application.Menus.Item("1282").Enabled) ;
                    //if(o.Mode ==  BoFormMode.fm_FIND_MODE )
                    // Application .Forms .Item(FormUID).Menu.Item("1282").Activate();
                    if (o.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                        Application.ActivateMenuItem("1282");

                    //   System.Runtime.InteropServices.Marshal.ReleaseComObject(o);   
                });
                tsk.Start();
                //  clickAddMenue(o);
                // CurrentForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
            }
            if (ForcePLDEnabled)
                try
                {
                    PLDAdder.PLDAdder pldadder = new PLDAdder.PLDAdder();
                    CurrentForm.ReportType = pldadder.getReportTypeCode(CurrentForm.Title, CurrentForm.BusinessObject.Type, Initializer.ADDON_NAME, CurrentForm.BusinessObject.Type);
                }
                catch (Exception ex) { ex.AppendInLogFile(); }
            //            try
            //            {
            //                if(!CurrentForm.IsSystem)
            //                {
            //                    System.Drawing.Rectangle Rect;

            //                    Rect = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea ;


            //CurrentForm.Left = Rect.Width - 10;

            //CurrentForm.Top = Rect.Top - 10;

            //                }
            //            }
            //            catch
            //            { }
            OnFormLoad(o);
            try
            {
                Application.ActivateMenuItem("1297");
            }
            catch (Exception ex) { ex.AppendInLogFile(); }
            EnforceViewModeOnCLose();
            if (HandleFormModeChanged)
            {
                _LastMode = CurrentForm.Mode;
                Timer.Elapsed -= Timer_Elapsed;
                Timer.Elapsed += Timer_Elapsed;
                Timer.Start();

            }
        }
        protected bool HandleFormModeChanged = true;
        void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            Task tsk = new Task(new Action(() =>
            {
                Timer.Stop();
                try
                {
                    if (Application.Forms.ActiveForm.UniqueID == _FormUID)
                    {
                        if (_LastMode != CurrentForm.Mode)
                        {
                            ModeChanged(CurrentForm.Mode, _LastMode);

                            _LastMode = CurrentForm.Mode;
                        }
                    }
                }
                catch (Exception ex) { ex.AppendInLogFile(); }
                Timer.Start();

            }));
            tsk.Start();
        }

        protected virtual void ModeChanged(BoFormMode boFormMode, BoFormMode _LastMode)
        {
        }
        protected void SetEditableForAddModeOnly(String ItemUID)
        {
            try
            {
                GetItem(ItemUID).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                GetItem(ItemUID).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }
            catch (Exception ex) { ex.AppendInLogFile(); }
        }
        protected void SetVisibleForAddModeOnly(String ItemUID)
        {
            GetItem(ItemUID).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            GetItem(ItemUID).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, (int)SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True);

        }
        virtual protected void OnMenuClicked(String MenuUID) { }
        virtual protected bool before_et_FORM_DATA_LOAD(SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo)
        {
            return true;
        }
        virtual protected bool OnFormDataAdded(string ObjectKey)
        {
            return true;
        }


        virtual protected bool OnFormDataUpdated(string key)
        {
            return true;
        }
        virtual protected bool OnValidate(SAPbouiCOM.ItemEvent pval)
        {
            return true;
        }
        virtual protected bool onDefaultButtonClicked(SAPbouiCOM.ItemEvent pval)
        {
            return true;
        }
        protected virtual bool onFormClosed(SAPbouiCOM.ItemEvent pVal) { return true; }
        protected virtual void OnFormLoad(Object o) { }
        protected virtual bool OnFormDataLoaded() { return true; }
        protected virtual bool onet_GOT_FOCUS(ItemEvent pVal) { return true; }
        protected virtual bool onITEM_PRESSED(ItemEvent pval) { return true; }
        protected virtual bool OnFormActivated(ItemEvent pVal) { return true; }
        protected virtual bool onComboSelected(ItemEvent pVal) { return true; }
        protected virtual bool OnCFLLoad(ChooseFromListEvent chooseFromListEvent)
        {
            return true;
        }
        protected virtual bool ItemEvent(ItemEvent pVal) { return true; }
        protected virtual bool OnLostFocus(ItemEvent pVal) { return true; }
        protected virtual bool OnResize(ItemEvent pVal) { return true; }
        /// <summary>
        /// When form is fully loaded and visible
        /// </summary>
        /// <param name="o">current instance of the form</param>

        protected virtual bool BeforeCloseMenuClicked() { return true; }
        protected virtual bool BeforeAddMenuClicked() { return true; }
        protected virtual bool BeforeFindMenuClicked() { return true; }
        protected virtual bool Beforeet_CHOOSE_FROM_LIST(SAPbouiCOM.ChooseFromListEvent cflEventArgs)
        {

            return true;
        }
        protected virtual bool BeforeDefaultButtonClick(SAPbouiCOM.ItemEvent pval)
        {
            return true;
        }
        protected virtual bool BeforeFormDataAddedORUpdated(BoEventTypes boEventTypes)
        {
            return true;

        }
        protected virtual bool BeforeFormDataAdded()
        {
            return true;
        }
        protected virtual bool BeforeFormDataUpdated()
        {
            return true;
        }
        protected void OpenLinkedObject(BoLinkedObject type)
        {
            SAPbouiCOM.Item ToItm;
            SAPbouiCOM.Item lnkToitm;
            try
            {

                ToItm = CurrentForm.Items.Add("U_TOITM", BoFormItemTypes.it_EDIT);
                lnkToitm = CurrentForm.Items.Add("lnkToitm", BoFormItemTypes.it_LINKED_BUTTON);
                ToItm.Width = 1;

            }
            catch
            {
                ToItm = CurrentForm.Items.Item("U_TOITM");
                lnkToitm = CurrentForm.Items.Item("lnkToitm");

            }

            ToItm.Visible = true;
            lnkToitm.Visible = true;
            ToItm.Enabled = true;
            ToItm.Left = CurrentForm.Width - 100;
            lnkToitm.Left = CurrentForm.Width - 100;
            lnkToitm.LinkTo = ToItm.UniqueID;
            (lnkToitm.Specific as SAPbouiCOM.LinkedButton).LinkedObject = type;// BoLinkedObject.lf_GoodsIssue;
            if (type == BoLinkedObject.lf_GoodsIssue)
            {

                string query = string.Format("Select \"DocEntry\" from \"OIGE\" where  \"U_basetype\" =  '{0}' and \"U_basentry\" = '{1}'", CurrentForm.BusinessObject.Type, MasterDS.GetValue("DocEntry", 0).Trim());
                var recset = Company.DoQuery(query);
                (ToItm.Specific as SAPbouiCOM.EditText).Value = recset.Fields.Item(0).Value.ToString();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(recset); GC.Collect();

            }
            else
                if (type == BoLinkedObject.lf_Invoice)
            {

                string query = string.Format("Select \"DocEntry\" from \"oinv\" where  \"U_basetype\" =  '{0}' and \"U_basentry\" = '{1}'", CurrentForm.BusinessObject.Type, MasterDS.GetValue("DocEntry", 0).Trim());
                var recset = Company.DoQuery(query);
                (ToItm.Specific as SAPbouiCOM.EditText).Value = recset.Fields.Item(0).Value.ToString();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(recset); GC.Collect();

            }
            else
                    if (type == BoLinkedObject.lf_PurchaseInvoice)
            {

                string query = string.Format("Select \"DocEntry\" from \"opch\" where  \"U_basetype\" =  '{0}' and \"U_basentry\" = '{1}'", CurrentForm.BusinessObject.Type, MasterDS.GetValue("DocEntry", 0).Trim());
                var recset = Company.DoQuery(query);
                (ToItm.Specific as SAPbouiCOM.EditText).Value = recset.Fields.Item(0).Value.ToString();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(recset); GC.Collect();

            }
            lnkToitm.Click();
            ToItm.Enabled = false;
            ToItm.Visible = false;
            lnkToitm.Visible = false;
        }

        protected void OpenLinkedObject(string type)
        {
            SAPbouiCOM.Item ToItm;
            SAPbouiCOM.Item lnkToitm;
            try
            {

                ToItm = CurrentForm.Items.Add("U_TOITM", BoFormItemTypes.it_EDIT);
                lnkToitm = CurrentForm.Items.Add("lnkToitm", BoFormItemTypes.it_LINKED_BUTTON);
                ToItm.Width = 1;

            }
            catch
            {
                ToItm = CurrentForm.Items.Item("U_TOITM");
                lnkToitm = CurrentForm.Items.Item("lnkToitm");

            }

            ToItm.Visible = true;
            lnkToitm.Visible = true;
            ToItm.Enabled = true;
            ToItm.Left = CurrentForm.Width - 100;
            lnkToitm.Left = CurrentForm.Width - 100;
            lnkToitm.LinkTo = ToItm.UniqueID;
            (lnkToitm.Specific as SAPbouiCOM.LinkedButton).LinkedObjectType = "u" + type;// BoLinkedObject.lf_GoodsIssue;


            string query = string.Format("Select \"DocEntry\" from \"@{0}\" where  \"U_basetype\" =  '{1}' and \"U_basentry\" = '{2}'", type, CurrentForm.BusinessObject.Type, MasterDS.GetValue("DocEntry", 0).Trim());
            var recset = Company.DoQuery(query);
            (ToItm.Specific as SAPbouiCOM.EditText).Value = recset.Fields.Item(0).Value.ToString();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(recset); GC.Collect();

            //Application.ActivateMenuItem("3079");
            //Application.ActivateMenuItem("1281");
            //var form = Application.Forms.ActiveForm;
            //var EditText = form.Items.Item("7").Specific as SAPbouiCOM.EditText;
            //EditText.Value = Company.DoQuery("Select DocEntry from \"OIGE\" where  U_basetype =  '{0}' and U_basentry = '{1}'", CurrentForm.BusinessObject.Key, MasterDS.GetValue("DocEntry", 0).Trim()).Fields.Item(0).Value.ToString();
            //form.Items.Item("1").Click();

            lnkToitm.Click();
            ToItm.Enabled = false;
            ToItm.Visible = false;
            lnkToitm.Visible = false;
        }
        protected virtual bool Beforeet_GOT_FOCUS(ItemEvent pVal) { return true; }
        protected virtual bool BeforeLostFocus(ItemEvent pVal) { return true; }
        void EnforceViewModeOnCLose()
        {
            try
            {
                if (_ItemToClose.Count > 0 && MasterDS.GetValue("Status", 0).Trim().ToLower() == "c")
                {
                    foreach (var item in _ItemToClose)
                    {
                        GetItem(item).Enabled = false;
                    }
                }
                else
                {

                    foreach (var item in _ItemToClose)
                    {
                        GetItem(item).Enabled = true;
                    }

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
        internal void Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {

            var success = true;
            try
            {
                if (Application.Forms.ActiveForm.UniqueID == _FormUID)
                {

                    if (pVal.BeforeAction && this.Application.Forms.ActiveForm.TypeEx == FormType && pVal.MenuUID == "1286")
                    {
                        success = BeforeCloseMenuClicked();
                    }
                    else
                        if (pVal.BeforeAction && this.Application.Forms.ActiveForm.TypeEx == FormType && pVal.MenuUID == "1282")
                    {
                        success = BeforeAddMenuClicked();
                    }
                    else
                            if (pVal.BeforeAction && this.Application.Forms.ActiveForm.TypeEx == FormType && pVal.MenuUID == "1281")
                    {
                        success = BeforeFindMenuClicked();
                    }
                    if (pVal.BeforeAction && this.Application.Forms.ActiveForm.TypeEx == FormType)
                    {
                        success = success && BeforeMenuClicked(pVal.MenuUID);
                    }
                    if (!pVal.BeforeAction)
                    {

                        OnMenuClicked(pVal.MenuUID);
                    }

                }
            }
            catch (Exception ex) { ex.AppendInLogFile(); }
            BubbleEvent = success;
        }

        protected virtual bool BeforeMenuClicked(string Menuid)
        {
            return true;
        }
        protected virtual bool application_FormDataEvent(SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo) { return true; }
        internal void Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            try
            {
                // Console.WriteLine("Form Type: {0} ,EventType = {1}, BeforeAction = {2},Action Success={3},FormTypeEx ={4},FormUID ={5},ObjectKey ={6},Type = {7}", BusinessObjectInfo.FormTypeEx, BusinessObjectInfo.EventType, BusinessObjectInfo.BeforeAction, BusinessObjectInfo.ActionSuccess, 
                //     BusinessObjectInfo.FormTypeEx, BusinessObjectInfo.FormUID,BusinessObjectInfo .ObjectKey ,BusinessObjectInfo .Type );
                //  if (!FormTypes.Contains(BusinessObjectInfo.FormTypeEx)) { BubbleEvent = true; return; }


                if (BusinessObjectInfo.FormTypeEx == FormType)
                {
                    application_FormDataEvent(BusinessObjectInfo);

                    //_FormUID = (String.IsNullOrEmpty(_FORMUID)) ? BusinessObjectInfo.FormUID : _FORMUID;
                    BaseHandler.isSaved = true;

                    #region On Action Success
                    if (BusinessObjectInfo.ActionSuccess && !BusinessObjectInfo.BeforeAction)
                    {
                        switch (BusinessObjectInfo.EventType)
                        {
                            case BoEventTypes.et_FORM_DATA_ADD:
                                try
                                {
                                    var value = getObjectKeyFromXML(BusinessObjectInfo.ObjectKey);

                                    BaseHandler.isSaved = OnFormDataAdded(value);
                                }
                                catch (Exception ex) { ex.AppendInLogFile(); }
                                break;
                            case BoEventTypes.et_FORM_DATA_LOAD:
                                //if (ForceViewModeonClose)
                                //{
                                //    if (MasterDS.GetValue("Status", 0).Trim() == "C" && CurrentForm.Mode != BoFormMode.fm_VIEW_MODE)
                                //    {

                                //        CurrentForm.Mode = BoFormMode.fm_VIEW_MODE;
                                //    }
                                //    if (MasterDS.GetValue("Status", 0).Trim() == "O" && CurrentForm.Mode == BoFormMode.fm_VIEW_MODE)
                                //    {
                                //        CurrentForm.Mode = BoFormMode.fm_OK_MODE;
                                //    }
                                //}
                                EnforceViewModeOnCLose();

                                BaseHandler.isSaved = OnFormDataLoaded();
                                break;
                            case BoEventTypes.et_FORM_DATA_UPDATE:
                                var key = getObjectKeyFromXML(BusinessObjectInfo.ObjectKey);

                                BaseHandler.isSaved = OnFormDataUpdated(key);

                                try
                                {
                                    if (MasterDS.GetValue("Status", 0).Trim() == "C" && MasterDS.GetValue("Canceled", 0).Trim() == "N")
                                    {
                                        bool succ = onFormDataClose(key);
                                        if (!succ)
                                        {

                                        }
                                        BaseHandler.isSaved = BaseHandler.isSaved && succ;
                                    }
                                }
                                catch (Exception)
                                {

                                }
                                break;
                        }

                    }
                    #endregion
                    else
                    #region Before Action
                        if (!BusinessObjectInfo.ActionSuccess && BusinessObjectInfo.BeforeAction)
                    {

                        switch (BusinessObjectInfo.EventType)
                        {
                            case BoEventTypes.et_FORM_DATA_ADD:
                                BaseHandler.isSaved = ValidationCheck() && IsMendatoryValidated() && BaseHandler.isSaved && BeforeFormDataAdded() && BeforeFormDataAddedORUpdated(BusinessObjectInfo.EventType) && _BeforeFormDataAdded();

                                break;

                            case BoEventTypes.et_FORM_DATA_LOAD:
                                BaseHandler.isSaved = before_et_FORM_DATA_LOAD(BusinessObjectInfo);
                                break;
                            case BoEventTypes.et_FORM_DATA_UPDATE:
                                BaseHandler.isSaved = ValidationCheck()
                                    && IsMendatoryValidated()
                                    && BeforeFormDataUpdated() && BeforeFormDataAddedORUpdated(BusinessObjectInfo.EventType);
                                if (BaseHandler.isSaved)
                                {
                                    BaseHandler.isSaved = _BeforeFormDataUpdated();
                                }
                                break;
                        }
                        if (BaseHandler.isSaved && (BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_ADD || BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_UPDATE))
                            Attachment_SAVE();
                    }
                    #endregion
                    else
                    #region On Action Failed
                            if (!BusinessObjectInfo.ActionSuccess && !BusinessObjectInfo.BeforeAction)
                    {
                        switch (BusinessObjectInfo.EventType)
                        {
                            case BoEventTypes.et_FORM_DATA_ADD:
                                break;
                            case BoEventTypes.et_FORM_DATA_LOAD:
                                break;
                            case BoEventTypes.et_FORM_DATA_UPDATE:
                                break;
                        }
                    }
                    #endregion


                }
            }
            catch (Exception ex)
            {
                if (System.Diagnostics.Debugger.IsAttached)
                {
                    Application.SetStatusBarMessage("This Error Appears only in Debug Mode:" + ex.Message, BoMessageTime.bmt_Short, true);
                }

            }

            BubbleEvent = BaseHandler.isSaved;
        }
        protected virtual bool OnItemChanged(SAPbouiCOM.ItemEvent pVal)
        {
            return true;
        }
        internal void Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            var success = true;
            try
            {
                //if (!FormTypes.Contains(pVal.FormTypeEx)) 
                //{ BubbleEvent = true; return; }
                //   this.Application.ItemEvent -= BaseFormApplication_ItemEvent;
                if (pVal.FormTypeEx == FormType)
                {

                    //_FormUID = (String.IsNullOrEmpty(_FORMUID)) ? pVal.FormUID : _FORMUID;
                    _itemUID = pVal.ItemUID;
                    CurrentForm.Freeze(true);
                    try
                    {

                        #region After Action
                        if (pVal.Action_Success && !pVal.Before_Action)
                        {
                            try
                            {
                                if (pVal.ItemChanged)
                                {
                                    //ApplyCFLName(pVal);

                                    OnItemChanged(pVal);
                                }
                            }
                            catch (Exception ex)
                            {
                                if (System.Diagnostics.Debugger.IsAttached)
                                    Console.WriteLine(ex.Message + " " + ex.StackTrace);
                            }
                            EnforceViewModeOnCLose();
                            switch (pVal.EventType)
                            {
                                case BoEventTypes.et_ALL_EVENTS:

                                    break;
                                case BoEventTypes.et_B1I_SERVICE_COMPLETE:
                                    break;
                                case BoEventTypes.et_CHOOSE_FROM_LIST:
                                    ApplyCFLName(pVal);
                                    //success = _OnLostFocus(pVal);
                                    success = OnCFLLoad((SAPbouiCOM.ChooseFromListEvent)pVal);

                                    break;
                                case BoEventTypes.et_CLICK:
                                    if (pVal.ItemUID == CurrentForm.DefButton)
                                    {
                                        success = onDefaultButtonClicked(pVal);
                                    }
                                    success = onButtonClicked(pVal);
                                    success = onClicked(pVal);
                                    break;
                                case BoEventTypes.et_COMBO_SELECT:

                                    success = onComboSelected(pVal);
                                    break;
                                case BoEventTypes.et_DATASOURCE_LOAD:
                                    break;
                                case BoEventTypes.et_DOUBLE_CLICK:
                                    success = success && on_et_DOUBLE_CLICK(pVal);
                                    break;
                                case BoEventTypes.et_Drag:
                                    break;
                                case BoEventTypes.et_EDIT_REPORT:
                                    break;
                                case BoEventTypes.et_FORMAT_SEARCH_COMPLETED:
                                    break;
                                case BoEventTypes.et_FORM_ACTIVATE:

                                    success = OnFormActivated(pVal) && _OnFormActivated(pVal);

                                    break;
                                case BoEventTypes.et_FORM_CLOSE:
                                    if (HandleFormModeChanged)
                                        Timer.Elapsed -= Timer_Elapsed;
                                    success = onFormClosed(pVal);

                                    break;
                                case BoEventTypes.et_FORM_DATA_ADD:
                                    isSaved = true;
                                    break;
                                case BoEventTypes.et_FORM_DATA_DELETE:
                                    break;
                                case BoEventTypes.et_FORM_DATA_LOAD:

                                    isSaved = true;
                                    break;
                                case BoEventTypes.et_FORM_DATA_UPDATE:
                                    isSaved = true;
                                    break;
                                case BoEventTypes.et_FORM_DEACTIVATE:
                                    success = true;
                                    isSaved = true;
                                    break;
                                case BoEventTypes.et_FORM_DRAW:
                                    //   timer.Start();
                                    break;
                                case BoEventTypes.et_FORM_KEY_DOWN:
                                    success = on_et_KEY_DOWN(pVal);
                                    break;
                                case BoEventTypes.et_FORM_LOAD:
                                    timer.Start();

                                    // OnFormLoaded(CurrentForm);
                                    break;
                                case BoEventTypes.et_FORM_MENU_HILIGHT:
                                    break;
                                case BoEventTypes.et_FORM_RESIZE:
                                    success = OnResize(pVal);
                                    break;
                                case BoEventTypes.et_FORM_UNLOAD:
                                    break;
                                case BoEventTypes.et_FORM_VISIBLE:
                                    {


                                        if (CurrentForm.Mode == BoFormMode.fm_VIEW_MODE)
                                        {


                                            //           OnFormLoaded(CurrentForm);
                                            //if (timer != null)
                                            //{
                                            //    timer.Dispose();
                                            //    timer = null;
                                            //}
                                            //timer = new System.Timers.Timer(100);
                                            //timer.Elapsed += timer_Elapsed;
                                            //    timer.Start();
                                        }
                                        success = true;
                                    }
                                    break;
                                case BoEventTypes.et_GOT_FOCUS:

                                    success = onet_GOT_FOCUS(pVal);
                                    break;
                                case BoEventTypes.et_GRID_SORT:
                                    break;
                                case BoEventTypes.et_ITEM_PRESSED:

                                    success = onITEM_PRESSED(pVal);

                                    OnBrowsButtonCLick_ATT(pVal);
                                    break;
                                case BoEventTypes.et_KEY_DOWN:
                                    break;
                                case BoEventTypes.et_LOST_FOCUS:
                                    success = OnLostFocus(pVal) && _OnLostFocus(pVal);
                                    try
                                    {
                                        if (CurrentForm.DefButton == "2")
                                            CurrentForm.DefButton = "1";
                                    }
                                    catch (Exception ex) { ex.AppendInLogFile(); }
                                    break;
                                case BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
                                    break;
                                case BoEventTypes.et_MATRIX_LINK_PRESSED:
                                    break;
                                case BoEventTypes.et_MATRIX_LOAD:
                                    break;
                                case BoEventTypes.et_MENU_CLICK:
                                    break;
                                case BoEventTypes.et_PICKER_CLICKED:
                                    break;
                                case BoEventTypes.et_PRINT:
                                    break;
                                case BoEventTypes.et_PRINT_DATA:
                                    break;
                                case BoEventTypes.et_PRINT_LAYOUT_KEY:
                                    break;
                                case BoEventTypes.et_RIGHT_CLICK:
                                    break;
                                case BoEventTypes.et_UDO_FORM_BUILD:
                                    break;
                                case BoEventTypes.et_UDO_FORM_OPEN:
                                    break;
                                case BoEventTypes.et_VALIDATE:

                                    success = OnValidate(pVal);
                                    if (success)
                                    {
                                        success = _OnValidate(pVal);
                                    }
                                    break;
                                default:
                                    break;
                            }

                        }
                        #endregion
                        else
                        #region Before Action
                          if (pVal.Before_Action)
                        {

                            try
                            {
                                if (pVal.ItemChanged)
                                {
                                    //ApplyCFLName(pVal);

                                    success = success && BeforeItemChanged(pVal);
                                }
                            }
                            catch (Exception ex)
                            {
                                if (System.Diagnostics.Debugger.IsAttached)
                                    Console.WriteLine(ex.Message + " " + ex.StackTrace);
                            }
                            switch (pVal.EventType)
                            {
                                case BoEventTypes.et_ALL_EVENTS:
                                    break;
                                case BoEventTypes.et_B1I_SERVICE_COMPLETE:
                                    break;
                                case BoEventTypes.et_CHOOSE_FROM_LIST:
                                    success = success && _OnCFLLoad((SAPbouiCOM.ChooseFromListEvent)pVal);
                                    success = success && Beforeet_CHOOSE_FROM_LIST((SAPbouiCOM.ChooseFromListEvent)pVal);
                                    try
                                    {

                                        if (((ChooseFromListEvent)pVal).ChooseFromListUID.ToLower() == "CFL_Asset".ToLower())
                                        {

                                            ApplyCFLNameFill("CFL_Asset", "U_AssetNam", "ItemName", 0);
                                            var customer = CurrentForm.ChooseFromLists.Item("CFL_Asset");
                                            var conds = new SAPbouiCOM.Conditions();
                                            var cond = conds.Add();
                                            cond.CondVal = "F";
                                            cond.Alias = "ItemType";
                                            cond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;

                                            //cond = conds.Add(); 
                                            //cond.Relationship = BoConditionRelationship.cr_AND;
                                            //cond.CondVal = "N";
                                            //cond.Alias = "frozenFor";
                                            //cond.Operation = BoConditionOperation.co_EQUAL;
                                            customer.SetConditions(conds);
                                        }
                                        if (((ChooseFromListEvent)pVal).ChooseFromListUID.ToLower() == "CFL_Cust".ToLower())
                                        {
                                            ApplyCFLNameFill("CFL_Cust", "U_CardName", "CardName", 0);
                                            var customer = CurrentForm.ChooseFromLists.Item("CFL_Cust");
                                            var conds = new SAPbouiCOM.Conditions();
                                            var cond = conds.Add();
                                            cond.CondVal = "C";
                                            cond.Alias = "CardType";
                                            cond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                            customer.SetConditions(conds);
                                        }
                                        if (((ChooseFromListEvent)pVal).ChooseFromListUID.ToLower() == "CFL_Ven".ToLower())
                                        {

                                            ApplyCFLNameFill("CFL_Ven", "U_CardName", "CardName", 0);
                                            var customer = CurrentForm.ChooseFromLists.Item("CFL_Ven");
                                            var conds = new SAPbouiCOM.Conditions();
                                            var cond = conds.Add();
                                            cond.CondVal = "S";
                                            cond.Alias = "CardType";
                                            cond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                            customer.SetConditions(conds);
                                        }

                                    }
                                    catch (Exception ex) { ex.AppendInLogFile(); }
                                    break;
                                case BoEventTypes.et_CLICK:
                                    if (pVal.ItemUID == CurrentForm.DefButton)
                                    {
                                        success = BeforeDefaultButtonClick(pVal);
                                    }
                                    if (success)
                                        success = BeforeClick(pVal);

                                    break;
                                case BoEventTypes.et_COMBO_SELECT:
                                    success = BeforComboSelected(pVal);
                                    break;
                                case BoEventTypes.et_DATASOURCE_LOAD:
                                    break;
                                case BoEventTypes.et_DOUBLE_CLICK:
                                    break;
                                case BoEventTypes.et_Drag:
                                    break;
                                case BoEventTypes.et_EDIT_REPORT:
                                    break;
                                case BoEventTypes.et_FORMAT_SEARCH_COMPLETED:
                                    break;
                                case BoEventTypes.et_FORM_ACTIVATE:
                                    break;
                                case BoEventTypes.et_FORM_CLOSE:
                                    break;
                                case BoEventTypes.et_FORM_DATA_ADD:

                                    break;
                                case BoEventTypes.et_FORM_DATA_DELETE:
                                    break;
                                case BoEventTypes.et_FORM_DATA_LOAD:

                                    EnforceViewModeOnCLose();
                                    break;
                                case BoEventTypes.et_FORM_DATA_UPDATE:
                                    break;
                                case BoEventTypes.et_FORM_DEACTIVATE:
                                    break;
                                case BoEventTypes.et_FORM_DRAW:
                                    break;
                                case BoEventTypes.et_FORM_KEY_DOWN:
                                    break;
                                case BoEventTypes.et_FORM_LOAD:
                                    break;
                                case BoEventTypes.et_FORM_MENU_HILIGHT:
                                    break;
                                case BoEventTypes.et_FORM_RESIZE:
                                    break;
                                case BoEventTypes.et_FORM_UNLOAD:
                                    break;
                                case BoEventTypes.et_FORM_VISIBLE:
                                    break;
                                case BoEventTypes.et_GOT_FOCUS:

                                    success = Beforeet_GOT_FOCUS(pVal);
                                    break;
                                case BoEventTypes.et_GRID_SORT:
                                    break;
                                case BoEventTypes.et_ITEM_PRESSED:
                                    success = before_et_ITEM_PRESSED(pVal);
                                    break;
                                case BoEventTypes.et_KEY_DOWN:
                                    success = success && Before_et_KEY_DOWN(pVal);
                                    break;
                                case BoEventTypes.et_LOST_FOCUS:
                                    success = BeforeLostFocus(pVal);
                                    break;
                                case BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
                                    break;
                                case BoEventTypes.et_MATRIX_LINK_PRESSED:
                                    break;
                                case BoEventTypes.et_MATRIX_LOAD:
                                    break;
                                case BoEventTypes.et_MENU_CLICK:
                                    break;
                                case BoEventTypes.et_PICKER_CLICKED:
                                    success = Before_et_PICKER_CLICKED(pVal);
                                    break;
                                case BoEventTypes.et_PRINT:
                                    break;
                                case BoEventTypes.et_PRINT_DATA:
                                    break;
                                case BoEventTypes.et_PRINT_LAYOUT_KEY:
                                    break;
                                case BoEventTypes.et_RIGHT_CLICK:
                                    success = success && BeforeRightClick(pVal);
                                    break;
                                case BoEventTypes.et_UDO_FORM_BUILD:
                                    break;
                                case BoEventTypes.et_UDO_FORM_OPEN:
                                    break;
                                case BoEventTypes.et_VALIDATE:
                                    success = success && BeforValidate(pVal);
                                    break;
                                default:
                                    break;
                            }

                        }
                        #endregion

                        success = success && ItemEvent(pVal);
                    }
                    catch (Exception ex)
                    {
                        // if (System.Diagnostics .Debugger .IsAttached )
                        //       ex.printAtStatusBar();               
                        Console.WriteLine(ex.Message);
                    }

                }
            }
            catch (Exception ex)
            {
                ex.PrintString();
            }

            CurrentForm.Freeze(false);

            //     this.Application.ItemEvent += BaseFormApplication_ItemEvent;
            BubbleEvent = success;
        }

        protected virtual bool Before_et_PICKER_CLICKED(SAPbouiCOM.ItemEvent pVal)
        {
            return true;
        }

        protected virtual bool BeforValidate(SAPbouiCOM.ItemEvent pVal)
        {
            return true;
        }

        protected virtual bool BeforeItemChanged(SAPbouiCOM.ItemEvent pVal)
        {
            return true;
        }

        protected virtual bool on_et_DOUBLE_CLICK(SAPbouiCOM.ItemEvent pVal)
        {
            return true;
        }

        protected virtual bool before_et_ITEM_PRESSED(ItemEvent pVal)
        {
            return true;
        }


        public void GetPrimaryKey(string TableName, out string PrimKey)
        {
            PrimKey = "";
            if (Company.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012 || Company.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008 ||
                Company.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014)
            {
                string query = "SELECT column_name as primarykey FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS AS TC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE AS p ON TC.CONSTRAINT_TYPE = 'PRIMARY KEY' AND TC.CONSTRAINT_NAME = p.CONSTRAINT_NAME and p.table_name='" + TableName + "' ORDER BY p.ORDINAL_POSITION";
                var recset = Company.DoQuery(query);
                PrimKey = recset.Fields.Item(0).Value.ToString();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(recset); GC.Collect();
            }
            else if (Company.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
            {
                var recset = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                recset.DoQuery(string.Format("select \"COLUMN_NAME\" from \"CONSTRAINTS\" where \"TABLE_NAME\"='{0}' and \"IS_PRIMARY_KEY\"='TRUE'", TableName));
                PrimKey = recset.Fields.Item(0).ToString();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(recset); GC.Collect();
            }
        }


        internal void Application_LayoutKeyEvent(ref LayoutKeyInfo eventInfo, out bool BubbleEvent)
        {
            try
            {
                var form = Application.Forms.Item(eventInfo.FormUID);

                //if (form.TypeEx == FormType)
                //    _FormUID = (String.IsNullOrEmpty(_FORMUID)) ? eventInfo.FormUID : _FORMUID;
                var docentry = MasterDS.GetValue("DocEntry", 0).Trim();
                eventInfo.LayoutKey = docentry;

            }
            catch
            {
            }
            BubbleEvent = true;
        }
        internal void Application_RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            bool success = true;
            try
            {
                if (!eventInfo.BeforeAction)
                {
                    switch (eventInfo.EventType)
                    {
                        case BoEventTypes.et_RIGHT_CLICK:
                            success = onRightClick(eventInfo);
                            break;
                    }
                }
                if (eventInfo.BeforeAction)
                {
                    switch (eventInfo.EventType)
                    {
                        case BoEventTypes.et_RIGHT_CLICK:
                            success = BeforeRightClick(eventInfo);
                            break;
                    }
                }
            }
            catch (Exception ex) { throw ex; }
            BubbleEvent = success;
        }

        protected virtual bool onFormDataClose()
        {
            return true;
        }
        protected virtual bool onFormDataClose(string Key)
        {
            return true;
        }

        protected virtual bool on_et_KEY_DOWN(ItemEvent pVal)
        {
            return true;
        }
        protected virtual bool Before_et_KEY_DOWN(ItemEvent pVal)
        {
            return true;
        }

        protected virtual bool onRightClick(ItemEvent pVal)
        {
            return true;
        }

        protected virtual bool BeforeRightClick(ItemEvent pVal)
        {
            return true;
        }

        protected virtual bool BeforeClick(ItemEvent pVal)
        {
            return true;
        }

        protected virtual bool onButtonClicked(ItemEvent pVal)
        {
            return true;
        }
        protected virtual bool onClicked(ItemEvent pVal)
        {
            return true;
        }

        protected virtual bool BeforComboSelected(ItemEvent pVal)
        {
            return true;
        }

        void BaseFormtimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            timer.Stop();
            //OnFormLoaded(CurrentForm);
        }

        protected virtual bool ValidationCheck()
        {
            return true;
        }
        private bool IsMendatoryValidated()
        {

            bool yes = true;
            try
            {
                foreach (var item in MendatoryFields)
                {
                    yes = !string.IsNullOrEmpty(MasterDS.GetValue(item.Key, 0).Trim());
                    if (!yes)
                    {
                        Application.SetStatusBarMessage(item.Value, BoMessageTime.bmt_Medium, true);
                        break;
                    }
                }
            }
            catch (Exception ex) { ex.AppendInLogFile(); }
            return yes;
        }
        private void ApplyCFLName(ItemEvent pVal)
        {
            var cflevent = (SAPbouiCOM.ChooseFromListEvent)pVal;

            if (_CFLArray.Keys.Contains(cflevent.ChooseFromListUID))
                foreach (var value in _CFLArray[cflevent.ChooseFromListUID])
                {

                    try
                    {
                        var mat = GetItem(pVal.ItemUID);
                        if (mat.Specific is SAPbouiCOM.Matrix)
                            (mat.Specific as SAPbouiCOM.Matrix).FlushToDataSource();
                    }
                    catch (Exception ex) { ex.AppendInLogFile(); }
                    var rowid = pVal.Row > 0 ? pVal.Row - 1 : 0;
                    dynamic str;

                    #region has colum
                    bool tr = false;
                    for (int i = 0; i < cflevent.SelectedObjects.Columns.Count; i++)
                    {
                        if (cflevent.SelectedObjects.Columns.Item(i).Name == value.Alias)
                        {
                            tr = true;
                            break;
                        }
                    }
                    #endregion
                    if (tr)
                    {
                        str = cflevent.SelectedObjects.GetValue(value.Alias, 0);
                        if (str is DateTime)
                        {
                            str = ((DateTime)str).ToString("yyyyMMdd");
                        }
                    }
                    else
                    {
                        SAPbobsCOM.CompanyService companyservice = null;
                        SAPbobsCOM.GeneralService GeneralService = null;
                        SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                        SAPbobsCOM.GeneralData oGeneralData = null;
                        companyservice = Company.GetCompanyService();
                        GeneralService = companyservice.GetGeneralService(cflevent.SelectedObjects.UniqueID);
                        oGeneralParams = ((SAPbobsCOM.GeneralDataParams)(GeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)));

                        oGeneralParams.SetProperty(cflevent.SelectedObjects.Columns.Item(0).Name, value);
                        oGeneralData = GeneralService.GetByParams(oGeneralParams);

                        str = oGeneralData.GetProperty(value.Alias);

                    }
                    CurrentForm.DataSources.DBDataSources.Item(value.datasourceNumber).SetValue(value.ItemUID, rowid, str);
                    try
                    {
                        var mat = GetItem(pVal.ItemUID);
                        if (mat.Specific is SAPbouiCOM.Matrix)
                            (mat.Specific as SAPbouiCOM.Matrix).LoadFromDataSourceEx();
                    }
                    catch (Exception ex) { ex.AppendInLogFile(); }
                    //if (value.datasourceNumber > 0)
                    //    try
                    //    {
                    //        for (int i = 0; i < CurrentForm.Items.Count; i++)
                    //        {
                    //            var m = CurrentForm.Items.Item(i).Specific;
                    //            if (m is SAPbouiCOM.Matrix)
                    //            {
                    //                (m as SAPbouiCOM.Matrix).LoadFromDataSource();
                    //            }
                    //        }
                    //    }
                    //    catch (Exception ex) { ex.PrintString(); }
                }
        }

        protected DataTable GetDataTableWithQuery(String Query)
        {
            var Dt = new DataTable();
            Dt.ExecuteQuery(Query);
            return Dt;
        }
        protected SAPbobsCOM.Recordset GetRecordSet(String Query)
        {
            var recset = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            recset.DoQuery(Query);
            recset.MoveFirst();
            return recset;
        }
        protected SAPbobsCOM.Recordset GetRecordSet(String Query, params object[] args)
        {

            var recset = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            recset.DoQuery(string.Format(Query, args));
            recset.MoveFirst();
            return recset;
        }
        private static void Add(String FormUID)
        {
            if (!FormTypes.Contains(FormUID))
                FormTypes.Add(FormUID);

        }
        /// <summary>
        /// Filters a CFL removing Previouse Conditions
        /// </summary>
        /// <param name="recset">DB representative Recordset</param>
        /// <param name="projectcfl">Cfl to be filtered</param>
        /// <param name="CondAlias">Alias name of CFL</param>
        /// <param name="DBField">DB Field to compare to Alias</param>
        public void filterCFLwhereEqual(SAPbobsCOM.Recordset recset, ChooseFromList projectcfl, String CondAlias, String DBField)
        {
            SAPbouiCOM.Form frm = CurrentForm;
            recset.MoveFirst();
            if (recset.RecordCount > 0)
            {
                //var columns = cflevent.SelectedObjects.Columns;
                var conds = new SAPbouiCOM.Conditions();
                var count = 0;
                while (!recset.EoF)
                {

                    var cond = conds.Add();
                    cond.Alias = CondAlias;
                    cond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    cond.CondVal = recset.Fields.Item(DBField).Value.ToString();
                    recset.MoveNext();

                    count++;
                    if (count < recset.RecordCount) cond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;

                }
                projectcfl.SetConditions(conds);
            }
            else
            {
                var conds = new Conditions();
                var cond = conds.Add();
                cond.Alias = CondAlias;
                cond.CondVal = "";
                cond.Operation = BoConditionOperation.co_EQUAL;

                projectcfl.SetConditions(conds);
            }


        }
        public void filterCFLwhereEqual(String query, string projectcfluid, String CondAlias, String DBField)
        {
            SAPbobsCOM.Recordset recset = GetRecordSet(query);
            var projectcfl = CurrentForm.ChooseFromLists.Item(projectcfluid);
            filterCFLwhereEqual(recset, projectcfl, CondAlias, DBField);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(recset); GC.Collect();
            //SAPbouiCOM.Form frm = CurrentForm;
            //recset.MoveFirst();
            //if (recset.RecordCount > 0)
            //{
            //    //var columns = cflevent.SelectedObjects.Columns;
            //    var conds = new SAPbouiCOM.Conditions();
            //    var count = 0;
            //    while (!recset.EoF)
            //    {

            //        var cond = conds.Add();
            //        cond.Alias = CondAlias;
            //        cond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            //        cond.CondVal = recset.Fields.Item(DBField).Value.ToString();
            //        recset.MoveNext();

            //        count++;
            //        if (count < recset.RecordCount) cond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;

            //    }
            //    projectcfl.SetConditions(conds);
            //}
            //else
            //{
            //    var conds = new Conditions();
            //    var cond = conds.Add();
            //    cond.Alias = CondAlias;
            //    cond.CondVal = "";
            //    cond.Operation = BoConditionOperation.co_BETWEEN;
            //    projectcfl.SetConditions(conds);
            //}


        }

        public Form LoadFromXML(string FileName)
        {
            var path = System.Windows.Forms.Application.StartupPath + "\\" + FileName;
            var xml = System.IO.File.ReadAllText(path);
            SAPbouiCOM.FormCreationParams fcp;
            fcp = Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams) as SAPbouiCOM.FormCreationParams;
            fcp.XmlData = System.Xml.Linq.XDocument.Parse(xml).ToString();
            return Application.Forms.AddEx(fcp);


        }
        protected void ApplyDisableOnClose(string ItemID)
        {
            if (!_ItemToClose.Contains(ItemID))
            {
                _ItemToClose.Add(ItemID);
            }


        }
        protected void ApplyCFLNameFill(String CFLUID, string ItemUID, string Alias, int datasourceNumber)
        {
            try
            {
                if (_CFLArray.Keys.Contains(CFLUID))
                {
                    var list = _CFLArray[CFLUID];
                    list.Add(new { ItemUID, Alias, datasourceNumber });
                }
                else
                {
                    var list = new List<dynamic>();
                    list.Add(new { ItemUID, Alias, datasourceNumber });
                    _CFLArray.Add(CFLUID, list);
                }
            }
            catch (Exception ex) { ex.AppendInLogFile(); }
        }

        internal string GetFormType()
        {
            return this.FormType;
        }

        public void Dispose()
        {
            //this.Application.ItemEvent -= BaseFormApplication_ItemEvent;
            //this.Application.RightClickEvent -= Application_RightClickEvent;
            //this.Application.LayoutKeyEvent -= Application_LayoutKeyEvent;
            //this.Application.MenuEvent -= BaseFormApplication_MenuEvent;

            // timer.Elapsed -= BaseFormtimer_Elapsed;
        }

    }
}
