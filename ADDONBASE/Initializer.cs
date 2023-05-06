using ADDONBASE.Extensions;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Xml.Linq;
namespace ADDONBASE
{
    internal class ConnectToB1
    {

        public static SAPbobsCOM.Company GetCompany()
        {
            SAPbobsCOM.Company _company = null;

            if (_company == null)
            {
                _company = new SAPbobsCOM.Company()
                {
                    LicenseServer = ConfigurationManager.AppSettings["LicenseServer"],
                    DbServerType = (SAPbobsCOM.BoDataServerTypes)Convert.ToInt32(ConfigurationManager.AppSettings["DataBaseType"]),

                    Server = ConfigurationManager.AppSettings["DataBaseServer"],
                    CompanyDB = ConfigurationManager.AppSettings["DataBaseName"],
                    UserName = ConfigurationManager.AppSettings["CompanyUserName"],
                    Password = ConfigurationManager.AppSettings["CompanyPassword"],
                    DbUserName = ConfigurationManager.AppSettings["DataBaseUserName"],
                    DbPassword = ConfigurationManager.AppSettings["DataBasePassword"],
                    language = SAPbobsCOM.BoSuppLangs.ln_English
                };
            }
            if (!_company.Connected)
                _company.Connect();
            return _company;

        }

        internal /* TRANSINFO: WithEvents */ SAPbouiCOM.Application SBO_Application;
        internal SAPbobsCOM.Company oCompany;

        private void SetApplication()
        {

            // *******************************************************************
            // // Use an SboGuiApi object to establish connection
            // // with the SAP Business One application and return an
            // // initialized appliction object
            // *******************************************************************

            SAPbouiCOM.SboGuiApi SboGuiApi = null;
            string sConnectionString = null;

            SboGuiApi = new SAPbouiCOM.SboGuiApi();

            // // by following the steps specified above, the following
            // // statment should be suficient for either development or run mode
            sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
            try { sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1)); }
            catch (Exception ex) { ex.AppendInLogFile(); }

            // // connect to a running SBO Application
            try
            {
                SboGuiApi.Connect(sConnectionString);

            }
            catch (Exception)
            {
                sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";

                SboGuiApi.Connect(sConnectionString);

            }

            // // get an initialized application object

            SBO_Application = SboGuiApi.GetApplication(-1);

        }


        private int SetConnectionContext()
        {
            int setConnectionContextReturn = 0;

            string sCookie = null;
            string sConnectionContext = null;
            int lRetCode = 0;

            // // First initialize the Company object

            oCompany = this.SBO_Application.Company.GetDICompany() as SAPbobsCOM.Company;

            // // Acquire the connection context cookie from the DI API.
            sCookie = oCompany.GetContextCookie();

            // // Retrieve the connection context string from the UI API using the
            // // acquired cookie.
            sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie);

            // // before setting the SBO Login Context make sure the company is not
            // // connected

            if (oCompany.Connected == true)
            {
                oCompany.Disconnect();
            }

            // // Set the connection context information to the DI API.
            setConnectionContextReturn = oCompany.SetSboLoginContext(sConnectionContext);

            return setConnectionContextReturn;
        }

        internal static bool isHana
        {
            get
            {
                return Initializer._Company.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB;
            }
        }

        private int ConnectToCompany()
        {
            int connectToCompanyReturn = 0;

            // // Establish the connection to the company database.
            //  connectToCompanyReturn = oCompany.Connect();
            if (!oCompany.Connected)
            {
                oCompany = SBO_Application.Company.GetDICompany() as SAPbobsCOM.Company;
                if (!oCompany.Connected)
                    connectToCompanyReturn = oCompany.Connect();
                else
                    connectToCompanyReturn = 0;
            }
            return connectToCompanyReturn;
        }
        SAPbouiCOM.Framework.Application frameworkApplication = null;

        // UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
        private void Class_Initialize_Renamed()
        {



            // //*************************************************************
            // // send an "hello world" message
            // //*************************************************************


        }

        public ConnectToB1()
            : base()
        {
            try
            {

                if (Environment.GetCommandLineArgs().GetValue(1).ToString().Length < 1)
                {
                    frameworkApplication = new SAPbouiCOM.Framework.Application();
                }
                else
                {
                    frameworkApplication = new SAPbouiCOM.Framework.Application(Environment.GetCommandLineArgs().GetValue(1).ToString());
                }
                this.SBO_Application = SAPbouiCOM.Framework.Application.SBO_Application;
                this.oCompany = SBO_Application.Company.GetDICompany() as SAPbobsCOM.Company;

                SBO_Application.SetStatusBarMessage("DI Connected To: " + oCompany.CompanyName, BoMessageTime.bmt_Short, false);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
    }
    public class Initializer
    {

        private bool restarted = false;
        SAPbobsCOM.Company _company;
        SAPbouiCOM.Application _application;
        private SAPbobsCOM.Company company1;
        public SAPbobsCOM.Company company
        {
            get
            {
                return _company;
            }
        }
        public SAPbouiCOM.Application application
        {
            get
            {
                return _application;
            }
        }
        internal static SAPbobsCOM.Company _Company { get; set; }
        internal static SAPbouiCOM.Application _Application { get; set; }
        internal static string ADDON_NAME
        {
            get
            {
                return ConfigurationManager.AppSettings["ADDON_Name"];
            }
        }
        public Initializer()
        {

            Initialize();

        }
        public SAPbouiCOM.Framework.Application frameworkApplication;
        void Connect()
        {
            try
            {
                if (Environment.GetCommandLineArgs().Count() < 2)
                {
                    frameworkApplication = new SAPbouiCOM.Framework.Application();
                }
                else
                {
                    frameworkApplication = new SAPbouiCOM.Framework.Application(Environment.GetCommandLineArgs().GetValue(1).ToString());
                }
                this._application = SAPbouiCOM.Framework.Application.SBO_Application;
                this._company = this._application.Company.GetDICompany() as SAPbobsCOM.Company;
                if (this._company.Connected)
                    this._application.SetStatusBarMessage("Connected To: " + _company.CompanyName, BoMessageTime.bmt_Short, false);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
        /// <summary>
        /// Initialize Addon
        /// </summary>
        void Initialize()
        {

            #region Connect to B1
            Connect();

            this.application.AppEvent += application_AppEvent;
            // this.application.MenuEvent += application_MenuEvent;
            #endregion
            Task InitializingTasks = new Task(() =>
             {
                 #region file Generator
                 try
                 {
                     var path = System.Windows.Forms.Application.StartupPath;

                     var sqlPath = path + @"\SCHEMAS";
                     if (!System.IO.Directory.Exists(sqlPath))
                     {
                         #region Make Entery Assembly Schema Files
                         try
                         {
                             Assembly a = Assembly.GetEntryAssembly();

                             var names = a.GetManifestResourceNames();

                             System.IO.Directory.CreateDirectory(sqlPath);
                             foreach (var item in names)
                             {
                                 if (item.StartsWith(a.ManifestModule.Name.Replace(".exe", "") + ".SCHEMAS."))
                                 {
                                     var i = item.Split('.');
                                     var p = String.Format(@"{0}\{1}.{2}", sqlPath, i[2], i[3]);
                                     if (!System.IO.File.Exists(p))
                                         a.ExtractSaveResource(item, p);
                                 }
                             }

                         }
                         catch (Exception ex) { ex.AppendInLogFile(); }
                         #endregion
                         #region Make Calling Assembly Schema Files
                         try
                         {
                             Assembly a = Assembly.GetExecutingAssembly();

                             var names = a.GetManifestResourceNames();

                             System.IO.Directory.CreateDirectory(sqlPath);
                             foreach (var item in names)
                             {
                                 if (item.StartsWith(a.ManifestModule.Name.Replace(".dll", "") + ".SCHEMAS."))
                                 {
                                     var i = item.Split('.');
                                     var p = String.Format(@"{0}\{1}.{2}", sqlPath, i[2], i[3]);
                                     if (!System.IO.File.Exists(p))
                                         a.ExtractSaveResource(item, p);
                                 }
                             }

                         }
                         catch (Exception ex) { ex.AppendInLogFile(); }
                         #endregion

                     }
                     var frmsPath = path + @"\FORMS";
                     if (!System.IO.Directory.Exists(frmsPath))
                     {
                         #region Make Entry Assembly Forms Files
                         try
                         {
                             Assembly a = Assembly.GetEntryAssembly();

                             var names = a.GetManifestResourceNames();

                             System.IO.Directory.CreateDirectory(frmsPath);
                             foreach (var item in names)
                             {
                                 if (item.StartsWith(a.ManifestModule.Name.Replace(".exe", "") + ".FORMS."))
                                 {
                                     var i = item.Split('.');
                                     var p = String.Format(@"{0}\{1}.{2}", frmsPath, i[2], i[3]);
                                     if (!System.IO.File.Exists(p))
                                         a.ExtractSaveResource(item, p);
                                 }
                             }
                         }
                         catch
                         {
                         }
                         #endregion
                         #region Make Calling Assembly Forms Files
                         try
                         {
                             Assembly a = Assembly.GetExecutingAssembly();

                             var names = a.GetManifestResourceNames();

                             System.IO.Directory.CreateDirectory(frmsPath);
                             foreach (var item in names)
                             {
                                 if (item.StartsWith(a.ManifestModule.Name.Replace(".dll", "") + ".FORMS."))
                                 {
                                     var i = item.Split('.');
                                     var p = String.Format(@"{0}\{1}.{2}", frmsPath, i[2], i[3]);
                                     if (!System.IO.File.Exists(p))
                                         a.ExtractSaveResource(item, p);
                                 }
                             }
                         }
                         catch
                         {
                         }
                         #endregion

                     }
                     var srfPath = path + @"\srf";
                     if (!System.IO.Directory.Exists(srfPath))
                     {
                         #region Make Entry Assembly SRF Files
                         try
                         {
                             Assembly a = Assembly.GetEntryAssembly();

                             var names = a.GetManifestResourceNames();

                             System.IO.Directory.CreateDirectory(srfPath);
                             foreach (var item in names)
                             {
                                 if (item.StartsWith(a.ManifestModule.Name.Replace(".exe", "") + ".srf."))
                                 {
                                     var i = item.Split('.');
                                     var p = String.Format(@"{0}\{1}.{2}", srfPath, i[2], i[3]);
                                     if (!System.IO.File.Exists(p))
                                         a.ExtractSaveResource(item, p);
                                 }
                             }
                         }
                         catch (Exception ex) { ex.AppendInLogFile(); }
                         #endregion
                         #region Make Calling Assembly SRF Files
                         try
                         {
                             Assembly a = Assembly.GetExecutingAssembly();

                             var names = a.GetManifestResourceNames();

                             System.IO.Directory.CreateDirectory(srfPath);
                             foreach (var item in names)
                             {
                                 if (item.StartsWith(a.ManifestModule.Name.Replace(".dll", "") + ".srf."))
                                 {
                                     var i = item.Split('.');
                                     var p = String.Format(@"{0}\{1}.{2}", srfPath, i[2], i[3]);
                                     if (!System.IO.File.Exists(p))
                                         a.ExtractSaveResource(item, p);
                                 }
                             }
                         }
                         catch (Exception ex) { ex.AppendInLogFile(); }
                         #endregion

                     }

                 }
                 catch (Exception ex) { ex.AppendInLogFile(); }
                 #endregion
                 #region Create UDT UDF
                 var goIN = true;
                 try { goIN = ConfigurationManager.AppSettings["CreateUDFS"].ToString().Trim().ToLower() == "true"; }
                 catch (Exception ex) { ex.AppendInLogFile(); }
                 if (goIN)
                 {

                     try
                     {
                         var files = Directory.GetFiles(System.Windows.Forms.Application.StartupPath + "\\SCHEMAS\\");
                         var created = false;
                         string[] strs = null;
                         foreach (var file in files)
                         {
                             udCreatorfromXML.udcreator creator = new udCreatorfromXML.udcreator(this.company);
                             var filname = Path.GetFileName(file);
                             if (filname.StartsWith("udt"))
                             {
                                 creator.createTablesfromXML(file);
                             }
                         }
                         foreach (var file in files)
                         {
                             udCreatorfromXML.udcreator creator = new udCreatorfromXML.udcreator(this.company);
                             var filname = Path.GetFileName(file);

                             if (filname.StartsWith("udf"))
                             {
                                 creator.createUDFFromXML(file);
                             }
                         }
                         foreach (var file in files)
                         {
                             udCreatorfromXML.udcreator creator = new udCreatorfromXML.udcreator(this.company);
                             var filname = Path.GetFileName(file);

                             if (filname.StartsWith("udo"))
                             {
                                 created = creator.createUDOFromXML(file);
                                 strs = creator.getUDONames(file).ToArray();
                             }
                         }
                         if (created)
                         {
                             this.application.MessageBox("UDO(s) Added You should Restart Company and Addon");
                             this.application.ActivateMenuItem("3329");
                             restarted = true;
                         }
                         if (!created)
                         {
                             UpdateUDOFORMS();
                         }
                     }
                     catch (Exception ex)
                     {
                         ex.PrintString();
                     }
                 }
                 #endregion
                 #region Add Reports

                 try
                 {
                     Reporter reporter = new Reporter(this.company);

                     reporter.UploadReports();
                 }
                 catch (Exception ex) { ex.AppendInLogFile(); }
                 #endregion
                 AddStoredProcedures sp = new AddStoredProcedures(this.company);

             });
            InitializingTasks.Start();

            menu = new Menu(this.application, this.company);
            _Company = this.company;
            _Application = this._application;

            Assembly enteryAssembly = System.Reflection.Assembly.GetEntryAssembly();
            var resrcs = enteryAssembly.GetTypes();
            //var resrcs = (from v in enteryAssembly.GetTypes()
            //             where v.Namespace.Contains("BusinessLogic")
            //             select  v).ToArray();

            foreach (var file in resrcs)
            {
                try
                {
                    if (!file.FullName.Contains("+") && file.FullName.Contains(".BusinessLogic."))
                    {
                        var obj = enteryAssembly.CreateInstance(file.FullName.PrintString().ToString());

                        if (obj is BaseHandler)
                        {
                            var str = (obj as BaseHandler).GetFormType();
                            lstFormTypeClassType.Add(str, file.FullName.PrintString().ToString());
                            (obj as BaseHandler).Dispose();
                            obj = null;
                        }
                        else
                            lstobjects.Add(obj);

                    }

                }
                catch (Exception ex) { ex.AppendInLogFile(); }
            }
            timer = new System.Timers.Timer(1000);
            timer.Elapsed += timer_Elapsed;
            //timer.Start();


            _Application.ItemEvent += CurrentApplication_ItemEvent;
            _Application.RightClickEvent += CurrentApplication_RightClickEvent;
            _Application.LayoutKeyEvent += CurrentApplication_LayoutKeyEvent;
            _Application.MenuEvent += CurrentBaseFormApplication_MenuEvent;
            _Application.FormDataEvent += CurrentBaseApplication_FormDataEvent;

            this.application.MenuEvent += application_MenuEvent;

        }
        System.Timers.Timer timer;
        public event _IApplicationEvents_FormDataEventEventHandler GENERAL_FormDataEvent;
        bool CurrentBaseApplicationStarted = false;
        private void CurrentBaseApplication_FormDataEvent(ref BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            //_Application.FormDataEvent -= CurrentBaseApplication_FormDataEvent;
            try
            {

                if (GetCurrentHandler(BusinessObjectInfo.FormUID) != null && !CurrentBaseApplicationStarted)
                {
                    CurrentBaseApplicationStarted = true;
                    GetCurrentHandler(BusinessObjectInfo.FormUID).Application_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);

                    CurrentBaseApplicationStarted = false;

                }
                else BubbleEvent = true;

            }
            catch (Exception ex)
            {
                BubbleEvent = true;
                ex.PrintString();
            }
            //  _Application.FormDataEvent += CurrentBaseApplication_FormDataEvent;
        }
        bool CurrentBaseFormApplicationStarted = false;
        private void CurrentBaseFormApplication_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            //_application.MenuEvent -= CurrentBaseFormApplication_MenuEvent;
            try
            {
                if (CurrentHandler != null && !CurrentBaseFormApplicationStarted)
                {
                    CurrentBaseFormApplicationStarted = true;
                    CurrentHandler.Application_MenuEvent(ref pVal, out BubbleEvent);
                    CurrentBaseFormApplicationStarted = false;
                }
                else BubbleEvent = true;
            }
            catch (Exception ex)
            {
                BubbleEvent = true;
                ex.PrintString();
            }
            // _application.MenuEvent += CurrentBaseFormApplication_MenuEvent;
        }

        bool CurrentApplicationStarted = false;
        private void CurrentApplication_LayoutKeyEvent(ref LayoutKeyInfo eventInfo, out bool BubbleEvent)
        {
            // _Application.LayoutKeyEvent -= CurrentApplication_LayoutKeyEvent;
            try
            {
                if (GetCurrentHandler(eventInfo.FormUID) != null && !CurrentApplicationStarted)
                {
                    CurrentApplicationStarted = true;
                    GetCurrentHandler(eventInfo.FormUID).Application_LayoutKeyEvent(ref eventInfo, out BubbleEvent);
                    CurrentApplicationStarted = false;
                }
                else BubbleEvent = true;
            }
            catch (Exception ex)
            {
                BubbleEvent = true;
                ex.PrintString();
            }
            // _Application.LayoutKeyEvent += CurrentApplication_LayoutKeyEvent;
        }

        bool RightClickEventStarted = false;
        private void CurrentApplication_RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            //  _Application.RightClickEvent -= CurrentApplication_RightClickEvent;
            try
            {
                if (CurrentHandler != null && !RightClickEventStarted)
                {
                    RightClickEventStarted = true;
                    CurrentHandler.Application_RightClickEvent(ref eventInfo, out BubbleEvent);
                    RightClickEventStarted = false;
                }
                else BubbleEvent = true;
            }
            catch (Exception ex)
            {
                BubbleEvent = true;
                ex.PrintString();
            }
        }
        bool EventStarted = false;
        void UpdateFormList()
        {
            _Application.ItemEvent -= CurrentApplication_ItemEvent;
            try
            {
                if (CurrentHandler == null)
                {
                    CurrentHandler = CreateInstance(application.Forms.ActiveForm.TypeEx, application.Forms.ActiveForm.UniqueID);


                }
                else if (CurrentHandler.GetFormType() != application.Forms.ActiveForm.TypeEx)
                {
                    CurrentHandler = CreateInstance(application.Forms.ActiveForm.TypeEx, application.Forms.ActiveForm.UniqueID);

                }



            }
            catch (KeyNotFoundException ex)
            {

            }

            _Application.ItemEvent += CurrentApplication_ItemEvent;
        }
        //  string LoadedFormUID = "";
        private void CurrentApplication_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            try
            {
                if (pVal.EventType == BoEventTypes.et_FORM_VISIBLE && pVal.ActionSuccess)
                {
                    if (pVal.getForm().Items.Count == 0)
                    {
                        Task tsk = new Task(new Action(() =>
                        {

                            try
                            {
                                if (CurrentHandler == null)
                                {
                                    CurrentHandler = CreateInstance(application.Forms.Item(FormUID).TypeEx, FormUID);

                                }
                                else if (CurrentHandler.GetFormType() != application.Forms.Item(FormUID).TypeEx)
                                {
                                    CurrentHandler = CreateInstance(application.Forms.Item(FormUID).TypeEx, FormUID);
                                }




                            }
                            catch (KeyNotFoundException ex)
                            {
                            }
                            catch (InvalidCastException ex)
                            {

                                System.Windows.Forms.Application.Exit();
                            }
                            catch (Exception ex)
                            {

                            }

                        }));
                        tsk.Start();
                    }

                }

                if (GetCurrentHandler(pVal.FormUID) != null && !EventStarted)
                {
                    EventStarted = true;
                    GetCurrentHandler(pVal.FormUID).Application_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                    EventStarted = false;
                }
                else BubbleEvent = true;
            }
            catch (Exception ex)
            {
                ex.PrintString();
                EventStarted = false;
                BubbleEvent = true;
            }

            (pVal.EventType.ToString() + " ActionSuccess =" + pVal.ActionSuccess.ToString() + " BeforeAction =" + pVal.BeforeAction.ToString() + " Item Count=" + pVal.getForm().Items.Count.ToString()).PrintString();
            //   _Application.ItemEvent -= CurrentApplication_ItemEvent;
            //  _Application.ItemEvent += CurrentApplication_ItemEvent;

        }

        private BaseHandler GetCurrentHandler(string p)
        {
            if (FormBaseHandler.Count(x => x.Key == p) > 0)
                return FormBaseHandler[p];
            else
                return null;

        }

        void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            //Task tsk = new Task(new Action(() =>
            //{
            //    (sender as System.Timers.Timer).Stop();

            //    try
            //    {
            //        if (CurrentHandler == null)
            //        {
            //            CurrentHandler = CreateInstance(application.Forms.Item (LoadedFormUID).TypeEx);

            //        }
            //        else if (CurrentHandler.GetFormType() != application.Forms.Item(LoadedFormUID).TypeEx)
            //        {
            //            //  CurrentHandler.Dispose();
            //            CurrentHandler = CreateInstance(application.Forms.Item(LoadedFormUID).TypeEx);
            //        }




            //    }
            //    catch (KeyNotFoundException ex)
            //    {
            //        // CurrentHandler = CreateInstance(application.Forms.ActiveForm.TypeEx);

            //        //  CurrentHandler = null;
            //    }
            //    catch (InvalidCastException ex)
            //    {

            //        System.Windows.Forms.Application.Exit();
            //    }
            //    catch (Exception ex)
            //    {

            //    }

            //    //(sender as System.Timers.Timer).Start();
            //}));
            //tsk.Start();

        }

        List<object> lstobjects = new List<object>();
        #region FormTypeToHandler
        Dictionary<string, BaseHandler> FormBaseHandler = new Dictionary<string, BaseHandler>();
        BaseHandler CurrentHandler
        {
            get
            {
                try
                {
                    //if(System.Diagnostics .Debugger .IsAttached )
                    //Console.WriteLine(FormBaseHandler.Count);
                    foreach (var item in FormBaseHandler)
                    {
                        if (!application.Forms.HasForm(item.Key))
                        {
                            FormBaseHandler.Remove(item.Key);

                        }
                    }
                    if (FormBaseHandler.Count(x => x.Key == application.Forms.ActiveForm.UniqueID.ToString()) > 0)
                    {
                        return FormBaseHandler[application.Forms.ActiveForm.UniqueID];

                    }
                    else
                        return null;
                }
                catch
                {
                    return null;
                }
            }
            set
            {
                if (value is BaseHandler)
                {
                    if (!FormBaseHandler.ContainsKey(application.Forms.ActiveForm.UniqueID))
                    {
                        //if(string.IsNullOrEmpty ((value as BaseHandler)._FORMUID))
                        //    (value as BaseHandler)._FORMUID = application.Forms.ActiveForm.UniqueID;
                        FormBaseHandler.Add(application.Forms.ActiveForm.UniqueID, value);
                    }

                }
            }
        }
        Dictionary<string, string> lstFormTypeClassType = new Dictionary<string, string>();
        BaseHandler CreateInstance(string FormType, string LoadedFormUID)
        {
            var frm = application.Forms.Item(LoadedFormUID);
            if (!this.FormBaseHandler.ContainsKey(frm.UniqueID))
            {
                Assembly enteryAssembly = System.Reflection.Assembly.GetEntryAssembly();
                if (lstFormTypeClassType.ContainsKey(FormType))
                {
                    var inst = enteryAssembly.CreateInstance(lstFormTypeClassType[FormType]) as BaseHandler;
                    inst._FORMUID = frm.UniqueID;


                    return inst;
                }
                else return null;
            }
            else
            {
                return this.FormBaseHandler[frm.UniqueID];
            }

        }
        #endregion
        public static string CompanyName
        {
            get
            {
                return _Company.CompanyName;
            }
        }
        public Initializer(out SAPbobsCOM.Company company)
        {

            #region Connect to B1
            ConnectToB1 connection = new ConnectToB1();
            _Company = connection.oCompany;
            this._company = connection.oCompany;
            // this.application.MenuEvent += application_MenuEvent;
            #endregion

            company = this.company;


        }
        public Initializer(out SAPbobsCOM.Company company, out SAPbouiCOM.Application application)
        {
            Initialize();

            company = this.company;
            application = this.application;


        }
        void CopyFilesToPath(String sqlPath, string tstString)
        {
            if (!System.IO.Directory.Exists(sqlPath))
            {
                #region Make Entery Assembly Schema Files
                try
                {
                    Assembly a = Assembly.GetEntryAssembly();

                    var names = a.GetManifestResourceNames();

                    System.IO.Directory.CreateDirectory(sqlPath);
                    foreach (var item in names)
                    {
                        if (item.StartsWith(a.ManifestModule.Name.Replace(".exe", "") + tstString))
                        {
                            var i = item.Split('.');
                            var p = String.Format(@"{0}\{1}.{2}", sqlPath, i[2], i[3]);
                            if (!System.IO.File.Exists(p))
                                a.ExtractSaveResource(item, p);
                        }
                    }

                }
                catch (Exception ex) { ex.AppendInLogFile(); }
                #endregion
                #region Make Calling Assembly Schema Files
                try
                {
                    Assembly a = Assembly.GetExecutingAssembly();

                    var names = a.GetManifestResourceNames();

                    System.IO.Directory.CreateDirectory(sqlPath);
                    foreach (var item in names)
                    {
                        if (item.StartsWith(a.ManifestModule.Name.Replace(".dll", "") + ".SCHEMAS."))
                        {
                            var i = item.Split('.');
                            var p = String.Format(@"{0}\{1}.{2}", sqlPath, i[2], i[3]);
                            if (!System.IO.File.Exists(p))
                                a.ExtractSaveResource(item, p);
                        }
                    }

                }
                catch (Exception ex) { ex.AppendInLogFile(); }
                #endregion

            }
        }
        public Initializer(SAPbobsCOM.Company company, SAPbouiCOM.Application application)
        {


            this._application = application;
            this._company = company;
            this.application.AppEvent += application_AppEvent;
            this.application.MenuEvent += application_MenuEvent;
            #region file Generator
            try
            {
                var path = System.Windows.Forms.Application.StartupPath;
                var sqlPath = path + @"\SCHEMAS";
                CopyFilesToPath(sqlPath, ".SCHEMAS.");
                var frmsPath = path + @"\FORMS";
                CopyFilesToPath(frmsPath, ".FORMS.");
                var srfPath = path + @"\srf";
                CopyFilesToPath(srfPath, ".srf.");
            }
            catch (Exception ex) { ex.AppendInLogFile(); }
            #endregion
            #region Create UDT UDF
            try
            {
                var files = Directory.GetFiles(System.Windows.Forms.Application.StartupPath + "\\SCHEMAS\\");
                var created = false;
                string[] strs = null;
                foreach (var file in files)
                {
                    udCreatorfromXML.udcreator creator = new udCreatorfromXML.udcreator(this.company);
                    var filname = Path.GetFileName(file);
                    if (filname.StartsWith("udt"))
                    {
                        creator.createTablesfromXML(file);
                    }
                }
                foreach (var file in files)
                {
                    udCreatorfromXML.udcreator creator = new udCreatorfromXML.udcreator(this.company);
                    var filname = Path.GetFileName(file);

                    if (filname.StartsWith("udf"))
                    {
                        creator.createUDFFromXML(file);
                    }
                }
                foreach (var file in files)
                {
                    udCreatorfromXML.udcreator creator = new udCreatorfromXML.udcreator(this.company);
                    var filname = Path.GetFileName(file);

                    if (filname.StartsWith("udo"))
                    {
                        created = creator.createUDOFromXML(file);
                        strs = creator.getUDONames(file).ToArray();
                    }
                }
                if (created)
                {
                    this.application.MessageBox("UDO(s) Added You should Restart Company and Addon");
                    this.application.ActivateMenuItem("3329");
                    restarted = true;
                }
                if (!created)
                {
                    UpdateUDOFORMS();
                }
            }
            catch (Exception ex) { ex.AppendInLogFile(); }
            #endregion
            #region Add Reports

            try
            {
                Reporter reporter = new Reporter(this.company);

                reporter.UploadReports();
            }
            catch (Exception ex) { ex.AppendInLogFile(); }
            #endregion

            AddStoredProcedures sp = new AddStoredProcedures(this.company);

            menu = new Menu(this.application, this.company);
            try
            {
                _Company = this.company;
                _Application = this._application;
                Assembly enteryAssembly = System.Reflection.Assembly.GetEntryAssembly();

            }
            catch
            {

            }

            company = this.company;
            application = this.application;


        }


        internal static List<string> UDONames
        {
            get
            {
                var files = Directory.GetFiles(System.Windows.Forms.Application.StartupPath + "\\SCHEMAS\\");

                List<string> strs = new List<string>();

                foreach (var file in files)
                {
                    udCreatorfromXML.udcreator creator = new udCreatorfromXML.udcreator(Initializer._Company);
                    var filname = Path.GetFileName(file);

                    if (filname.StartsWith("udo"))
                    {
                        try
                        {
                            var sts = creator.getUDONames(file).ToArray();
                            foreach (var item in sts)
                            {
                                strs.Add(item);
                            }
                        }
                        catch { }
                    }
                }
                return strs;
            }
        }
        /// <summary>
        /// Update UDO Forms
        /// </summary>
        void UpdateUDOFORMS()
        {
            #region PR_FormsUpdate
            #region Open All unattached Forms
            //var files = Directory.GetFiles(System.Windows.Forms.Application.StartupPath + "\\SCHEMAS\\");

            var strs = UDONames;

            //foreach (var file in files)
            //{
            //    udCreatorfromXML.udcreator creator = new udCreatorfromXML.udcreator(this.company);
            //    var filname = Path.GetFileName(file);

            //    if (filname.StartsWith("udo"))
            //    {
            //        strs = creator.getUDONames(file).ToArray();
            //    }
            //}
            var total = strs.Count();
            var count = 0;
            var pb = ADDONBASE.Initializer._Application.StatusBar.CreateProgressBar("UDOS", total, false);

            foreach (string str in strs)
            {
                count++;
                try
                {
                    var udo = this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD) as SAPbobsCOM.IUserObjectsMD;

                    udo.GetByKey(str);
                    var xml1 = System.IO.File.ReadAllText(@"FORMS\" + str + ".srf");
                    var xml = udo.FormSRF;
                    XDocument doc = default(XDocument);
                    XDocument doc2 = default(XDocument);
                    try
                    {
                        doc = XDocument.Parse(xml);
                    }
                    catch (Exception ex) { ex.AppendInLogFile(); }
                    try
                    {
                        doc2 = XDocument.Parse(xml1);
                    }
                    catch (Exception ex) { ex.AppendInLogFile(); }

                    if (doc == null || !doc.ToString(SaveOptions.DisableFormatting).Equals(doc2.ToString(SaveOptions.DisableFormatting)))
                    {
                        try
                        {
                            this.application.ActivateMenuItem(str);
                            this.application.Forms.ActiveForm.Close();
                        }
                        catch (Exception ex) { ex.AppendInLogFile(); }

                    }

                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(udo); }

                    catch (Exception ex) { ex.AppendInLogFile(); }
                    udo = null;
                }
                catch (FileNotFoundException ex) { }
                pb.Value = count;
            }
            pb.Stop();
            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(pb); }

            catch (Exception ex) { ex.AppendInLogFile(); }
            #endregion
            pb = ADDONBASE.Initializer._Application.StatusBar.CreateProgressBar("UDOS", total, false);
            count = 0;
            foreach (string str in strs)
            {
                count++;
                string xml = "";

                try
                {
                    var udo = this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD) as SAPbobsCOM.IUserObjectsMD;

                    udo.GetByKey(str);
                    var xml1 = "";

                    try
                    {
                        xml1 = System.IO.File.ReadAllText(@"FORMS\" + str + ".srf");

                        #region Save xml1
                        xml = udo.FormSRF;
                        XDocument doc = default(XDocument);
                        XDocument doc2 = default(XDocument);
                        try
                        {
                            doc = XDocument.Parse(xml);
                        }
                        catch (Exception ex) { ex.AppendInLogFile(); }
                        try
                        {
                            doc2 = XDocument.Parse(xml1);
                        }
                        catch (Exception ex) { ex.AppendInLogFile(); }

                        if (doc == null || !doc.ToString(SaveOptions.DisableFormatting).Equals(doc2.ToString(SaveOptions.DisableFormatting)))
                        {
                            var recset = _company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                            xml1 = xml1.Replace("'", "''");
                            var s = string.Format("update \"OUDO\" set \"NewFormSrf\" = N'{0}' where \"Code\" = '{1}'", xml1, str);
                            recset.DoQuery(s);//string.Format("update OUDO set NewFormSrf = '{0}' where Code = '{1}';", xml1, str));

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(recset); GC.Collect();
                            udo.FormSRF = xml1;
                            var i = udo.Update();
                            if (i != 0)
                            {
                                //  this.application.MessageBox(this.company.GetLastErrorDescription());

                            }
                            else
                            {
                                try { System.Runtime.InteropServices.Marshal.ReleaseComObject(udo); }

                                catch (Exception ex) { ex.AppendInLogFile(); }
                                udo = null;
                                this.application.ActivateMenuItem(str);
                            }
                            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(udo); }

                            catch (Exception ex) { ex.AppendInLogFile(); }
                            udo = null;
                        }

                        #endregion
                    }
                    catch (Exception ex) { ex.AppendInLogFile(); }

                }
                catch (Exception ex) { ex.AppendInLogFile(); }
                pb.Value = count;
            }
            pb.Stop();
            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(pb); }

            catch (Exception ex) { ex.AppendInLogFile(); }
            //}
            #endregion

        }
        void application_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            if (pVal.MenuUID == "Atudo")
            {
                UpdateUDOFORMS();
            }
            BubbleEvent = true;
        }

        Menu menu;
        void application_AppEvent(BoAppEventTypes EventType)
        {
            if (EventType == BoAppEventTypes.aet_CompanyChanged)
            {

                // System.Windows.Forms.Application.Exit();
                menu.AddMenuItems();
            }
            if (EventType == BoAppEventTypes.aet_ShutDown)
            {
                application.SetStatusBarMessage(string.Format("Addon {0} is closing", ConfigurationManager.AppSettings["ADDON_Name"]), BoMessageTime.bmt_Short, false);
                try { System.Windows.Forms.Application.Exit(); }
                catch (Exception ex) { ex.AppendInLogFile(); }
            }
        }


        //public static void ExtractSaveResource(String filename, String location)
        //{
        //    Assembly a = Assembly.GetEntryAssembly();
        //    Stream resFilestream = a.GetManifestResourceStream(filename);
        //    if (resFilestream != null)
        //    {
        //        BinaryReader br = new BinaryReader(resFilestream);
        //        FileStream fs = new FileStream(location, FileMode.Create); // say 
        //        BinaryWriter bw = new BinaryWriter(fs);
        //        byte[] ba = new byte[resFilestream.Length];
        //        resFilestream.Read(ba, 0, ba.Length);
        //        bw.Write(ba);
        //        br.Close();
        //        bw.Close();
        //        resFilestream.Close();
        //    }
        //    // this.Close(); 
        //}

        /// <summary>
        /// Dispose Function
        /// </summary>
        public void dispose()
        {

            var path = System.Windows.Forms.Application.StartupPath + @"\SCHEMAS";
            try { System.IO.Directory.Delete(path, true); }
            catch (Exception ex) { ex.AppendInLogFile(); }
            var path2 = System.Windows.Forms.Application.StartupPath + @"\FORMS";
            try { System.IO.Directory.Delete(path2, true); }
            catch (Exception ex) { ex.AppendInLogFile(); }
            path2 = System.Windows.Forms.Application.StartupPath + @"\srf";
            try { System.IO.Directory.Delete(path2, true); }
            catch (Exception ex) { ex.AppendInLogFile(); }
            path2 = System.Windows.Forms.Application.StartupPath + @"\Stored_Procedure";
            try { System.IO.Directory.Delete(path2, true); }
            catch (Exception ex) { ex.AppendInLogFile(); }
            path2 = System.Windows.Forms.Application.StartupPath + @"\SPCLIENT";
            try { System.IO.Directory.Delete(path2, true); }
            catch (Exception ex) { ex.AppendInLogFile(); }
            path2 = System.Windows.Forms.Application.StartupPath + @"\Reports";
            try { System.IO.Directory.Delete(path2, true); }
            catch (Exception ex) { ex.AppendInLogFile(); }

        }
        public static void ExtractFiles(string rptPath, string Namespace)
        {
            //var rptPath = path + @"\CRReports";
            if (!System.IO.Directory.Exists(rptPath))
            {
                #region Make Entry Assembly SRF Files
                try
                {
                    Assembly a = Assembly.GetEntryAssembly();

                    var names = a.GetManifestResourceNames();

                    System.IO.Directory.CreateDirectory(rptPath);
                    foreach (var item in names)
                    {
                        if (item.StartsWith(a.ManifestModule.Name.Replace(".exe", "") + "." + Namespace + "."))
                        {
                            var i = item.Split('.');
                            var p = String.Format(@"{0}\{1}.{2}", rptPath, i[2], i[3]);
                            if (!System.IO.File.Exists(p))
                                a.ExtractSaveResource(item, p);
                        }
                    }
                }
                catch (Exception ex) { ex.AppendInLogFile(); }
                #endregion
                #region Make Calling Assembly SRF Files
                try
                {
                    Assembly a = Assembly.GetExecutingAssembly();

                    var names = a.GetManifestResourceNames();

                    System.IO.Directory.CreateDirectory(rptPath);
                    foreach (var item in names)
                    {
                        if (item.StartsWith(a.ManifestModule.Name.Replace(".dll", "") + "." + Namespace + "."))
                        {
                            var i = item.Split('.');
                            var p = String.Format(@"{0}\{1}.{2}", rptPath, i[2], i[3]);
                            if (!System.IO.File.Exists(p))
                                a.ExtractSaveResource(item, p);
                        }
                    }
                }
                catch (Exception ex) { ex.AppendInLogFile(); }
                #endregion

            }
        }

        //   public event EventHandler SBO_Close;

        public void Run()
        {
            frameworkApplication.Run();
        }
    }
}
