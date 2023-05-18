using ADDONBASE.Extensions;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace ADDONBASE
{
    public class _Initializer
    {

        public static bool IsMenuResultClear = true;
        private static SAPbouiCOM.Framework.Application _FRAMEWORK_APPLICATION;
        private static SAPbouiCOM.Application _SBO_Application;
        private static SAPbobsCOM.Company _Company;
        public static SAPbouiCOM.Framework.Application FRAMEWORK_APPLICATION
        {
            get
            {
                try
                {
                    #region Delete LogFile If Bigger than Size
                    var path = System.Environment.CurrentDirectory + "\\Error_LOG.txt";
                    if (System.IO.File.Exists(path))
                    {
                        var info = new System.IO.FileInfo(path);

                        System.IO.File.Delete(path);

                    }
                    #endregion
                    if (_FRAMEWORK_APPLICATION == null)
                    {
                        if (Environment.GetCommandLineArgs().Count() < 2)
                        {
                            _FRAMEWORK_APPLICATION = new SAPbouiCOM.Framework.Application();
                        }
                        else
                        {
                            _FRAMEWORK_APPLICATION = new SAPbouiCOM.Framework.Application(Environment.GetCommandLineArgs().GetValue(1).ToString());
                        }
                        _FRAMEWORK_APPLICATION.BeforeInitialized += _FRAMEWORK_APPLICATION_BeforeInitialized;
                        _FRAMEWORK_APPLICATION.AfterInitialized += _FRAMEWORK_APPLICATION_AfterInitialized;
                        SAPbouiCOM.Framework.Application.SBO_Application.AppEvent += application_AppEvent;
                    }
                }
                catch (Exception ex)
                {
                    ex.AppendInLogFile();
                }

                return _FRAMEWORK_APPLICATION;
            }
        }


        static void UpdateUDOFORMS()
        {
            #region PR_FormsUpdate
            #region Open All unattached Forms

            var strs = UDONames;


            var total = strs.Count();
            var count = 0;
            var pb = SBO_Application.StatusBar.CreateProgressBar("UDOS", total, false);

            foreach (string str in strs)
            {
                count++;
                try
                {
                    var udo = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD) as SAPbobsCOM.IUserObjectsMD;

                    udo.GetByKey(str);
                    var xml1 = System.IO.File.ReadAllText(@"FORMS\" + str + ".srf");
                    var xml = udo.FormSRF;
                    XDocument doc = default(XDocument);
                    XDocument doc2 = default(XDocument);
                    try
                    {
                        if (!string.IsNullOrEmpty(xml))
                            doc = XDocument.Parse(xml);
                    }
                    catch (Exception ex) { ex.AppendInLogFile(); }
                    try
                    {
                        if (!string.IsNullOrEmpty(xml1))
                            doc2 = XDocument.Parse(xml1);
                    }
                    catch (Exception ex) { ex.AppendInLogFile(); }

                    if (doc == null || !doc.ToString(SaveOptions.DisableFormatting).Equals(doc2.ToString(SaveOptions.DisableFormatting)))
                    {
                        try
                        {
                            _SBO_Application.ActivateMenuItem(str);
                            _SBO_Application.Forms.ActiveForm.Close();
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
            pb = _SBO_Application.StatusBar.CreateProgressBar("UDOS", total, false);
            count = 0;
            foreach (string str in strs)
            {
                count++;
                string xml = "";

                try
                {
                    var udo = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD) as SAPbobsCOM.IUserObjectsMD;

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
                            if (!string.IsNullOrEmpty(xml))
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
                            var recset = _Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                            xml1 = xml1.Replace("'", "''");
                            var s = string.Format("update \"OUDO\" set \"NewFormSrf\" = N'{0}' where \"Code\" = '{1}'", xml1, str);
                            recset.DoQuery(s);//string.Format("update OUDO set NewFormSrf = '{0}' where Code = '{1}';", xml1, str));
                            Marshal.ReleaseComObject(recset); GC.Collect();
                            udo.FormSRF = xml1;
                            var i = udo.Update();
                            if (i != 0)
                            {
                                //  _SBO_Application.MessageBox(Company.GetLastErrorDescription());

                            }
                            else
                            {
                                try { System.Runtime.InteropServices.Marshal.ReleaseComObject(udo); }

                                catch (Exception ex) { ex.AppendInLogFile(); }
                                udo = null;
                                _SBO_Application.ActivateMenuItem(str);
                            }
                            try { if (udo != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(udo); }

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
        static void application_AppEvent(BoAppEventTypes EventType)
        {
            if (EventType == BoAppEventTypes.aet_CompanyChanged)
            {

                // System.Windows.Forms.Application.Exit();
            }
            else
            if (EventType == BoAppEventTypes.aet_ShutDown)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(string.Format("Addon {0} is closing", ConfigurationManager.AppSettings["ADDON_Name"]));
                try { System.Windows.Forms.Application.Exit(); }
                catch (Exception ex) { ex.AppendInLogFile(); }
            }
        }
        public static string GetQuery(string key)
        {
            string dbType = "SQL";
            switch (Company.DbServerType)
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
        public static string GetQuery(string key, params object[] args)
        {
            var query =  string.Format(GetQuery(key), args);
           
            Logger.Logger.Log(query);
            var LogPath = System.IO.Path.GetTempPath();
            Logger.Logger.CreateLog(LogPath, key + ".txt");
            Logger.Logger.ClearLog();
            return query;
        }

        static string GetXmlNodeValue(string file, string xPath)
        {
            var doc = new XmlDocument();
            doc.Load(file);
            var xmlPath = string.Empty;
            var node = doc.DocumentElement.SelectSingleNode(xPath);
            return node.InnerText;
        }
        private static void _FRAMEWORK_APPLICATION_AfterInitialized(object sender, EventArgs e)
        {
            var companyname = Company.CompanyName;
            SBO_Application.SetStatusBarMessage(string.Format("Initialized Successfully! Connected to : {0}", companyname), BoMessageTime.bmt_Medium, false);
            System.Threading.Tasks.Task InitializingTasks = new System.Threading.Tasks.Task(() =>
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
                try { goIN = System.Configuration.ConfigurationManager.AppSettings["CreateUDFS"].ToString().Trim().ToLower() == "true"; }
                catch (Exception ex) { ex.AppendInLogFile(); }
                if (goIN)
                {
                    var assembly = Assembly.GetEntryAssembly();
                    List<string> classNamesUDT = new List<string>();
                    List<string> classNamesUDF = new List<string>();
                    List<string> classNamesUDO = new List<string>();

                    foreach(Type type in assembly.GetTypes())
                    {
                        if (Attribute.IsDefined(type, typeof(Attributes.TableNameAttribute)) && !type.IsAbstract && !type.IsInterface)
                        {
                            classNamesUDT.Add(type.FullName+", "+ assembly.FullName);
                        }
                    }
                    foreach (var type in assembly.GetTypes())
                    {
                        if (Attribute.IsDefined(type, typeof(Attributes.TableNameAttribute)) && !type.IsAbstract && !type.IsInterface)
                        {
                            classNamesUDF.Add(type.FullName + ", " + assembly.FullName);
                        }

                    }
                    foreach (Type type in assembly.GetTypes())
                    {
                        if (Attribute.IsDefined(type, typeof(Attributes.UDONameAttribute)) && !type.IsAbstract && !type.IsInterface)
                        {
                            classNamesUDO.Add(type.FullName + ", " + assembly.FullName);
                        }
                    }
                    foreach(var classname in classNamesUDT)
                    {
                        BusinessLogic.SAPB1Helper.CreateUDT(classname, Company);
                    }
                    foreach(var classname in classNamesUDF)
                    {
                        BusinessLogic.SAPB1Helper.CreateUDF(classname, Company);
                    }
                    //get all class names from the classes having attribute UDONameAttribute
                    foreach (var className in classNamesUDO)
                    {
                        BusinessLogic.SAPB1Helper.CreateUDO(className, Company);
                    }
                    try
                    {
                        var files = Directory.GetFiles(System.Windows.Forms.Application.StartupPath + "\\SCHEMAS\\");
                        var created = false;
                        //  string[] strs = null;
                        foreach (var file in files)
                        {
                            udCreatorfromXML.udcreator creator = new udCreatorfromXML.udcreator(Company);
                            var filname = Path.GetFileName(file);
                            if (filname.StartsWith("udt"))
                            {
                                creator.createTablesfromXML(file);
                            }
                        }
                        foreach (var file in files)
                        {
                            udCreatorfromXML.udcreator creator = new udCreatorfromXML.udcreator(Company);
                            var filname = Path.GetFileName(file);

                            if (filname.StartsWith("udf"))
                            {
                                creator.createUDFFromXML(file);
                            }
                        }

                        if (ConfigurationManager.AppSettings.AllKeys.Contains("UDONAMES"))
                        {
                            created = true;
                            try
                            {
                                AddUDO.Invoke();
                            }
                            catch { }
                        }
                        else
                            foreach (var file in files)
                            {
                                udCreatorfromXML.udcreator creator = new udCreatorfromXML.udcreator(Company);
                                var filname = Path.GetFileName(file);

                                if (filname.StartsWith("udo"))
                                {
                                    created = creator.createUDOFromXML(file);
                                    //  strs = creator.getUDONames(file).ToArray();
                                }
                            }
                        if (created)
                        {

                            UpdateUDOFORMS();
                            SBO_Application.MessageBox("UDO(s) Added You should Restart Company and Addon");
                        }
                        foreach (var file in files)
                        {
                            udCreatorfromXML.udcreator creator = new udCreatorfromXML.udcreator(Company);
                            var filname = Path.GetFileName(file);

                            if (filname.StartsWith("fms"))
                            {
                                creator.createFMSFromXML(file);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ex.AppendInLogFile();
                    }

                }
                #endregion
                #region Add Reports

                try
                {
                    Reporter reporter = new Reporter(Company);

                    reporter.UploadReports();
                }
                catch (Exception ex) { ex.AppendInLogFile(); }
                #endregion
                #region AttachUDOForms
                try
                {
                    if (System.Diagnostics.Debugger.IsAttached)
                    {
                        var udosCSV = UDONames.List2CSV();
                        if (!string.IsNullOrEmpty(udosCSV))
                        {
                            var recset = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                            recset.DoQuery(string.Format(@"select ""Code"", ""NewFormSrf"" from oudo where ""Code"" in ({0})", udosCSV));
                            while (!recset.EoF)
                            {
                                var Code = recset.Fields.Item(0).Value.ToString();
                                var NewFormSrf = recset.Fields.Item(1).Value.ToString();

                                System.IO.File.WriteAllText(Path.Combine("Forms", Code + ".srf"), NewFormSrf);
                                recset.MoveNext();
                            }
                            Marshal.ReleaseComObject(recset);

                        }


                    }
                }
                catch (Exception ex)
                {
                    ex.AppendInLogFile();
                }
                #endregion
                AddStoredProcedures sp = new AddStoredProcedures(Company);
                var menu = new Menu(SBO_Application, Company);

            });
            InitializingTasks.Start();
            SBO_Application.MenuEvent += application_MenuEvent;
        }
        public static Action AddUDO;
        public static void RunExplicitly()
        {
            SBO_Application.SetStatusBarMessage(string.Format("Initialized Successfully! Connected to : {0}", Company.CompanyName), BoMessageTime.bmt_Medium, false);
            System.Threading.Tasks.Task InitializingTasks = new System.Threading.Tasks.Task(() =>
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
                try { goIN = System.Configuration.ConfigurationManager.AppSettings["CreateUDFS"].ToString().Trim().ToLower() == "true"; }
                catch (Exception ex) { ex.AppendInLogFile(); }
                if (goIN)
                {
                    try
                    {
                        var files = Directory.GetFiles(System.Windows.Forms.Application.StartupPath + "\\SCHEMAS\\");
                        var created = false;
                        //  string[] strs = null;
                        foreach (var file in files)
                        {
                            udCreatorfromXML.udcreator creator = new udCreatorfromXML.udcreator(Company);
                            var filname = Path.GetFileName(file);
                            if (filname.StartsWith("udt"))
                            {
                                creator.createTablesfromXML(file);
                            }
                        }
                        foreach (var file in files)
                        {
                            udCreatorfromXML.udcreator creator = new udCreatorfromXML.udcreator(Company);
                            var filname = Path.GetFileName(file);

                            if (filname.StartsWith("udf"))
                            {
                                creator.createUDFFromXML(file);
                            }
                        }
                        foreach (var file in files)
                        {
                            udCreatorfromXML.udcreator creator = new udCreatorfromXML.udcreator(Company);
                            var filname = Path.GetFileName(file);

                            if (filname.StartsWith("udo"))
                            {
                                created = creator.createUDOFromXML(file);
                                //  strs = creator.getUDONames(file).ToArray();
                            }
                        }
                        if (created)
                        {

                            UpdateUDOFORMS();
                            SBO_Application.MessageBox("UDO(s) Added You should Restart Company and Addon");
                        }

                    }
                    catch (Exception ex)
                    {
                        ex.AppendInLogFile();
                    }
                }
                #endregion
                #region Add Reports

                try
                {
                    Reporter reporter = new Reporter(Company);

                    reporter.UploadReports();
                }
                catch (Exception ex) { ex.AppendInLogFile(); }
                #endregion
                AddStoredProcedures sp = new AddStoredProcedures(Company);
                var menu = new Menu(SBO_Application, Company);

            });
            InitializingTasks.Start();
            SBO_Application.MenuEvent += application_MenuEvent;
        }



        static void application_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            if (pVal.MenuUID == "Atudo")
            {
                UpdateUDOFORMS();
            }
            BubbleEvent = true;
        }
        private static void _FRAMEWORK_APPLICATION_BeforeInitialized(object sender, EventArgs e)
        {
            SBO_Application.SetStatusBarMessage("Initializing!", BoMessageTime.bmt_Short, false);
        }

        public static SAPbouiCOM.Application SBO_Application
        {
            get
            {
                try
                {
                    if (_SBO_Application == null)
                    {
                        _SBO_Application = SAPbouiCOM.Framework.Application.SBO_Application;
                        //    _SBO_Application.ItemEvent += _SBO_Application_ItemEvent;
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                return _SBO_Application;
            }

        }
        internal static Func<string, ItemEvent, bool> itemevent;
        static void _SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = itemevent.Invoke(FormUID, pVal);
            //   Console.WriteLine("FormUID={0}, ItemEvent={1}, ActionSuccess {2},BeforeAaction={3},ItemID ={4},ColumnID ={5}", pVal.FormUID, pVal.EventType, pVal.ActionSuccess, pVal.Before_Action, pVal.ItemUID, pVal.ColUID);
            //   BubbleEvent = true;
        }
        public static SAPbobsCOM.Company Company
        {
            get
            {
                if (_Company == null)
                {
                    _Company = SBO_Application.Company.GetDICompany() as SAPbobsCOM.Company;
                }
                return _Company;
            }
            set
            { _Company = value; }

        }
        internal static List<string> UDONames
        {
            get
            {

                List<string> strs = new List<string>();
                //Write here code for getting udo names from addon other location
                if (ConfigurationManager.AppSettings.AllKeys.Contains("UDONAMES"))
                {
                    var strings = ConfigurationManager.AppSettings["UDONAMES"].Split(',');
                    foreach (var item in strings)
                    {
                        strs.Add(item);
                    }

                }
                else
                {

                    var files = Directory.GetFiles(System.Windows.Forms.Application.StartupPath + "\\SCHEMAS\\");
                    foreach (var file in files)
                    {
                        udCreatorfromXML.udcreator creator = new udCreatorfromXML.udcreator(Initializer._Company);
                        var filname = Path.GetFileName(file);

                        if (filname.StartsWith("udo"))
                        {
                            var sts = creator.getUDONames(file).ToArray();
                            foreach (var item in sts)
                            {
                                strs.Add(item);
                            }
                        }

                    }
                }
                return strs;
            }
        }

    }
}