using ADDONBASE.Extensions;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System;
using System.IO;
using System.Text;
using System.Xml;
namespace ADDONBASE
{
    public class _UserFormBase : UserFormBase
    {
        protected void ExtractQuery(string query, string queryName)
        {
            var outputPath = Path.Combine(Path.GetTempPath(), queryName);
            System.IO.File.WriteAllText(outputPath, query);
        }
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
        protected SAPbouiCOM.IForm CurrentForm
        {
            get
            {
                return this.UIAPIRawForm;
            }
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
                BubbleEvent = _Initializer.IsMenuResultClear;
            }
            catch (Exception ex)
            { ex.AppendInLogFile(); BubbleEvent = true; }
            _Initializer.IsMenuResultClear = true;
        }
        protected virtual bool BeforeCloseMenuClicked() { return true; }

        protected virtual bool BeforeMenuClicked(string p)
        {
            return true;
        }
        protected SAPbobsCOM.Recordset GetRecordSet(String Query, params object[] args)
        {

            var recset = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            recset.DoQuery(string.Format(Query, args));
            recset.MoveFirst();
            return recset;
        }
        protected SAPbobsCOM.Recordset GetRecordSet()
        {

            var recset = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            return recset;
        }
        public override void OnInitializeComponent()
        {
            Application.MenuEvent -= Application_MenuEvent;
            Application.MenuEvent += Application_MenuEvent;
            base.OnInitializeComponent();
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
        public string GetQuery(string key, params object[] args)
        {

            var query = string.Format(GetQuery(key), args);

            Logger.Logger.Log(query);
            var LogPath = System.IO.Path.GetTempPath();
            Logger.Logger.CreateLog(LogPath, key + ".txt");
            Logger.Logger.ClearLog();
            return query;
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
