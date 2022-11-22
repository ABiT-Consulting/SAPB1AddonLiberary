using ADDONBASE.Extensions;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
namespace ADDONBASE
{
    public class _SystemFormBase : SystemFormBase
    {
        protected object GetFirstFromQuery(string Query, params object[] obj)
        {
            object value = null;
            var recset = GetRecordSet(string.Format(Query, obj));
            recset.MoveFirst();
            if (!recset.EoF)
                value = recset.Fields.Item(0).Value;

            System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
            GC.Collect();
            return value;
        }

        protected List<object> GetValuesFromQuery(string Query, params object[] obj)
        {
            List<object> objs = new List<object>();

            var recset = GetRecordSet(string.Format(Query, obj));
            recset.MoveFirst();
            int i = 0;
            while (!recset.EoF)
            {
                objs.Add(recset.Fields.Item(i).Value);
                i++;
                recset.MoveNext();
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
            GC.Collect();
            return objs;
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
        protected SAPbobsCOM.Recordset GetRecordSet(String Query, params object[] args)
        {

            var recset = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            recset.DoQuery(string.Format(Query, args));
            recset.MoveFirst();
            return recset;
        }
        protected string getObjectKeyFromXML(String XML)
        {
            return Extensions.Extensions.getObjectKeyFromXML(XML);
        }
        #region Menu Handler
        public override void OnInitializeComponent()
        {
            Application.MenuEvent -= Application_MenuEvent;
            Application.MenuEvent += Application_MenuEvent;
            base.OnInitializeComponent();
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

        protected virtual bool BeforeMenuClicked(string p)
        {
            return true;
        }

        protected virtual bool BeforeCloseMenuClicked() { return true; }

        #endregion
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
    public struct DatabaseTypes
    {
        public const string HANA = "HANA";
        public const string SQL = "SQL";
        public const string ORACLE = "ORACLE";
    }
}
