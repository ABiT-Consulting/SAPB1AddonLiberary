using ADDONBASE.Extensions;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
namespace ADDONBASE
{
    public struct SearchParams
    {
        public string sql;
        public string QueryName;
        public string CatCode;
        public string FormID;
        public string ItemID;
        public long QueryID;

        public string HDRTable;
        //RowParams
        public string RowTable;
        public string ColumnName;
        //
    }

    public class CAddFormattedSearch
    {

        private void RemoveFormattedSearch(string formId, string itemId)
        {
            var B1Comp = _Initializer.Company;
            FormattedSearches fs = (FormattedSearches)B1Comp.GetBusinessObject(BoObjectTypes.oFormattedSearches);
            
            int count = fs.Browser.RecordCount;
            for (int i = 0; i < count; i++)
            {
                
                if (fs.GetByKey(i))
                {
                    if (fs.Remove() != 0)
                    {
                        B1Comp.GetLastErrorDescription().PrintString();
                    }
                    break;
                }
                fs.Browser.MoveNext();
            }
        }

        private string AddFormattedSearch(string zFormID,
              string zItemID,
              string zTargetColumn,
              BoFormattedSearchActionEnum zAction,
              int zQueryID, BoYesNoEnum zByField,
              BoYesNoEnum zRefresh,
              BoYesNoEnum zForceRefresh,
              string zFieldName = "")
        {
            RemoveFormattedSearch(zFormID, zItemID);
            var B1Comp = _Initializer.Company;
            FormattedSearches fs = (FormattedSearches)B1Comp.GetBusinessObject(BoObjectTypes.oFormattedSearches);

            fs.FormID = zFormID;
            fs.ItemID = zItemID;
            fs.ColumnID = zTargetColumn;
            fs.Action = zAction;
            fs.QueryID = zQueryID;
            fs.FieldID = zFieldName;
            fs.ForceRefresh = zForceRefresh;
            fs.Refresh = zRefresh;
            fs.ByField = zByField;
            if (fs.Add() == 0)
            {
                return B1Comp.GetNewObjectKey();
            }
            else
            {
                //B1App.MessageBox(B1Comp.GetLastErrorDescription)
                B1Comp.GetLastErrorDescription().PrintString();
                return "-1";
            }

        }
        public void createContactNameSearches(SearchParams zSP, bool zItemIsMatrix = true)
        {
            var fieldname = String.Empty;
            var sqlstmt = string.Empty;
            var B1App = Initializer._Application;

            try
            {
                sqlstmt = zSP.sql;
                if (zItemIsMatrix)
                    fieldname = zSP.ColumnName;
                else
                    fieldname = zSP.ItemID;

                //  sqlstmt = zSP.sql.Replace ( "&HDRTABLE", zSP.HDRTable);
                //     sqlstmt += "$[" +zSP.RowTable +"." +fieldname +"]";
                zSP.QueryID = CreateUserQry(zSP.QueryName, sqlstmt, Convert.ToInt32(zSP.CatCode));

                //then connect the FS to forms/fields needed
                if (zSP.QueryID != -1)
                    AddFormattedSearch(zSP.FormID, zSP.ItemID, zSP.ColumnName, BoFormattedSearchActionEnum.bofsaQuery, (int)zSP.QueryID, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, fieldname);

            }
            catch (Exception ex)
            {
                B1App.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
            }
        }
        private long CreateUserQry(string zName, string zSQL, int zCategory)
        {

            var B1Comp = _Initializer.Company;
            Recordset rs = B1Comp.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            long qryID = 0;
            string qryIDA = null;
            //stupid B1 error is sending some bad characters at end of _Active 
            var s = string.Format("select \"IntrnalKey\" from \"OUQR\" where \"QName\" = '{0}'", zName);// "select intrnalkey from ouqr where qname = '{0}'");
            rs.DoQuery(s);
            if (!rs.EoF)
            {
                qryID =Convert.ToInt64( rs.Fields.Item("IntrnalKey").Value);
            }
            else
            {
                UserQueries uq = B1Comp.GetBusinessObject(BoObjectTypes.oUserQueries) as UserQueries;
                uq.QueryDescription = zName;
                uq.Query = zSQL;
                uq.QueryCategory = zCategory;
                if (uq.Add() == 0)
                {
                    try
                    {
                        qryIDA = B1Comp.GetNewObjectKey();
                        for (int x = qryIDA.Length; x >= 1; x += -1)
                        {
                            int result;
                            if (int.TryParse(qryIDA, out result))
                            {
                                qryID = result;
                                break; // TODO: might not be correct. Was : Exit For
                            }
                            else
                            {
                                qryIDA = qryIDA.Substring(0, qryIDA.Length - 1);
                            }

                        }

                    }
                    catch (System.Runtime.InteropServices.COMException ex)
                    {
                    }
                    catch (Exception ex2)
                    {
                         
                    }

                }
                else
                {
                    qryID = -1;
                    _Initializer.Company.GetLastErrorDescription().PrintString();
                }

            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); // 
            return qryID;

        }

    }
}
