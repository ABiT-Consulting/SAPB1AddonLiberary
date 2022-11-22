using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using ADDONBASE.Extensions;
namespace ADDONBASE
{
    public class EventQueryHandler
    {
        SAPbobsCOM.Company Company
        {
            get
            {
                return _Initializer.Company;
            }
        }
        SAPbouiCOM.Application Application
        {
            get
            {
                return _Initializer.SBO_Application;
            }
        }
        string querypath = "";
        public EventQueryHandler()
        {
            querypath = System.IO.Directory.GetCurrentDirectory() + "\\Queries\\Queries.xml";
        }
        public EventQueryHandler(string queryPath)
        {
            querypath = queryPath;
        }

        public void RunEvent(string eventNumber, SAPbouiCOM.Form oform = null)
        {
            var sqltype = "SQL";
            if (Company.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB) sqltype = "HANA";

            XDocument document = XDocument.Load(querypath);
            foreach (var v in document.Descendants("Query"))
            {
                if (v.Attribute("Event") != null && v.Attribute("Event").Value.ToLower() == eventNumber.ToLower())
                {
                    var query = v.Element(sqltype).Value;
                    var str = query;// "jkasdlue as 12&sdaj__3982[source=saj_/29]sj*2&7^;'asj[source=-js/.2]_jsld+=[source=283]";
                    var res = Regex.Matches(str, @"\$\[([A-Za-z0-9-_ \\/.:]+)\]").Cast<Match>().Select(match => match.Groups[1].Value).ToList();
                    foreach (var item in res)
                    {
                        var value = "";
                        if (item.Contains("."))
                        {
                            string dbname = item.Split('.')[0];
                            string fieldName = item.Split('.')[1];
                            value = oform.DataSources.DBDataSources.Item(dbname).GetValue(fieldName, 0).Trim();
                        }
                        else
                        {
                            value = oform.DataSources.UserDataSources.Item(item).Value;

                        }
                        query = query.Replace("$[" + item + "]", value);
                    }
                    Console.WriteLine(query);
                    var recset = Company.DoQuery(query);
                    if (recset.RecordCount > 0)
                    {
                        var name = recset.Fields.Item(0).Name;
                        var value = recset.Fields.Item(0).Value.ToString().Trim();
                        if (name.Contains("."))
                        {
                            var dbname = name.Split('.')[0];
                            var field = name.Split('.')[1];
                            oform.DataSources.DBDataSources.Item(dbname).SetValue(field, 0, value);
                        }
                        else
                        {
                            oform.DataSources.UserDataSources.Item(name).Value = value;
                        }
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
                    GC.Collect();

                }
            }

        }
    }
}
