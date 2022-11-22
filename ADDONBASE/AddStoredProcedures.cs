using ADDONBASE.Extensions;
using System;
using System.IO;
using System.Reflection;
using System.Threading;
namespace ADDONBASE
{
    public class AddStoredProcedures
    {
        string commandText;
        private void ExecuteSPs(Assembly thisAssembly, SAPbobsCOM.Company company, string str)
        {
            foreach (var n in thisAssembly.GetManifestResourceNames())
            {
                if (n.Substring(n.Length - 3) == "sql")
                {
                    using (Stream s = thisAssembly.GetManifestResourceStream(n))
                    {
                        using (StreamReader sr = new StreamReader(s))
                        {
                            try
                            {
                                commandText = sr.ReadToEnd();
                                if (n.StartsWith(thisAssembly.GetName().Name + ".Stored_Procedure"))
                                {
                                    var procname = n.Split('.')[2];
                                    var rs = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                                    rs.DoQuery(string.Format(str, commandText.Replace("'", "''"), procname));
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); GC.Collect();
                                }
                                if (n.StartsWith(thisAssembly.GetName().Name + ".SPCLIENT"))
                                {
                                    try
                                    {
                                        var rs = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                                        rs.DoQuery(commandText);
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); GC.Collect();
                                    }
                                    catch { }
                                }
                            }
                            catch (Exception ex)
                            {

                            }
                        }
                    }
                }
            }
        }
        internal AddStoredProcedures(SAPbobsCOM.Company company)
        {
            Thread th = new Thread(new ThreadStart(delegate ()
            {
                Assembly thisAssembly = Assembly.GetEntryAssembly();
                //SBO_SP_TransactionNotification
                //var names = from v in thisAssembly.GetManifestResourceNames() where v.Split ('.').Last ().Substring (0,2) == "SP" select v;
                var str = "";
                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ADDONBASE.spMain.sql"))
                {

                    try
                    {
                        using (StreamReader sr = new StreamReader(stream))
                        {
                            str = sr.ReadToEnd();
                        }
                    }
                    catch (Exception ex) { ex.AppendInLogFile(); }
                }
                ExecuteSPs(Assembly.GetExecutingAssembly(), company, str);
                ExecuteSPs(thisAssembly, company, str);


            }));
            th.Start();
        }
        public AddStoredProcedures()
        {
        }
        public void ExecuteSPs(string FolderPath)
        {
            var company = Initializer._Company;
            var files = Directory.GetFiles(FolderPath);

            var recset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            foreach (var Path in files)
            {
                try { recset.DoQuery(System.IO.File.ReadAllText(Path)); }
                catch (Exception ex) { Console.WriteLine(Path); }
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(recset); GC.Collect();
        }

    }
}
