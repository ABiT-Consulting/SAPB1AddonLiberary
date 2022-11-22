using ADDONBASE.Extensions;
using SAPbobsCOM;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
namespace ADDONBASE
{
    public class Reporter
    {
        SAPbobsCOM.Company oCompany;
        public Reporter(SAPbobsCOM.Company company)
        {
            oCompany = company;

        }
        public Reporter()
        {
            oCompany = Initializer._Company;

        }
        internal void UploadReports()
        {
            var assembly = Assembly.GetEntryAssembly();
            var resources = assembly.GetManifestResourceNames();
            var name_resources = (from v in resources where getFileName(v) == "NameInSAP.txt" select v).ToArray();

            var rep_resources = (from v in resources where getFileName(v).EndsWith(".rpt") select v).ToArray();
            int half = resources.Count() / 2;
            for (int j = 0; j < rep_resources.Length; j++)
            {
                var resource = rep_resources[j];
                var streamrept = Assembly.GetEntryAssembly().GetManifestResourceStream(resource.ToString());

                string[] ar = resource.Split('.');
                var Path = ar[3];//+".rpt";
                var fsr = new FileStream(Path, FileMode.OpenOrCreate);

                for (int i = 0; i < streamrept.Length - 1; i++)
                {
                    var b = (byte)streamrept.ReadByte();
                    fsr.WriteByte(b);
                }

                fsr.Close();
                #region get name
                var name = getNamespacePart(resource);
                var reptypes = name.Split('.');
                var reptype = reptypes[reptypes.Length - 1];
                name = (from v in name_resources where getNamespacePart(v) == name select v).First();
                name = GetResourceTextFile(name);
                //ReportDocument rd = new ReportDocument ();
                // rd.Load (Path );
                // name = rd.SummaryInfo .ReportTitle ;
                name = "";
                #endregion
                AddReport(Path, name, reptype);

                System.IO.File.Delete(Path);
            }
        }
        public void UploadReports(string FolderPath)
        {
            var files = Directory.GetFiles(FolderPath);
            foreach (var Path in files)
            {
                var Name = "";
                try
                {
                    //  ReportDocument rd = new ReportDocument();
                    //   rd.Load(Path);
                    // Name = rd.SummaryInfo.ReportTitle;
                }
                catch (Exception ex) { ex.AppendInLogFile(); }
                AddReport(Path, Name, "report");
            }
        }
        public void UploadPLDs(string FolderPath)
        {
            var files = Directory.GetFiles(FolderPath);
            foreach (var Path in files)
            {
                var Name = "";
                try
                {
                    //  ReportDocument rd = new ReportDocument();
                    //   rd.Load(Path);
                    // Name = rd.SummaryInfo.ReportTitle;
                }
                catch (Exception ex) { ex.AppendInLogFile(); }
                if (System.IO.Path.GetExtension(Path) == ".rpt")
                {
                    var objectTypePath = FolderPath + "\\" + System.IO.Path.GetFileNameWithoutExtension(Path) + ".txt";
                    var objectType = System.IO.File.ReadAllText(objectTypePath);
                    AddPLD(Path, Name, objectType);
                }
            }

        }
        private string getNamespacePart(String Path)
        {
            var res = Path.Split('.');
            var reply = "";
            for (int i = 0; i < res.Length - 2; i++)
            {
                reply += "." + res[i];
            }
            return reply.Substring(1);
        }
        private string getFileName(String Path)
        {
            var str = Path.Split('.');
            return str[str.Length - 2] + "." + str[str.Length - 1];
            //return Assembly.GetEntryAssembly().GetManifestResourceInfo(Path).FileName;
        }

        private string GetResourceTextFile(String Namespace)
        {
            string result = string.Empty;

            using (Stream stream = Assembly.GetEntryAssembly().
                        GetManifestResourceStream(Namespace))
            {
                using (StreamReader sr = new StreamReader(stream))
                {
                    result = sr.ReadToEnd();
                }
            }
            return result;
        }

        private void AddPLD(string rptFilePath, string ReportName, String ObjType)
        {
            ReportLayoutsService oLayoutService = (ReportLayoutsService)oCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService);
            ReportLayout oReport = (ReportLayout)oLayoutService.GetDataInterface(ReportLayoutsServiceDataInterfaces.rlsdiReportLayout);

            //Initialize critical properties 
            // 
            // Use TypeCode "RCRI" to specify a Crystal Report. 
            // Use other TypeCode to specify a layout for a document type. 
            // List of TypeCode types are in table RTYP. 
            if (string.IsNullOrEmpty(ReportName))
                ReportName = System.IO.Path.GetFileNameWithoutExtension(rptFilePath);

            oReport.Name = ReportName; //"BonusReport";
            oReport.TypeCode = ObjType;
            oReport.Author = oCompany.UserName;
            oReport.Category = ReportLayoutCategoryEnum.rlcCrystal;


            string newReportCode;
            try
            {
                // Add new object 

                ReportLayoutParams oNewReportParams = oLayoutService.AddReportLayout(oReport);


                // Get code of the added ReportLayout object 
                newReportCode = oNewReportParams.LayoutCode;
                Console.WriteLine(newReportCode);
            }

            catch (System.Exception err)
            {
                string errMessage = err.Message;
                return;
            }

            // Wpload .rpt file using SetBlob interface 
            //string rptFilePath = @"D:\AbacusDev\PayRoll\Source\External Reports\Employee Bonus Report\BonusReport_Rpt.rpt";

            CompanyService oCompanyService = oCompany.GetCompanyService();
            // Specify the table and _FORM to update 
            BlobParams oBlobParams = (BlobParams)oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams);
            oBlobParams.Table = "RDOC";
            oBlobParams.Field = "Template";

            // Specify the record whose blob _FORM is to be set 
            BlobTableKeySegment oKeySegment = oBlobParams.BlobTableKeySegments.Add();
            oKeySegment.Name = "DocCode";
            oKeySegment.Value = newReportCode;

            Blob oBlob = (Blob)oCompanyService.GetDataInterface(CompanyServiceDataInterfaces.csdiBlob);

            // Put the rpt file into buffer 
            FileStream oFile = new FileStream(rptFilePath, System.IO.FileMode.Open);
            int fileSize = (int)oFile.Length;
            byte[] buf = new byte[fileSize];
            oFile.Read(buf, 0, fileSize);
            oFile.Close();

            // Convert memory buffer to Base64 string 
            oBlob.Content = Convert.ToBase64String(buf, 0, fileSize);

            try
            {
                //Upload Blob to database 

                oCompanyService.SetBlob(oBlobParams, oBlob);
            }
            catch (System.Exception ex)
            {
                string errmsg = ex.Message;
            }

        }
        private void AddReport(string rptFilePath, string ReportName, string reptype)
        {
            ReportLayoutsService oLayoutService = (ReportLayoutsService)oCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService);
            ReportLayout oReport = (ReportLayout)oLayoutService.GetDataInterface(ReportLayoutsServiceDataInterfaces.rlsdiReportLayout);

            //Initialize critical properties 
            // 
            // Use TypeCode "RCRI" to specify a Crystal Report. 
            // Use other TypeCode to specify a layout for a document type. 
            // List of TypeCode types are in table RTYP. 
            if (string.IsNullOrEmpty(ReportName))
                ReportName = System.IO.Path.GetFileNameWithoutExtension(rptFilePath);

            oReport.Name = ReportName; //"BonusReport";
            oReport.TypeCode = "RCRI";
            oReport.Author = oCompany.UserName;
            if (reptype.ToLower() == "report")
                oReport.Category = ReportLayoutCategoryEnum.rlcCrystal;
            else oReport.Category = ReportLayoutCategoryEnum.rlcPLD;


            string newReportCode;
            try
            {
                // Add new object 

                ReportLayoutParams oNewReportParams = oLayoutService.AddReportLayout(oReport);


                // Get code of the added ReportLayout object 
                newReportCode = oNewReportParams.LayoutCode;
            }

            catch (System.Exception err)
            {
                string errMessage = err.Message;
                return;
            }

            // Wpload .rpt file using SetBlob interface 
            //string rptFilePath = @"D:\AbacusDev\PayRoll\Source\External Reports\Employee Bonus Report\BonusReport_Rpt.rpt";

            CompanyService oCompanyService = oCompany.GetCompanyService();
            // Specify the table and _FORM to update 
            BlobParams oBlobParams = (BlobParams)oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams);
            oBlobParams.Table = "RDOC";
            oBlobParams.Field = "Template";

            // Specify the record whose blob _FORM is to be set 
            BlobTableKeySegment oKeySegment = oBlobParams.BlobTableKeySegments.Add();
            oKeySegment.Name = "DocCode";
            oKeySegment.Value = newReportCode;

            Blob oBlob = (Blob)oCompanyService.GetDataInterface(CompanyServiceDataInterfaces.csdiBlob);

            // Put the rpt file into buffer 
            FileStream oFile = new FileStream(rptFilePath, System.IO.FileMode.Open);
            int fileSize = (int)oFile.Length;
            byte[] buf = new byte[fileSize];
            oFile.Read(buf, 0, fileSize);
            oFile.Close();

            // Convert memory buffer to Base64 string 
            oBlob.Content = Convert.ToBase64String(buf, 0, fileSize);

            try
            {
                //Upload Blob to database 

                oCompanyService.SetBlob(oBlobParams, oBlob);
            }
            catch (System.Exception ex)
            {
                string errmsg = ex.Message;
            }

        }
    }
}
