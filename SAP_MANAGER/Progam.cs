using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

using SAPbouiCOM.Framework;
namespace SAP_MANAGER
{

    class Program
    {
        public static SAPbouiCOM.Application oapp;
        public static SAPbobsCOM.Company oCompany;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new Application();
                }
                else
                {
                    oApp = new Application(args[0]);
                }
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                oapp = Application.SBO_Application;
                oCompany = oapp.Company.GetDICompany() as SAPbobsCOM.Company;
                Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;
                System.Windows.Forms.Application.Run();


            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            Application.SBO_Application.ItemEvent -= SBO_Application_ItemEvent;
            try
            {
                var item = "10000112";
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == item  )
                {
                    // Your code here 
                    var form = Application.SBO_Application.Forms.Item(pVal.FormUID).Items.Item(item).Visible = false;
                    BubbleEvent = false;
                }
            }
            catch (Exception ex)
            {
         
            }
            finally
            {
                Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;
            }
        }

        private static bool IsAbitUser()
        {
            // Windows logon name only (case‑insensitive)
            return string.Equals(Environment.UserName, "abit", StringComparison.OrdinalIgnoreCase);

            // ‑‑ If you need the full DOMAIN\user form, use:
            // var fullName = WindowsIdentity.GetCurrent().Name;  
            // return fullName.EndsWith(@"\abit", StringComparison.OrdinalIgnoreCase);
        }
        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    {
                        if (!IsAbitUser())
                        {


                            var item = "10000112";
                            Application.SBO_Application.Forms.GetForm("0", 0).Items.Item(item).Visible = false;
                            Application.SBO_Application.Menus.Item("43564").Enabled = false;
                            Application.SBO_Application.Menus.RemoveEx("43524");
                            Application.SBO_Application.Menus.RemoveEx("43523");
                            //  sssApplication.SBO_Application.Menus.RemoveEx("43564");
                        }
                    }
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }
    }
}
