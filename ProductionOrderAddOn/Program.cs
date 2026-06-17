using System;
using System.Collections.Generic;
using ProductionOrderAddOn.Services;
using SAPbouiCOM.Framework;

namespace ProductionOrderAddOn
{
    class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        public static SAPbouiCOM.Application SBO_Application;

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;

                if (args.Length < 1)
                {
                    oApp = new Application();
                    SBO_Application = Application.SBO_Application;
                }
                else
                {
                    //If you want to use an add-on identifier for the development license, you can specify an add-on identifier string as the second parameter.
                    //oApp = new Application(args[0], "XXXXX");
                    oApp = new Application(args[0]);
                }

                DatabaseInitializer.Init();

                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    CompanyService.Disconnect();
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    CompanyService.Disconnect();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    CompanyService.Disconnect();
                    break;
                default:
                    break;
            }
        }
    }
}
