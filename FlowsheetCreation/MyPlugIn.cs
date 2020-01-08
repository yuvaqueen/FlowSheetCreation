using System;
using System.AddIn;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Interop;
using System.Windows.Threading;
using Aucotec.EngineeringBase.Client.Runtime;

namespace FlowsheetCreation
{
    /// <summary>
    /// Implements Wizard FlowsheetCreation
    /// </summary>
    [AddIn("FlowsheetCreation", Description = "FlowsheetCreation", Publisher = "Yuvarani")]
    public class MyPlugIn : PlugInWizard
    {
        /// <summary>
        /// Runs the wizard.
        /// </summary>
        /// <param name="myApplication">Application object instance</param>	
        public override void Run(Application myApplication)
        {
          
          //  if (myApplication.Selection[0].TypeId == ObjectType.FctPlant)
          //  {
               //  Run(myApplication);
           // }
           // else
           // {
                MainWindow frm = new MainWindow();
                frm.DataContext = new VmMainWindow(myApplication);
                WindowInteropHelper wih = new WindowInteropHelper(frm);
                wih.Owner = myApplication.ActiveWindow.Handle;
                // frm.ShowDialog();

                // Make a synchronously shutdown
                if (!AppDomain.CurrentDomain.IsDefaultAppDomain())
                    Dispatcher.CurrentDispatcher.InvokeShutdown();
           // }
        }

       
    }
}

