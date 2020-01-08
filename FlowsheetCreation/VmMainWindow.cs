using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Input;
using Aucotec.EngineeringBase.Client.Runtime;

namespace FlowsheetCreation
{
    public class VmMainWindow : Helpers.VmBase
    {
        public string sSymbolOID;
        public List<StenFunc> SavedFol, SavedFunc;
        public List<Sheet> SavedStencil;
        public ObjectCollection sheets;
        public ObjectItem Function, Function2, Function3, oSel;
        public static ICommand CmdOpen { get; set; }
        public ObjectCollection FindStencil;
        public ObjectCollection FindStencil2;
        public Application myApplication;
        public IList<ExecuteSheetOperationRecordData> atParamData;

        public VmMainWindow(Application myApplication)
        {
            
            CmdOpen = new Helpers.RelayCommand(ExeOpen);
            this.myApplication = myApplication;
            SavedStencil = new List<Sheet>();
            SavedFol = new List<StenFunc>();
            SavedFunc = new List<StenFunc>();
            SecondWindow secondWindow = new SecondWindow();
            secondWindow.DataContext = new VmSecondWindow(myApplication, secondWindow);
            secondWindow.ShowDialog();

        }
        public class StenFunc
        {
            public ObjectItem FuncName;
            public ObjectItem StenName;
        }
        private void ExeOpen(object obj)
        {

       
        }
    }
}







