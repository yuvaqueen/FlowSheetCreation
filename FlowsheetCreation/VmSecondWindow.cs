using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows.Data;
using System.Windows.Input;
using System.Linq;
using Aucotec.EngineeringBase.Client.Runtime;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;
using FlowsheetCreation.Model;

namespace FlowsheetCreation
{
    public class VmSecondWindow : Helpers.VmBase
    {

        private string _cirText;
        public string CirText
        {
            get
            {
                return _cirText;
            }
            set
            {
                _cirText = value;
                OnPropertyChanged("CirText");
            }
        }
        public List<ObjectItem> funcNames;
        public System.Windows.Forms.ComboBox combobox1;
        public ObjectItem Function, Function2, Function3, oSel, Folder;
        public ObjectCollection sheets, FindStencil2, FindStencil;
        public Application myApplication;
        public ModelView MySelectedItem { get; set; }
        public ModelView SelTargetSheet { get; set; }
        public ModelView SelPosition { get; set; }
        public ModelView SelCir { get; set; }
        private ObservableCollection<ModelView> _MyData;
        public ObservableCollection<ModelView> MyData
        {
            get { return _MyData; }
            set
            {
                _MyData = value;
                OnPropertyChanged("MyData");
            }
        }
        public string partOf, circuitcomp, position, sSymbolOID;
        private ObservableCollection<ModelView> _position;
        public ObservableCollection<ModelView> Position
        {
            get { return _position; }
            set
            {
                _position = value;
                OnPropertyChanged("Position");
            }
        }
        private ObservableCollection<ModelView> _circomp;
        public ObservableCollection<ModelView> Circomp
        {
            get { return _circomp; }
            set
            {
                _circomp = value;
                OnPropertyChanged("Circomp");
            }
        }
        private ObservableCollection<ModelView> _targetSheet;
        public ObservableCollection<ModelView> TargetSheet
        {
            get { return _targetSheet; }
            set
            {
                _targetSheet = value;
                OnPropertyChanged("TargetSheet");
            }
        }
        public List<WorksheetRow> worksheetitems;
        public static ICommand CmdOpen { get; set; }
        public static ICommand CmdCancel { get; set; }
        public List<StenFunc> SavedFol, SavedFunc;
        public List<GatherInfo> SelectedFunctions;
        public GatherInfo GatherIdObject;
        public List<GatherObjRow> toTest;
        public int InformationOnname = 12199;
        public int Find2 = 101367;
        public int Find3 = 25;
        public int Dimension = 27434;
        public class StenFunc
        {
            public ObjectItem FuncName;
            public ObjectItem StenName;
        }
        public class GatherObjRow
        {
            public WorksheetRow RowValue;
            public ObjectItem StencilValue;
        }
        public class GatherInfo
        {
            public Sheet TargetSheet;
            public ObjectItem CircuitComp;
            public ObjectItem StencilComp;
            public ObjectItem FetchedStencil;
            public ObjectItem BrokenComponent;
            public string Pos;
            public Boolean IsPlaced = false;
        }
        public IList<ExecuteSheetOperationRecordData> atParamData;
        public List<Sheet> SavedStencil;
        public SecondWindow secondWindow;
        public void upPos()
        {
            if (MySelectedItem != null && MySelectedItem.SelPosition != null)
                MySelectedItem.PosText = MySelectedItem.SelPosition.Type2Name.ToString();

            if (MySelectedItem != null && MySelectedItem.SelTargetSheet != null)
                MySelectedItem.SheetText = MySelectedItem.SelTargetSheet.Type3Name.ToString();

            if (MySelectedItem != null && MySelectedItem.SelCir != null)
                MySelectedItem.CirText = MySelectedItem.SelCir.Type1Name.ToString();

        }

        public VmSecondWindow(Application myApplication, SecondWindow secondWindow)
        {

            toTest = new List<GatherObjRow>();
            funcNames = new List<ObjectItem>();
            this.secondWindow = secondWindow;
            SelectedFunctions = new List<GatherInfo>();
            SavedStencil = new List<Sheet>();
            SavedFol = new List<StenFunc>();
            SavedFunc = new List<StenFunc>();
            this.myApplication = myApplication;
            Circomp = new ObservableCollection<ModelView>();
            Position = new ObservableCollection<ModelView>();
            TargetSheet = new ObservableCollection<ModelView>();
            CmdOpen = new Helpers.RelayCommand(ExeOpen);
            CmdCancel = new Helpers.RelayCommand(ExeCancel);
            sheets = myApplication.ActiveProject.DrawingsFolder.FindObjects(ObjectKind.Sheet, SearchBehavior.Deep);
            FindStencil = myApplication.Folders.Stencils.FindObjects(ObjectKind.StencilCircuitComponent, SearchBehavior.Deep);

            var findstencil = from x in FindStencil from y in x.Children select y;
            findstencil.ToList().ForEach(items => { Circomp.Add(new ModelView() { TypeIcon = items.Image, Source = items, Type1Name = items.Name }); });

            foreach (Sheet item in sheets)
            {
                item.ExecuteFormula("A11544;", out partOf);
                TargetSheet.Add(new ModelView() { SourceSheet = item, TypeIcon = item.Image, Source = item, Type3Name = partOf + " " + item.Name });

            }
            #region POSITION

            Position.Add(new ModelView() { Type2Name = "A1" });
            Position.Add(new ModelView() { Type2Name = "A2" });
            Position.Add(new ModelView() { Type2Name = "A3" });
            Position.Add(new ModelView() { Type2Name = "A4" });
            Position.Add(new ModelView() { Type2Name = "A5" });
            Position.Add(new ModelView() { Type2Name = "A6" });
            Position.Add(new ModelView() { Type2Name = "A7" });
            Position.Add(new ModelView() { Type2Name = "A8" });
            Position.Add(new ModelView() { Type2Name = "A9" });
            Position.Add(new ModelView() { Type2Name = "A10" });
            Position.Add(new ModelView() { Type2Name = "A11" });
            Position.Add(new ModelView() { Type2Name = "A12" });
            Position.Add(new ModelView() { Type2Name = "A13" });
            Position.Add(new ModelView() { Type2Name = "A14" });
            Position.Add(new ModelView() { Type2Name = "A15" });
            Position.Add(new ModelView() { Type2Name = "A16" });
            Position.Add(new ModelView() { Type2Name = "B1" });
            Position.Add(new ModelView() { Type2Name = "B2" });
            Position.Add(new ModelView() { Type2Name = "B3" });
            Position.Add(new ModelView() { Type2Name = "B4" });
            Position.Add(new ModelView() { Type2Name = "B5" });
            Position.Add(new ModelView() { Type2Name = "B6" });
            Position.Add(new ModelView() { Type2Name = "B7" });
            Position.Add(new ModelView() { Type2Name = "B8" });
            Position.Add(new ModelView() { Type2Name = "B9" });
            Position.Add(new ModelView() { Type2Name = "B10" });
            Position.Add(new ModelView() { Type2Name = "B11" });
            Position.Add(new ModelView() { Type2Name = "B12" });
            Position.Add(new ModelView() { Type2Name = "B13" });
            Position.Add(new ModelView() { Type2Name = "B14" });
            Position.Add(new ModelView() { Type2Name = "B15" });
            Position.Add(new ModelView() { Type2Name = "B16" });
            Position.Add(new ModelView() { Type2Name = "C1" });
            Position.Add(new ModelView() { Type2Name = "C2" });
            Position.Add(new ModelView() { Type2Name = "C3" });
            Position.Add(new ModelView() { Type2Name = "C4" });
            Position.Add(new ModelView() { Type2Name = "C5" });
            Position.Add(new ModelView() { Type2Name = "C6" });
            Position.Add(new ModelView() { Type2Name = "C7" });
            Position.Add(new ModelView() { Type2Name = "C8" });
            Position.Add(new ModelView() { Type2Name = "C9" });
            Position.Add(new ModelView() { Type2Name = "C10" });
            Position.Add(new ModelView() { Type2Name = "C11" });
            Position.Add(new ModelView() { Type2Name = "C12" });
            Position.Add(new ModelView() { Type2Name = "C13" });
            Position.Add(new ModelView() { Type2Name = "C14" });
            Position.Add(new ModelView() { Type2Name = "C15" });
            Position.Add(new ModelView() { Type2Name = "C16" });
            Position.Add(new ModelView() { Type2Name = "D1" });
            Position.Add(new ModelView() { Type2Name = "D2" });
            Position.Add(new ModelView() { Type2Name = "D3" });
            Position.Add(new ModelView() { Type2Name = "D4" });
            Position.Add(new ModelView() { Type2Name = "D5" });
            Position.Add(new ModelView() { Type2Name = "D6" });
            Position.Add(new ModelView() { Type2Name = "D7" });
            Position.Add(new ModelView() { Type2Name = "D8" });
            Position.Add(new ModelView() { Type2Name = "D9" });
            Position.Add(new ModelView() { Type2Name = "D10" });
            Position.Add(new ModelView() { Type2Name = "D11" });
            Position.Add(new ModelView() { Type2Name = "D12" });
            Position.Add(new ModelView() { Type2Name = "D13" });
            Position.Add(new ModelView() { Type2Name = "D14" });
            Position.Add(new ModelView() { Type2Name = "D15" });
            Position.Add(new ModelView() { Type2Name = "D16" });
            Position.Add(new ModelView() { Type2Name = "E1" });
            Position.Add(new ModelView() { Type2Name = "E2" });
            Position.Add(new ModelView() { Type2Name = "E3" });
            Position.Add(new ModelView() { Type2Name = "E4" });
            Position.Add(new ModelView() { Type2Name = "E5" });
            Position.Add(new ModelView() { Type2Name = "E6" });
            Position.Add(new ModelView() { Type2Name = "E7" });
            Position.Add(new ModelView() { Type2Name = "E8" });
            Position.Add(new ModelView() { Type2Name = "E9" });
            Position.Add(new ModelView() { Type2Name = "E10" });
            Position.Add(new ModelView() { Type2Name = "E11" });
            Position.Add(new ModelView() { Type2Name = "E12" });
            Position.Add(new ModelView() { Type2Name = "E13" });
            Position.Add(new ModelView() { Type2Name = "E14" });
            Position.Add(new ModelView() { Type2Name = "E15" });
            Position.Add(new ModelView() { Type2Name = "E16" });
            Position.Add(new ModelView() { Type2Name = "F1" });
            Position.Add(new ModelView() { Type2Name = "F2" });
            Position.Add(new ModelView() { Type2Name = "F3" });
            Position.Add(new ModelView() { Type2Name = "F4" });
            Position.Add(new ModelView() { Type2Name = "F5" });
            Position.Add(new ModelView() { Type2Name = "F6" });
            Position.Add(new ModelView() { Type2Name = "F7" });
            Position.Add(new ModelView() { Type2Name = "F8" });
            Position.Add(new ModelView() { Type2Name = "F9" });
            Position.Add(new ModelView() { Type2Name = "F10" });
            Position.Add(new ModelView() { Type2Name = "F11" });
            Position.Add(new ModelView() { Type2Name = "F12" });
            Position.Add(new ModelView() { Type2Name = "F13" });
            Position.Add(new ModelView() { Type2Name = "F14" });
            Position.Add(new ModelView() { Type2Name = "F15" });
            Position.Add(new ModelView() { Type2Name = "F16" });
            Position.Add(new ModelView() { Type2Name = "G1" });
            Position.Add(new ModelView() { Type2Name = "G2" });
            Position.Add(new ModelView() { Type2Name = "G3" });
            Position.Add(new ModelView() { Type2Name = "G4" });
            Position.Add(new ModelView() { Type2Name = "G5" });
            Position.Add(new ModelView() { Type2Name = "G6" });
            Position.Add(new ModelView() { Type2Name = "G7" });
            Position.Add(new ModelView() { Type2Name = "G8" });
            Position.Add(new ModelView() { Type2Name = "G9" });
            Position.Add(new ModelView() { Type2Name = "G10" });
            Position.Add(new ModelView() { Type2Name = "G11" });
            Position.Add(new ModelView() { Type2Name = "G12" });
            Position.Add(new ModelView() { Type2Name = "G13" });
            Position.Add(new ModelView() { Type2Name = "G14" });
            Position.Add(new ModelView() { Type2Name = "G15" });
            Position.Add(new ModelView() { Type2Name = "G16" });
            Position.Add(new ModelView() { Type2Name = "H1" });
            Position.Add(new ModelView() { Type2Name = "H2" });
            Position.Add(new ModelView() { Type2Name = "H3" });
            Position.Add(new ModelView() { Type2Name = "H4" });
            Position.Add(new ModelView() { Type2Name = "H5" });
            Position.Add(new ModelView() { Type2Name = "H6" });
            Position.Add(new ModelView() { Type2Name = "H7" });
            Position.Add(new ModelView() { Type2Name = "H8" });
            Position.Add(new ModelView() { Type2Name = "H9" });
            Position.Add(new ModelView() { Type2Name = "H10" });
            Position.Add(new ModelView() { Type2Name = "H11" });
            Position.Add(new ModelView() { Type2Name = "H12" });
            Position.Add(new ModelView() { Type2Name = "H13" });
            Position.Add(new ModelView() { Type2Name = "H14" });
            Position.Add(new ModelView() { Type2Name = "H15" });
            Position.Add(new ModelView() { Type2Name = "H16" });
            Position.Add(new ModelView() { Type2Name = "J1" });
            Position.Add(new ModelView() { Type2Name = "J2" });
            Position.Add(new ModelView() { Type2Name = "J3" });
            Position.Add(new ModelView() { Type2Name = "J4" });
            Position.Add(new ModelView() { Type2Name = "J5" });
            Position.Add(new ModelView() { Type2Name = "J6" });
            Position.Add(new ModelView() { Type2Name = "J7" });
            Position.Add(new ModelView() { Type2Name = "J8" });
            Position.Add(new ModelView() { Type2Name = "J9" });
            Position.Add(new ModelView() { Type2Name = "J10" });
            Position.Add(new ModelView() { Type2Name = "J11" });
            Position.Add(new ModelView() { Type2Name = "J12" });
            Position.Add(new ModelView() { Type2Name = "J13" });
            Position.Add(new ModelView() { Type2Name = "J14" });
            Position.Add(new ModelView() { Type2Name = "J15" });
            Position.Add(new ModelView() { Type2Name = "J16" });
            Position.Add(new ModelView() { Type2Name = "K1" });
            Position.Add(new ModelView() { Type2Name = "K2" });
            Position.Add(new ModelView() { Type2Name = "K3" });
            Position.Add(new ModelView() { Type2Name = "K4" });
            Position.Add(new ModelView() { Type2Name = "K5" });
            Position.Add(new ModelView() { Type2Name = "K6" });
            Position.Add(new ModelView() { Type2Name = "K7" });
            Position.Add(new ModelView() { Type2Name = "K8" });
            Position.Add(new ModelView() { Type2Name = "K9" });
            Position.Add(new ModelView() { Type2Name = "K10" });
            Position.Add(new ModelView() { Type2Name = "K11" });
            Position.Add(new ModelView() { Type2Name = "K12" });
            Position.Add(new ModelView() { Type2Name = "K13" });
            Position.Add(new ModelView() { Type2Name = "K14" });
            Position.Add(new ModelView() { Type2Name = "K15" });
            Position.Add(new ModelView() { Type2Name = "K16" });
            Position.Add(new ModelView() { Type2Name = "L1" });
            Position.Add(new ModelView() { Type2Name = "L2" });
            Position.Add(new ModelView() { Type2Name = "L3" });
            Position.Add(new ModelView() { Type2Name = "L4" });
            Position.Add(new ModelView() { Type2Name = "L5" });
            Position.Add(new ModelView() { Type2Name = "L6" });
            Position.Add(new ModelView() { Type2Name = "L7" });
            Position.Add(new ModelView() { Type2Name = "L8" });
            Position.Add(new ModelView() { Type2Name = "L9" });
            Position.Add(new ModelView() { Type2Name = "L10" });
            Position.Add(new ModelView() { Type2Name = "L11" });
            Position.Add(new ModelView() { Type2Name = "L12" });
            Position.Add(new ModelView() { Type2Name = "L13" });
            Position.Add(new ModelView() { Type2Name = "L14" });
            Position.Add(new ModelView() { Type2Name = "L15" });
            Position.Add(new ModelView() { Type2Name = "L16" });
            Position.Add(new ModelView() { Type2Name = "M1" });
            Position.Add(new ModelView() { Type2Name = "M2" });
            Position.Add(new ModelView() { Type2Name = "M3" });
            Position.Add(new ModelView() { Type2Name = "M4" });
            Position.Add(new ModelView() { Type2Name = "M5" });
            Position.Add(new ModelView() { Type2Name = "M6" });
            Position.Add(new ModelView() { Type2Name = "M7" });
            Position.Add(new ModelView() { Type2Name = "M8" });
            Position.Add(new ModelView() { Type2Name = "M9" });
            Position.Add(new ModelView() { Type2Name = "M10" });
            Position.Add(new ModelView() { Type2Name = "M11" });
            Position.Add(new ModelView() { Type2Name = "M12" });
            Position.Add(new ModelView() { Type2Name = "M13" });
            Position.Add(new ModelView() { Type2Name = "M14" });
            Position.Add(new ModelView() { Type2Name = "M15" });
            Position.Add(new ModelView() { Type2Name = "M16" });
            #endregion
            MyData = new ObservableCollection<ModelView>();

            myApplication.Selection.ToList().ForEach(item => {
                item.ExecuteFormula("A22;", out partOf);
                item.ExecuteFormula("A27434;", out position);
                item.ExecuteFormula("A12199;", out circuitcomp); //A12199
                MyData.Add(new ModelView()
                {
                    Source = item,
                    PartOf = partOf,
                    Designation = item.Name,
                    Position = Position,
                    Circomp = Circomp,
                    CirName = partOf,
                    TargetSheet = TargetSheet,
                    Checkbox = true,
                    PosText = position,
                    CirText = circuitcomp,
                    SheetText = "Please Select"
                });
            });
        }

        private void ExeCancel(object obj)
        {
            secondWindow.Close();
        }

        public void ExeOpen(object obj)
        {
            using (WaitDialog wait = myApplication.Dialogs.CreateWaitDialog())   //creates waiting dialog while generating the report. 
            {
                Stopwatch sw = Stopwatch.StartNew();
              //  List<ObjectItem> ItemsWithNoPoint = new List<ObjectItem>();

                PlantModel PlantModel = new PlantModel();
                PlantModel.ExistingPlants = GetAllPlantFromProject(myApplication);

                wait.ShowDialogAsync(1, "Please wait...", AnimationType.FileMove);


                var selecteditem = from ModelView MySelectedItem in MyData where MySelectedItem.Checkbox == true where MySelectedItem.CirText != null && MySelectedItem.PosText != null && MySelectedItem.SheetText != null select MySelectedItem;

                selecteditem.ToList().ForEach(item =>
                {
                    var FindStencilChild1 = from x in FindStencil from y in x.Children where y.Name == item.CirText select y;
                    SelectedFunctions.AddRange((from x in FindStencilChild1 select new GatherInfo { Pos = item.PosText, CircuitComp = x, TargetSheet = item.SelTargetSheet.SourceSheet, StencilComp = item.Source }));
                });

                List<GatherInfo> DistinctSelectedFunctions = SelectedFunctions.Distinct().ToList();

                DistinctSelectedFunctions = ChangeInfoAndDimensionAttribute(DistinctSelectedFunctions);
                // fetch stencil 
                var FecthList = StencilBySelectedItem(DistinctSelectedFunctions);

                var SheetGrouped = from records in FecthList
                                   group records by records.TargetSheet.FullName into g
                                   where g.Count() >= 1
                                   select g.Key;

                var sheetsWithNoRef = (from records in FecthList where records.TargetSheet.Attributes.GetAttributeValue(AttributeId.SheetSize) != "A0" && records.TargetSheet.Attributes.GetAttributeValue(AttributeId.SheetSize) != "A1" && records.TargetSheet.Attributes.GetAttributeValue(AttributeId.SheetSize) != "A2" && records.TargetSheet.Attributes.GetAttributeValue(AttributeId.SheetSize) != "A3"
                                   select records.TargetSheet.Name).ToList();

                if (sheetsWithNoRef != null && sheetsWithNoRef.Count != 0)
                {
                    FecthList.ForEach(item => { if (sheetsWithNoRef.Contains(item.TargetSheet.Name)) { FecthList.RemoveAll(x => x == item); } });
                    var removedDup = sheetsWithNoRef.Distinct();
                    foreach (var item in removedDup)
                    {
                        System.Windows.MessageBox.Show("Please Add Sheet Size to" + item, "Sheet Size Undefined", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Stop);
                    }
                }
                else
                {
                    var SetFold = SetFoldForDevice(FecthList, SavedFol);

                    foreach (var name in SheetGrouped)
                    {
                        var group = FecthList.Where(x => x.TargetSheet.FullName == name).ToList();
                        var firstItem = group.FirstOrDefault();
                        firstItem.TargetSheet.Open(SheetOpenBehavior.AutoSave);


                        foreach (var Object in group)
                        {
                            GatherInfo PlacedObject = InserObjectToSheetAndCollectData(Object);
                            var toBedeleted = SavedFol.Where(x => x.FuncName.Id.Equals(PlacedObject.StencilComp.Id)).FirstOrDefault();
                            if (PlacedObject.IsPlaced == true)
                            {
                                PlantModel.NewPlants = GetAllPlantFromProject(myApplication);
                                //var PlantsFileterd = (from u in PlantModel.NewPlants
                                //                      from x in PlantModel.ExistingPlants
                                //                      where u.Id != x.Id
                                //                      select u).ToList();
                                List<ObjectItem> PlantsFileterd = new List<ObjectItem>();
                                foreach (var plant in PlantModel.NewPlants)
                                {
                                    if (!PlantModel.ExistingPlants.Contains(plant))
                                        PlantsFileterd.Add(plant);
                                }
                                string str = PlacedObject.StencilComp.Name.Substring(0, 3);
                                var CorrectPlants = PlantsFileterd.Where(x => x.Children.ToList().Any(c => c.Name.Contains(str))).FirstOrDefault();

                                //   if (CorrectPlants == null)
                                //   ItemsWithNoPoint.Add(Object.StencilComp);

                                if (CorrectPlants != null)
                                {
                                    var ChildrenOfPlantToDelete = PlantsFileterd.Where(c => c.Id != CorrectPlants.Id).Select(v => v.Children).ToList();
                                    foreach (var childrens in ChildrenOfPlantToDelete)
                                    {
                                        DeleteInSheet(childrens.ToList(), PlacedObject.TargetSheet);
                                    }
                                }
                                var exisitngObjInPlant = GetAllIdOfCollectionExculdeFolder(PlacedObject.StencilComp.Parent.Children);

                                BreakUpComponent(PlacedObject, CorrectPlants);

                                var NewUpdateOfPlant = GetAllIdOfCollectionExculdeFolder(PlacedObject.StencilComp.Parent.Children);
                                var NewUpdateOfPlantNoFolder = from i in PlacedObject.StencilComp.Parent.Children where i.Kind != ObjectKind.Folder select i;

                                NewUpdateOfPlant.RemoveAll(x => exisitngObjInPlant.Contains(x));

                                var ToBeRemoved = NewUpdateOfPlantNoFolder.ToList().Where(x => NewUpdateOfPlant.Contains(x.Id) && !x.Name.Equals(PlacedObject.StencilComp.Name)).ToList();
                                DeleteInSheet(ToBeRemoved, PlacedObject.TargetSheet);
                                ToBeRemoved.ForEach(item =>
                            {
                                item.Delete();
                            });

                                var NewUpdateOfPlantAfterDel = PlacedObject.StencilComp.Parent.Children.ToList();
                                var NewUpdateOfPlantAfterDelFDel = from i in NewUpdateOfPlantAfterDel where i.Kind != ObjectKind.Folder select i;

                                PlacedObject.FetchedStencil = (from x in NewUpdateOfPlantAfterDelFDel where x.Name == PlacedObject.StencilComp.Name && x.Id != PlacedObject.StencilComp.Id select x).FirstOrDefault();

                                var importantFold = SavedFol.Select(x => x).Where(f => f.StenName.Name == PlacedObject.StencilComp.Name).FirstOrDefault();
                                if (PlacedObject.FetchedStencil == null) { }

                                if (PlacedObject.FetchedStencil != null)
                                {

                                    var checksamechild = from fetchstencil in PlacedObject.FetchedStencil.Children from StencilCom in PlacedObject.StencilComp.Children where fetchstencil.Name == StencilCom.Name select new { fetchstencil, StencilCom };

                                    var selecteddata1 = PlacedObject.FetchedStencil.Children.ToList();

                                    PlacedObject.StencilComp.Children.ToList().ForEach(item => { selecteddata1.RemoveAll(x => x.Name.Contains(item.Name)); });

                                    var objectTodelete = (from check in selecteddata1 where PlacedObject.StencilComp.Name == check.Parent.Name select check).ToList();
                                    DeleteInSheet(objectTodelete, PlacedObject.TargetSheet);



                                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                    /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                                    checksamechild.ToList().ForEach(items => { items.fetchstencil.MoveTo(importantFold.StenName); });
                                    PlacedObject.FetchedStencil.Children.ToList().ForEach(child => { if (child.TypeId != ObjectType.PinHydraulicPneumatic) { child.Delete(); } });


                                    //merges attribute of two diffrent object 
                                    mergeAttribute(importantFold);

                                    importantFold.StenName.Children.ToList().ForEach(child => { child.MoveTo(PlacedObject.FetchedStencil); });

                                    //if (toBedeleted != null) {
                                    //    if (toBedeleted.FuncName != null)
                                    //        toBedeleted.FuncName.Delete();
                                    //    if (toBedeleted.StenName != null)
                                    //        toBedeleted.StenName.Delete();
                                    //        }
                                }
                            }
                            else
                            {
                                //if (toBedeleted != null)
                                //{
                                //    if (toBedeleted.StenName != null)
                                //        toBedeleted.StenName.Delete();
                                //}
                                //System.Windows.MessageBox.Show("Please have a reference name for sheet", "No Sheet Size", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Exclamation);
                            }

                        }
                        firstItem.TargetSheet.Close();
                    }

                    PlantModel.NewPlants = GetAllPlantFromProject(myApplication);

                    List<ObjectItem> PlantstoDelete = new List<ObjectItem>();
                    foreach (var plant in PlantModel.NewPlants)
                    {
                        if (!PlantModel.ExistingPlants.Contains(plant))
                            PlantstoDelete.Add(plant);
                    }
                    PlantstoDelete.ToList().ForEach(unwatedItem => { unwatedItem.Delete(); });
                    //foreach (var unwatedItem in PlantstoDelete)
                    //{
                    //    unwatedItem.Delete();
                    //}

                    //  var todelete = (from i in ItemsWithNoPoint from j in SavedFol where i.Id == j.FuncName.Id select j.FuncName.Id).ToList();
                    //   SavedFol.ToList().ForEach(item => {{ if (!todelete.Contains(item.FuncName.Id)) { item.FuncName.Delete(); } } item.StenName.Delete(); });
                    SavedFol.ForEach(x => { x.FuncName.Delete(); x.StenName.Delete(); });
                }
                sw.Stop(); 
                wait.CloseDialog();

                secondWindow.Close();
            }
        }
        private void DeleteInSheet(List<ObjectItem> ItemsToDelete, Sheet TargetSheet)
        {
            List<ObjectItem> AllChildren = new List<ObjectItem>();
            foreach (var item in ItemsToDelete)
            {
                AllChildren.Add(item);
                if (item.Children != null)
                {
                    foreach (var firstStack in item.Children)
                    {
                        AllChildren.Add(firstStack);
                        if (firstStack.Children != null)
                        {
                            foreach (var SecondStack in firstStack.Children)
                            {
                                AllChildren.Add(SecondStack);
                                if (SecondStack.Children != null)
                                {
                                    foreach (var ThirdStack in SecondStack.Children)
                                    {
                                        AllChildren.Add(ThirdStack);
                                        if (ThirdStack.Children != null)
                                        {
                                            foreach (var FourthStack in ThirdStack.Children)
                                            {
                                                AllChildren.Add(FourthStack);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            var test = AllChildren;
            AllChildren.ForEach(s =>
            {
                if (s.TypeId != ObjectType.PinHydraulicPneumatic)
                {

                    var sourceAssoc = from assoc in s.SourceAssociations select assoc;

                    sourceAssoc.ToList().ForEach(sa =>
                    {

                        sSymbolOID = sa.RelatedObject.Id;

                        atParamData = new List<ExecuteSheetOperationRecordData>();
                        atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.OpExecSheetDeleteSymbol, Value = 0 });
                        atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetRef2Symbol, Value = sSymbolOID });
                        TargetSheet.ExecuteSheetOperation(ref atParamData);
                        TargetSheet.Store();
                    });
                }
            });
        }

        private List<string> GetAllIdOfCollectionExculdeFolder(ObjectCollection collection)
        {
            return (from i in collection where i.Kind != ObjectKind.Folder select i.Id).ToList();
        }

        private List<GatherInfo> ChangeInfoAndDimensionAttribute(List<GatherInfo> DistinctSelectedFunctions)
        {
            DistinctSelectedFunctions.ToList().ForEach(item =>
            {
                item.StencilComp.Attributes.SetAttributeValue((AttributeId)InformationOnname, item.CircuitComp.Name);
                item.StencilComp.Attributes.SetAttributeValue((AttributeId)Dimension, item.Pos);
                item.StencilComp.Store();
            });
            if (DistinctSelectedFunctions.Count.ToString() == "0")
                System.Windows.MessageBox.Show("Please select the values");

            return DistinctSelectedFunctions;

        }

        private GatherInfo InserObjectToSheetAndCollectData(GatherInfo objectToBeplaced)
        {

            var syncDesig = objectToBeplaced.FetchedStencil.Attributes.GetAttributeValue(AttributeId.SymbolSyncDesignation);
            var masterId = objectToBeplaced.FetchedStencil.Parent.Id;
            string reference = masterId + "#" + syncDesig;
            atParamData = new List<ExecuteSheetOperationRecordData>();
            atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.OpExecSheetDropSymbol });
            atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetRef2Master, Value = reference });

            string value = objectToBeplaced.Pos;

            //10386
            int i = 0;
            if (objectToBeplaced.TargetSheet.Attributes.GetAttributeValue(AttributeId.SheetSize) == "A0")
            {
                i++;
                atParamData = sheetPlacementDestination(value, atParamData);
            }
            else if (objectToBeplaced.TargetSheet.Attributes.GetAttributeValue(AttributeId.SheetSize) == "A1")
            {
                i++;
                atParamData = sheetPlacementDestination1(value, atParamData);
            }
            else if (objectToBeplaced.TargetSheet.Attributes.GetAttributeValue(AttributeId.SheetSize) == "A2")
            {
                i++;
                atParamData = sheetPlacementDestination2(value, atParamData);
            }
            else if (objectToBeplaced.TargetSheet.Attributes.GetAttributeValue(AttributeId.SheetSize) == "A3")
            {
                i++;
                atParamData = sheetPlacementDestination3(value, atParamData);
            }

            if (i != 0)
            {
                atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetRef2Obj, Value = objectToBeplaced.FetchedStencil.Id });  //Motorshape.id ... this is the OID string.

                objectToBeplaced.TargetSheet.ExecuteSheetOperation(ref atParamData);
                objectToBeplaced.TargetSheet.Store();


                Worksheet wsf;

                wsf = objectToBeplaced.TargetSheet.OpenWorksheet("Functions", interactiveOnly: false);

                WorksheetRow LastEntry = null;

                worksheetitems = wsf.Rows.Where(s => s.ObjectItem.Parent.IsDeleted == false).Where(s => s.ObjectItem.Id == objectToBeplaced.StencilComp.Id).Select(s => s).ToList();

                LastEntry = worksheetitems.LastOrDefault();

                if (LastEntry != null)
                    objectToBeplaced.BrokenComponent = LastEntry.ObjectItem;

                objectToBeplaced.IsPlaced = true;
            }
            return objectToBeplaced;

        }
        private IList<ExecuteSheetOperationRecordData> sheetPlacementDestination(string positionChosen, IList<ExecuteSheetOperationRecordData> atParamData)
        {

            switch (positionChosen)
            {
                #region A
                case "A1":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 42 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 801 });
                    break;

                case "A2":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 111 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 801 });
                    break;

                case "A3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 186 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 801 });
                    break;


                case "A4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 260 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 801 });
                    break;


                case "A5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 334 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 801 });
                    break;


                case "A6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 408 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 801 });
                    break;


                case "A7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 483});
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 801 });
                    break;


                case "A8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 556 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 801 });
                    break;
                case "A9":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 632 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 801 });
                    break;

                case "A10":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 705 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 801 });
                    break;
                case "A11":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 780 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 801 });
                    break;
                case "A12":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 855 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 801 });
                    break;

                case "A13":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 928 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 801 });
                    break;

                case "A14":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1004 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 801 });
                    break;

                case "A15":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1076 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 801 });
                    break;

                case "A16":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1148 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 801 });
                    break;


                #endregion
                #region B
                case "B1":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 42 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 737 });
                    break;


                case "B2":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 112 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 737 });
                    break;


                case "B3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 185 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 737 });
                    break;


                case "B4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 259 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 737 });
                    break;


                case "B5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 335 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 737 });
                    break;

                case "B6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 409 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 737 });
                    break;


                case "B7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 483 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 737 });
                    break;


                case "B8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 557 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 737 });
                    break;
                case "B9":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 631 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 737 });
                    break;
                case "B10":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 705 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 737 });
                    break;
                case "B11":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 780 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 737 });
                    break;
                case "B12":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 855 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 737 });
                    break;

                case "B13":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 930 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 737 });
                    break;

                case "B14":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1003 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 737 });
                    break;

                case "B15":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1077 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 737 });
                    break;

                case "B16":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1147 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 737 });
                    break;
                #endregion
                #region C
                case "C1":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 42 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 665 });
                    break;


                case "C2":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 112 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 665 });
                    break;


                case "C3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 185 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 665 });
                    break;


                case "C4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 259 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 665 });
                    break;


                case "C5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 335 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 665 });
                    break;


                case "C6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 409 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 665 });
                    break;


                case "C7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 483 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 665 });

                    break;

                case "C8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 557 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 665 });

                    break;
                case "C9":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 631 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 665 });
                    break;
                case "C10":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 780 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 665 });
                    break;
                case "C11":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 780 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 665 });
                    break;
                case "C12":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 855 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 665 });
                    break;

                case "C13":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 930 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 665 });
                    break;

                case "C14":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1003 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 665 });
                    break;

                case "C15":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1077 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 665 });
                    break;

                case "C16":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1147 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 665 });
                    break;
                #endregion
                #region D

                case "D1":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 42 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 595 });
                    break;


                case "D2":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 112 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 595 });
                    break;


                case "D3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 185 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 595 });
                    break;


                case "D4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 259 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 595 });
                    break;


                case "D5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 335 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 595 });
                    break;


                case "D6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 409 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 595 });
                    break;


                case "D7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 483 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 595 });
                    break;


                case "D8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 557 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 595 });


                    break;
                case "D9":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 631 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 595 });
                    break;
                case "D10":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 705 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 595 });
                    break;
                case "D11":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 780 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 595 });
                    break;
                case "D12":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 855 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 595 });
                    break;

                case "D13":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 930 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 595 });
                    break;

                case "D14":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1003 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 595 });
                    break;

                case "D15":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1077 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 595 });
                    break;

                case "D16":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1147 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 595 });
                    break;
                #endregion
                #region E
                case "E1":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 42 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 525 });


                    break;
                case "E2":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 112 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 525 });


                    break;
                case "E3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 185 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 525 });


                    break;
                case "E4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 259 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 525 });

                    break;
                case "E5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 335 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 525 });


                    break;
                case "E6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 409 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 525 });


                    break;
                case "E7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 483 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 525 });

                    break;

                case "E8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 557 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 525 });

                    break;
                case "E9":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 631 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 525 });
                    break;
                case "E10":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 705 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 525 });
                    break;
                case "E11":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 780 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 525 });
                    break;
                case "E12":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 855 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 525 });
                    break;

                case "E13":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 930 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 525 });
                    break;

                case "E14":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1003 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 525 });
                    break;

                case "E15":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1077 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 525 });
                    break;

                case "E16":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1147 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 525 });
                    break;
                #endregion
                #region F
                case "F1":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 42 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 455 });

                    break;

                case "F2":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 112 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 455 });

                    break;

                case "F3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 185 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 455 });
                    break;


                case "F4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 259 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 455 });
                    break;


                case "F5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 335 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 455 });

                    break;

                case "F6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 409 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 455 });
                    break;


                case "F7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 483 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 455 });

                    break;

                case "F8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 557 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 455 });
                    break;
                case "F9":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 631 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 455 });
                    break;
                case "F10":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 705 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 455 });
                    break;
                case "F11":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 780 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 455 });
                    break;
                case "F12":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 855 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 455 });
                    break;

                case "F13":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 930 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 455 });
                    break;

                case "F14":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1003 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 455 });
                    break;

                case "F15":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1077 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 455 });
                    break;

                case "F16":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1147 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 455 });
                    break;
                #endregion
                #region G
                case "G1":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 42 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 385 });
                    break;


                case "G2":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 112 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 385 });
                    break;


                case "G3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 185 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 385 });
                    break;


                case "G4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 259 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 385 });
                    break;


                case "G5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 335 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 385 });
                    break;

                case "G6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 409 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 385 });
                    break;


                case "G7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 483 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 385 });
                    break;


                case "G8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 557 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 385 });
                    break;
                case "G9":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 631 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 385 });
                    break;
                case "G10":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 705 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 385 });
                    break;
                case "G11":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 780 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 385 });
                    break;
                case "G12":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 855 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 385 });
                    break;

                case "G13":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 930 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 385 });
                    break;

                case "G14":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1003 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 385 });
                    break;

                case "G15":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1077 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 385 });
                    break;

                case "G16":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1147 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 385 });
                    break;
                #endregion
                #region H
                case "H1":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 42 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;


                case "H2":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 112 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;


                case "H3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 185 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;


                case "H4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 259 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;


                case "H5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 335 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;

                case "H6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 409 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;


                case "H7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 483 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;


                case "H8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 557 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;
                case "H9":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 631 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;
                case "H10":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 705 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;
                case "H11":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 780 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;
                case "H12":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 855 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;

                case "H13":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 930 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;

                case "H14":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1003 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;

                case "H15":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1077 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;

                case "H16":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1147 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;
                #endregion
                #region J
                case "J1":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 42 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;


                case "J2":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 112 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;


                case "J3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 185 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;


                case "J4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 259 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;


                case "J5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 335 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;

                case "J6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 409 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 737 });
                    break;


                case "J7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 483 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;


                case "J8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 557 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;
                case "J9":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 631 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;
                case "J10":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 705 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;
                case "I11":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 780 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;
                case "J12":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 855 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;

                case "J13":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 930 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;

                case "J14":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1003 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;

                case "J15":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1077 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;

                case "J16":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1147 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;
                #endregion
                #region K
                case "K1":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 42 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 176 });
                    break;


                case "K2":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 112 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 176 });
                    break;


                case "K3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 185 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 176 });
                    break;


                case "K4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 259 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 176 });
                    break;


                case "K5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 335 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 176 });
                    break;

                case "K6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 409 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 176 });
                    break;


                case "K7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 483 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 176 });
                    break;


                case "K8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 557 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 176 });
                    break;
                case "K9":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 631 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 176 });
                    break;
                case "K10":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 705 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 176 });
                    break;
                case "K11":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 780 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 176 });
                    break;
                case "K12":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 855 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 176 });
                    break;

                case "K13":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 930 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 176 });
                    break;

                case "K14":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1003 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 176 });
                    break;

                case "K15":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1077 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 176 });
                    break;

                case "K16":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1147 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 176 });
                    break;
                #endregion
                #region L
                case "L1":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 42 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 104 });
                    break;


                case "L2":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 112 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 104 });
                    break;


                case "L3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 185 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 104 });
                    break;


                case "L4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 259 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 104 });
                    break;


                case "L5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 335 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 104 });
                    break;

                case "L6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 409 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 104 });
                    break;


                case "L7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 483 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 104 });
                    break;


                case "L8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 557 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 104 });
                    break;
                case "L9":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 631 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 104 });
                    break;
                case "L10":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 705 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 104 });
                    break;
                case "L11":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 780 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 104 });
                    break;
                case "L12":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 855 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 104 });
                    break;

                case "L13":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 930 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 104 });
                    break;

                case "L14":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1003 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 104 });
                    break;

                case "L15":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1077 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 104 });
                    break;

                case "L16":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1147 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 104 });
                    break;
                #endregion 
                #region M
                case "M1":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 42 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;


                case "M2":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 112 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;


                case "M3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 185 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;


                case "M4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 259 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;


                case "M5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 335 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;

                case "M6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 409 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;


                case "M7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 483 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;


                case "M8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 557 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;
                case "M9":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 631 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;
                case "M10":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 705 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;
                case "M11":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 780 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;
                case "M12":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 855 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;

                case "M13":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 930 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;

                case "M14":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1003 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;

                case "M15":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1077 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;

                case "M16":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 1147 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;
                    #endregion
            }

            return atParamData;
        }

        private IList<ExecuteSheetOperationRecordData> sheetPlacementDestination3(string positionChosen, IList<ExecuteSheetOperationRecordData> atParamData)
        {
            switch (positionChosen)
            {
                #region A
                case "A1":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 28 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 270 });
                    break;

                case "A2":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 78 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 270 });
                    break;

                case "A3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 131 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 270 });
                    break;


                case "A4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 184 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 270 });
                    break;


                case "A5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 236 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 270 });
                    break;


                case "A6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 289 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 270 });
                    break;


                case "A7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 341 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 270 });
                    break;


                case "A8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 390 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 270 });
                    break;


                #endregion
                #region B
                case "B1":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 28 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 222 });
                    break;

                case "B2":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 78 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 222 });
                    break;

                case "B3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 131 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 222 });
                    break;


                case "B4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 184 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 222 });
                    break;


                case "B5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 236 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 222 });
                    break;


                case "B6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 289 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 222 });
                    break;


                case "B7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 341 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 222 });
                    break;


                case "B8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 390 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 222 });
                    break;


                #endregion
                #region C
                case "C1":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 28 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 173 });
                    break;

                case "C2":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 78 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 173 });
                    break;

                case "C3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 131 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 173 });
                    break;


                case "C4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 184 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 173 });
                    break;


                case "C5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 236 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 173 });
                    break;


                case "C6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 289 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 173 });
                    break;


                case "C7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 341 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 173 });
                    break;


                case "C8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 390 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 173 });
                    break;


                #endregion
                #region D
                case "D1":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 28 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 124 });
                    break;

                case "D2":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 78 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 124 });
                    break;

                case "D3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 131 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 124 });
                    break;


                case "D4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 184 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 124 });
                    break;


                case "D5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 236 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 124 });
                    break;


                case "D6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 289 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 124 });
                    break;


                case "D7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 341 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 124 });
                    break;


                case "D8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 390 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 124 });
                    break;


                #endregion
                #region E
                case "E1":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 25 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 70 });


                    break;
                case "E2":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 75 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 70 });


                    break;
                case "E3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 130 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 70 });


                    break;
                case "E4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 185 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 70 });

                    break;
                case "E5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 240 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 70 });


                    break;
                case "E6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 270 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 70 });


                    break;
                case "E7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 340 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 70 });

                    break;

                case "E8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 390 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 70 });

                    break;

                #endregion
                #region F
                case "F1":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 28 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 27 });
                    break;

                case "F2":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 78 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 27 });
                    break;

                case "F3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 131 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 27 });
                    break;


                case "F4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 184 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 27 });
                    break;


                case "F5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 236 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 27 });
                    break;


                case "F6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 289 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 27 });
                    break;


                case "F7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 341 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 27 });
                    break;


                case "F8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 390 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 27 });
                    break;


                    #endregion
            }
            return atParamData;
        }

        private IList<ExecuteSheetOperationRecordData> sheetPlacementDestination2(string positionChosen, IList<ExecuteSheetOperationRecordData> atParamData)
        {
            switch (positionChosen)
            {
                #region A
                case "A1":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 43 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 380 });
                    break;

                case "A2":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 111 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 380 });
                    break;

                case "A3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 185 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 380 });
                    break;


                case "A4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 260 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 380 });
                    break;


                case "A5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 334 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 380 });
                    break;


                case "A6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 408 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 380 });
                    break;


                case "A7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 482 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 380 });
                    break;


                case "A8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 552 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 380 });
                    break;


                #endregion
                #region B
                case "B1":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 43 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;

                case "B2":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 111 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;

                case "B3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 185 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;


                case "B4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 260 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;


                case "B5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 334 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;


                case "B6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 408 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;


                case "B7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 482 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;


                case "B8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 552 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 315 });
                    break;


                #endregion
                #region C
                case "C1":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 43 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;

                case "C2":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 111 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;

                case "C3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 185 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;


                case "C4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 260 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;


                case "C5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 334 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;


                case "C6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 408 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;


                case "C7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 482 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;


                case "C8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 552 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 245 });
                    break;


                #endregion
                #region D
                case "D1":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 43 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 175 });
                    break;

                case "D2":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 111 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 175 });
                    break;

                case "D3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 185 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 175 });
                    break;


                case "D4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 260 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 175 });
                    break;


                case "D5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 334 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 175 });
                    break;


                case "D6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 408 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 175 });
                    break;


                case "D7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 482 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 175 });
                    break;


                case "D8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 552 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 175 });
                    break;


                #endregion
                #region E
                case "E1":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 43 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 105 });
                    break;

                case "E2":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 111 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 105 });
                    break;

                case "E3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 185 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 105 });
                    break;


                case "E4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 260 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 105 });
                    break;


                case "E5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 334 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 105 });
                    break;


                case "E6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 408 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 105 });
                    break;


                case "E7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 482 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 105 });
                    break;


                case "E8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 552 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 105 });
                    break;


                #endregion
                #region F
                case "F1":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 43 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;

                case "F2":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 111 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;

                case "F3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 185 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;


                case "F4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 260 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;


                case "F5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 334 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;


                case "F6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 408 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;


                case "F7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 482 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;


                case "F8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 552 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 40 });
                    break;


                    #endregion
            }
            return atParamData;
        }

        private IList<ExecuteSheetOperationRecordData> sheetPlacementDestination1(string positionChosen, IList<ExecuteSheetOperationRecordData> atParamData)
        {
            switch (positionChosen)
            {
                #region A
                case "A1":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 38 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 554.6 });
                    break;

                case "A2":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 105 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 554.6 });
                    break;

                case "A3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 175 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 554.6 });
                    break;


                case "A4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 245 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 554.6 });
                    break;


                case "A5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 315 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 554.6 });
                    break;


                case "A6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 385 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 554.6 });
                    break;


                case "A7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 456 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 554.6 });
                    break;


                case "A8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 525 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 554.6 });
                    break;
                case "A9":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 596 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 554.6 });
                    break;

                case "A10":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 666 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 554.6 });
                    break;
                case "A11":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 736 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 554.6 });
                    break;
                case "A12":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 804 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 554.6 });
                    break;



                #endregion
                #region B
                case "B1":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 38 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 483.3 });
                    break;

                case "B2":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 105 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 483.3 });
                    break;

                case "B3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 175 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 483.3 });
                    break;


                case "B4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 245 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 483.3 });
                    break;


                case "B5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 315 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 483.3 });
                    break;


                case "B6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 385 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 483.3 });
                    break;


                case "B7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 456 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 483.3 });
                    break;


                case "B8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 525 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 483.3 });
                    break;
                case "B9":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 596 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 483.3 });
                    break;

                case "B10":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 666 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 483.3 });
                    break;
                case "B11":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 736 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 483.3 });
                    break;
                case "B12":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 804 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 483.3 });
                    break;



                #endregion
                #region C
                case "C1":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 38 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 409 });
                    break;

                case "C2":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 105 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 409 });
                    break;

                case "C3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 175 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 409 });
                    break;


                case "C4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 245 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 409 });
                    break;


                case "C5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 315 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 409 });
                    break;


                case "C6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 385 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 409 });
                    break;


                case "C7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 456 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 409 });
                    break;


                case "C8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 525 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 409 });
                    break;
                case "C9":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 596 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 409 });
                    break;

                case "C10":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 666 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 409 });
                    break;
                case "C11":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 736 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 409 });
                    break;
                case "C12":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 804 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 409 });
                    break;



                #endregion
                #region D
                case "D1":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 38 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 334 });
                    break;

                case "D2":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 105 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 334 });
                    break;

                case "D3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 175 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 334 });
                    break;


                case "D4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 245 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 334 });
                    break;


                case "D5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 315 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 334 });
                    break;


                case "D6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 385 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 334 });
                    break;


                case "D7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 456 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 334 });
                    break;


                case "D8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 525 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 334 });
                    break;
                case "D9":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 596 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 334 });
                    break;

                case "D10":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 666 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 334 });
                    break;
                case "D11":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 736 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 334 });
                    break;
                case "D12":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 804 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 334 });
                    break;

                #endregion
                #region E
                case "E1":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 38 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 259 });
                    break;

                case "E2":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 105 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 259 });
                    break;

                case "E3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 175 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 259 });
                    break;


                case "E4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 245 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 259 });
                    break;


                case "E5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 315 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 259 });
                    break;


                case "E6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 385 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 259 });
                    break;


                case "E7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 456 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 259 });
                    break;


                case "E8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 525 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 259 });
                    break;
                case "E9":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 596 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 259 });
                    break;

                case "E10":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 666 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 259 });
                    break;
                case "E11":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 736 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 259 });
                    break;
                case "E12":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 804 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 259 });
                    break;

                #endregion
                #region F
                case "F1":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 38 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 185 });
                    break;

                case "F2":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 105 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 185 });
                    break;

                case "F3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 175 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 185 });
                    break;


                case "F4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 245 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 185 });
                    break;


                case "F5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 315 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 185 });
                    break;


                case "F6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 385 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 185 });
                    break;


                case "F7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 456 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 185 });
                    break;


                case "F8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 525 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 185 });
                    break;
                case "F9":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 596 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 185 });
                    break;

                case "F10":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 666 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 185 });
                    break;
                case "F11":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 736 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 185 });
                    break;
                case "F12":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 804 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 185 });
                    break;

                #endregion
                #region G
                case "G1":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 38 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 111 });
                    break;

                case "G2":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 105 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 111 });
                    break;

                case "G3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 175 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 111 });
                    break;


                case "G4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 245 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 111 });
                    break;


                case "G5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 315 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 111 });
                    break;


                case "G6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 385 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 111 });
                    break;


                case "G7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 456 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 111 });
                    break;


                case "G8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 525 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 111 });
                    break;
                case "G9":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 596 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 111 });
                    break;

                case "G10":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 666 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 111 });
                    break;
                case "G11":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 736 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 111 });
                    break;
                case "G12":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 804 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 111 });
                    break;

                #endregion
                #region H
                case "H1":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 38 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 39 });
                    break;

                case "H2":
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 105 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 39 });
                    break;

                case "H3":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 175 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 39 });
                    break;


                case "H4":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 245 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 39 });
                    break;


                case "H5":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 315 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 39 });
                    break;


                case "H6":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 385 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 39 });
                    break;


                case "H7":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 456 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 39 });
                    break;


                case "H8":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 525 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 39 });
                    break;
                case "H9":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 596 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 39 });
                    break;

                case "H10":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 666 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 39 });
                    break;
                case "H11":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 736 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 39 });
                    break;
                case "H12":

                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosX, Value = 804 });
                    atParamData.Add(new ExecuteSheetOperationRecordData() { Qualifier = (int)ExecuteSheetOperationModes.ArgExecSheetPosY, Value = 39 });
                    break;

                    #endregion

            }
            return atParamData;
        }

        private ObjectCollection GetAllPlantFromProject(Application myApplication)
        {
            FilterExpression filter = myApplication.CreateFilter();
            filter.TypeId = ObjectType.FctPlant;
            var PlantCollection = myApplication.ActiveProject.FunctionsFolder.FindObjects(filter, SearchBehavior.Deep);

            return PlantCollection;
        }

        private ObjectItem BreakUpComponent(GatherInfo breakupCompObj, ObjectItem StencilList)
        {
            string outpuOfResult;
            if (StencilList == null)
            {

            }
            if (StencilList != null)
            {

            breakupCompObj.FetchedStencil = StencilList;
                //StencilList.RemoveAll(c => c.Id == breakupCompObj.FetchedStencil.Id);
             
                    //  System.Windows.MessageBox.Show("HI");
 
            //    if (Function2 != null)
            //    {
                    Function2 = breakupCompObj.FetchedStencil;   //row.ObjectItem;
                    if (Function2.Parent.IsDeleted == false)
                    {
                        if (Function2.TypeId == ObjectType.FctPlant)
                        {
                            GatherIdObject = new GatherInfo() { Pos = Function2.Id, FetchedStencil = Function2, StencilComp = breakupCompObj.StencilComp, TargetSheet = breakupCompObj.TargetSheet };
                            breakupCompObj.FetchedStencil.MoveTo(breakupCompObj.StencilComp.Parent);
                            var funcNames = breakupCompObj.FetchedStencil.Children.Where(s => s.Attributes.GetAttributeValue(AttributeId.Comment) == breakupCompObj.StencilComp.Attributes.GetAttributeValue((AttributeId)Find3)).Select(s => s).ToList();
                            funcNames.ForEach(item =>
                            {
                                item.Attributes.SetAttributeValue(AttributeId.Designation, breakupCompObj.StencilComp.Attributes.GetAttributeValue(AttributeId.Designation));

                                breakupCompObj.StencilComp.ExecuteFormula("A12199;", out outpuOfResult);
                                int test = 27434;
                                item.Attributes.SetAttributeValue((AttributeId)test, outpuOfResult);

                                item.Store();

                            });

                            if (breakupCompObj.FetchedStencil.Name.Contains(breakupCompObj.StencilComp.Attributes.GetAttributeValue((AttributeId)InformationOnname)))    //row.StencilValue.Attributes.TryFindById((AttributeId)Find, out AttributeItem attr2))  //item3.StencilComp.Attributes.TryFindById((AttributeId)Find, out AttributeItem attr2))
                            {
                            ExApplication exApplication = myApplication as ExApplication;
                            exApplication.ExtendedUtils.ExecuteCommand(OperationCommand.SynchronizeTreeToObject, breakupCompObj.FetchedStencil);

                            //ExApplication exApp = myApplication as ExApplication;
                            //]ObjectItem selection = myApplication.Selection.;
                            //if (selection != null)
                            //{
                            try
                                {
                                    (new System.Threading.Thread(CloseIt)).Start();
                                    ObjectItem startObject = breakupCompObj.FetchedStencil;
                              //  myApplication.Utils.RunMacroOrPlugIn("Test.Module1.Test", startObject);
                                  myApplication.Utils.RunMacroOrPlugIn("Break_up_component.Wizard.Run", startObject);

                                }
                                catch (Exception ex) //add reference to Aucotec.EngineeringBase.Client.Common
                                {
                                     System.Windows.MessageBox.Show("Macro was not found!", "", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                                }
                                //}

                            }

                        }
                    }
              //  }

            }
            return StencilList;

        }


        private void mergeAttribute(StenFunc importantFold)
        {
            var checksameInFolderchild = from folder in importantFold.StenName.Children from StencilCom in GatherIdObject.StencilComp.Children where folder.Name == StencilCom.Name select new { folder, StencilCom };

            checksameInFolderchild.ToList().ForEach(item =>
            {
                item.StencilCom.Attributes.ToList().ForEach(attr =>
                {
                    item.folder.Attributes.SetAttributeValue(attr.Id, attr.Value);
                    item.folder.Store();
                });
            });
        }

        private object SetFoldForDevice(List<GatherInfo> output, List<StenFunc> savedFol)
        {
            output.ForEach(item3 =>
            {
                Folder = item3.StencilComp.Parent.NewChild(ObjectKind.Folder);
                Folder.Attributes.SetAttributeValue(AttributeId.Designation, item3.StencilComp.Name);
                Folder.Store();
                SavedFol.Add(new StenFunc() { FuncName = item3.StencilComp, StenName = Folder });
            });
            return savedFol;
        }

        private void mergeFuntion(GatherInfo objMerge, List<StenFunc> folder)
        {


            //samefold.ToList().ForEach(i => { i.xc.MoveTo(i.StenName); });
            //stencil2.ToList().ForEach(item => { item.find.Children.ToList().ForEach(child => { child.Delete(); }); });
            //var samefoldtomerge = from i in SavedFol from b in i.StenName.Children select new { b };

            var empty = objMerge.StencilComp;
            //samefoldtomerge.ToList().ForEach(x => x.b.MoveTo(empty));

            //var data = objMerge.FetchedStencil.Select(v => v.StencilValue).ToList();
            //var toCompare = stencil2.Select(m => m.find).ToList();
            var existing = objMerge.StencilComp.Children.Where(o => o.Children.Any(i => i.Name != null)).SelectMany(j => j.Children).ToList();

            List<ObjectItem> listtoAdd = new List<ObjectItem>();

            foreach (var item in objMerge.StencilComp.Children)
            {

                var singleobjMerge = objMerge.FetchedStencil.Children.ToList().Select(d => d).Where(l => l.Name == item.Name).FirstOrDefault();
                listtoAdd.Add(singleobjMerge);


            }
            objMerge.StencilComp.Children.ToList().ForEach(child => { { child.Delete(); }; });
            var folderToMove = folder.Select(x => x).Where(c => c.StenName.Name == objMerge.StencilComp.Name).FirstOrDefault();
            foreach (var obj in listtoAdd)
            {
                obj.MoveTo(folderToMove.FuncName);
            }

        }

        private List<GatherInfo> StencilBySelectedItem(List<GatherInfo> gatherInfor2)
        {
            List<GatherInfo> gatherInfor3 = new List<GatherInfo>();
            int Find = 12199;
            FindStencil = myApplication.Folders.Stencils.FindObjects(ObjectKind.StencilCircuitComponent, SearchBehavior.Deep);

            var FindStencilChild = from item1 in FindStencil from items in item1.Children select items;
            var FindStencilChild2 = from items in FindStencilChild from item in gatherInfor2 where item.StencilComp.Attributes.GetAttributeValue((AttributeId)Find) == items.Name select new { items, item };


            gatherInfor3.AddRange((from i in FindStencilChild2 select new GatherInfo { TargetSheet = i.item.TargetSheet, CircuitComp = i.item.CircuitComp, Pos = i.item.Pos, StencilComp = i.item.StencilComp, FetchedStencil = i.items }));


            return gatherInfor3;


        }

        public void CloseIt()
        {
            try
            {
                System.Threading.Thread.Sleep(850);
                Microsoft.VisualBasic.Interaction.AppActivate(
                     System.Diagnostics.Process.GetCurrentProcess().Id);
                System.Windows.Forms.SendKeys.SendWait(" ");
            }
            catch (Exception ex)
            {

            }
        }


    }
}
