using Aucotec.EngineeringBase.Client.Runtime;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows.Media;
using static FlowsheetCreation.VmSecondWindow;

namespace FlowsheetCreation
{
    public class ModelView : Helpers.VmBase
    {
        private string _posText;
         public string PosText
        {
            get
            {
                    return _posText;
            }

            set
            {
                _posText = value;
                OnPropertyChanged("PosText");
            }
        }
        private string _sheetText;
        public string SheetText { 
            get {
               
                    return _sheetText;
              
              }
            set
            {
                _sheetText = value;
                OnPropertyChanged("SheetText");
            }
        }
        private string _cirText;
        public string CirText
        {
            get {  
                    return _cirText;
            }
            set
            {
                _cirText = value;
                OnPropertyChanged("CirText");
            }
        }
        private ModelView _SelCir;
        public ModelView SelCir
        {
            get { return _SelCir; }
            set {
                if (value != null)
                    _SelCir = value; OnPropertyChanged("SelCir"); }
        }
        private ModelView _SelTargetSheet;
        public ModelView SelTargetSheet
        {
            get { return _SelTargetSheet; }
            set
            {
                if (value != null)
                    _SelTargetSheet = value; OnPropertyChanged("SelTargetSheet");
            }
        }
        private ModelView _SelPosition;
        public ModelView SelPosition
        {
            get { return _SelPosition; }
            set {
                if(value != null)
                _SelPosition = value; OnPropertyChanged("SelPosition");
            }
        }     
        public ImageSource TypeIcon { get; set; }
        public string Type2Name { get; set; }
        public string Type3Name { get; set; }
        public string Type1Name { get; set; }
        public ObjectItem Source { get; set; }
        public Sheet SourceSheet { get; set; }
        public bool Funcbox { get; set; }
        public ObservableCollection<ModelView> Circomp { get; set; }
        public ObservableCollection<ModelView> Position { get; set; }
        public ObservableCollection<ModelView> TargetSheet { get; set; }
        public ModelView(ObjectItem source, Sheet sourceSheet)
        {
            Source = source;
            SourceSheet = sourceSheet;
            PartOf = source.Name;
            Designation = source.Name;
            TypeIcon = source.Image;
            Type1Name = source.Name;
            Type2Name = source.Name;
            Type3Name = source.Name;
        }

        public ModelView()
        {
        }

        private bool _Checkbox;
        public bool Checkbox
        {
            get
            {
                return _Checkbox;
            }
            set
            {
                _Checkbox = value;
                OnPropertyChanged("Checkbox");

            }
        }
        private string _cirName;
        public string CirName
        {

            get { return _cirName; }
            set
            {
                  _cirName = value;  OnPropertyChanged("CirName");
            }
        }
        private string _partOf;
        public string PartOf
        {

            get { return _partOf; }
            set
            {
                _partOf = value;
                OnPropertyChanged("PartOf");
            }
        }


        private string _designation;
        public string Designation
        {

            get { return _designation; }
            set
            {
                _designation = value;
                OnPropertyChanged("Designation");
            }
        }
        private string _drawing;
        public string Drawing
        {

            get { return _drawing; }
            set
            {
                _drawing = value;
                OnPropertyChanged("Drawing");
            }
        }



    }
 
}
