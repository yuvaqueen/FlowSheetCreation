
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Media;
using Aucotec.EngineeringBase.Client.Runtime;

namespace FlowsheetCreation
{
    public class ObjectItemLocation : Helpers.VmBase
    {
        private string _locdesignation;
        public string LocDesignation
        {

            get { return _locdesignation; }
            set
            {
                _locdesignation = value;
                OnPropertyChanged("LocDesignation");
            }
        }
        public Sheet Source { get; set; }
        private string _locPartOf;
        public string LocPartOf
        {

            get { return _locPartOf; }
            set
            {
                _locPartOf = value;
                OnPropertyChanged("LocPartOf");
            }
        }
        private bool _locCheckbox;
        public bool LocCheckbox
        {
            get
            {
                return _locCheckbox;
            }
            set
            {
                _locCheckbox = value;
                OnPropertyChanged("LocCheckbox");

            }
        }
        public ObjectItemLocation(Sheet source)
        {
            Source = source;
            LocPartOf = source.Name;
            LocDesignation = source.Name;
          
           
        }

        public ObjectItemLocation()
        {
        }
    }


}
