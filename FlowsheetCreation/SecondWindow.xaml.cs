using Aucotec.EngineeringBase.Client.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
namespace FlowsheetCreation
{
    /// <summary>
    /// Interaction logic for SecondWindow.xaml
    /// </summary>
    public partial class SecondWindow : Window
    {
      

        public SecondWindow()
        {
            InitializeComponent();
            
          
        }

        private void ComboBox_DropDownClosed(object sender, EventArgs e)
        {
            string circuitcomp;
            var vm = DataContext as VmSecondWindow;
            vm.upPos();
            var combo = sender as ComboBox;
            if (combo.Text == "Search")
            {
                vm.MySelectedItem.Source.ExecuteFormula("A12199;", out circuitcomp);
                combo.Text = circuitcomp;
            }
        }

        private void ComboBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            
            var vm = DataContext as VmSecondWindow;
            var combo = sender as ComboBox;
            combo.ItemsSource = vm.Circomp;
        

               vm.Circomp.Clear();
               vm.FindStencil = vm.myApplication.Folders.Stencils.FindObjects(ObjectKind.StencilCircuitComponent, SearchBehavior.Deep);
               var findstencil = from x in vm.FindStencil from y in x.Children select y;
            //   char[] charArray = combo.Text.ToCharArray();

               findstencil.ToList().ForEach( item =>
               {
                  // foreach (var c in combo.Text.ToCharArray())
                 //  {
                      // string test = c.ToString();
                       if (item.Name.Contains(combo.Text))
                       //  if (item.Name.Contains(combo.Text))
                       {
                           vm.Circomp.Add(new ModelView() { TypeIcon = item.Image, Type1Name = item.Name, Source = item });

                       }

                  // }

               } );



        }

        private void combo_DropDownOpened(object sender, EventArgs e)
        {
            var combo = sender as ComboBox;
            combo.Text = "Please Search";
 
           
        }

    }
}
