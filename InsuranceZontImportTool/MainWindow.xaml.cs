using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace InsuranceZontImportTool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private List<CheckBox> GetSelectedCheckBoxes(ItemCollection items)
        {
            List<CheckBox> list = new List<CheckBox>();
            foreach (TreeViewItem item in items)
            {
                UIElement elemnt = GetChildControl(item, "chkBox");
                if (elemnt != null)
                {
                    CheckBox chk = (CheckBox)elemnt;
                    if (chk.IsChecked.HasValue && chk.IsChecked.Value)
                    {
                        list.Add(chk);
                    }
                }

                List<CheckBox> l = GetSelectedCheckBoxes(item.Items);
                list = list.Concat(l).ToList();
            }

            return list;
        }

        private UIElement GetChildControl(DependencyObject parentObject, string childName)
        {

            UIElement element = null;

            if (parentObject != null)
            {
                int totalChild = VisualTreeHelper.GetChildrenCount(parentObject);
                for (int i = 0; i < totalChild; i++)
                {
                    DependencyObject childObject = VisualTreeHelper.GetChild(parentObject, i);

                    if (childObject is FrameworkElement && ((FrameworkElement)childObject).Name == childName)
                    {
                        element = childObject as UIElement;
                        break;
                    }

                    // get its child
                    element = GetChildControl(childObject, childName);
                    if (element != null) break;
                }
            }

            return element;
        }

        private void Window_Activated(object sender, EventArgs e)
        {

        }

        private void btnGetAdvisorFunds_Click(object sender, RoutedEventArgs e)
        {
            //Get the BPC PDF Folder
            System.Windows.Forms.FolderBrowserDialog folderDlg = new System.Windows.Forms.FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = false;

            if (folderDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // folderDlg.SelectedPath
            }
        }

        private void BindDirectoryTree(string path,TreeViewItem parentNode)
        {
            string[] files = Directory.GetFiles(path, "*.csv", SearchOption.TopDirectoryOnly);
            string[] folders = Directory.GetDirectories(path);

            if(folders != null && folders.Length > 0)
            {

                foreach (var item in folders)
                {
                    BindDirectoryTree(item,parentNode);
                }
                
            }


        }
    }

}
