using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
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
        
        private List<string> GetSelectedCheckBoxes(TreeViewItem node)
        {
            List<string> list = new List<string>();

            //BPC_PDF Node
            if (node.HasItems)
            {
                //Advisor Nodes
                foreach (TreeViewItem childNode in node.Items)
                {
                    if(childNode.Header is CheckBox)
                    {
                        CheckBox nodeHeader = (CheckBox)childNode.Header;
                        if (nodeHeader.IsChecked.Value)
                        {
                            string advisoName = nodeHeader.Content.ToString();
                            list.Add(advisoName);
                        }
                    }
                }
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

        List<CorpInfo> _corpInfos = new List<CorpInfo>();
        TreeViewItem _tvBPCFolderNode;
        private void btnGetAdvisorFunds_Click(object sender, RoutedEventArgs e)
        {
            //Get the BPC PDF Folder
            System.Windows.Forms.FolderBrowserDialog folderDlg = new System.Windows.Forms.FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = false;
            folderDlg.SelectedPath = @"C:\Temp\INSURANCEZONE\BPC_PDF";

            if (folderDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (tvAdvisorFundsData.HasItems)
                {
                    tvAdvisorFundsData.Items.Clear();
                    _corpInfos = new List<CorpInfo>();
                }

                txtAdvisorFundFolder.Text = folderDlg.SelectedPath;
                _tvBPCFolderNode = new TreeViewItem() {Header= System.IO.Path.GetFileName(folderDlg.SelectedPath),IsExpanded = true }; //CreateTreeViewItem(System.IO.Path.GetFileName(folderDlg.SelectedPath), true, true);
                
                //Advisor Folder level expected inside the BPC_PDF
                string[] folders = Directory.GetDirectories(folderDlg.SelectedPath);

                int fileIndx = 1;
                foreach (var folder in folders)
                {
                    string advisorFolderName = System.IO.Path.GetFileName(folder);
                    TreeViewItem tvAdvisorNode = CreateTreeViewItem(advisorFolderName, true, true);
                    tvAdvisorNode.Tag = folder;

                    _tvBPCFolderNode.Items.Add(tvAdvisorNode);

                    string[] files = Directory.GetFiles(folder, "*.csv", SearchOption.AllDirectories);
                    foreach (var item in files)
                    {
                        DataTable fileData = GetFileData(item);
                        int recordCount = fileData.Rows.Count;
                        _corpInfos.Add(new CorpInfo() {ID = fileIndx, Advisor = advisorFolderName, FileName = item,Data = fileData,RecordCount = recordCount });
                        fileIndx++;
                    }
                }

                tvAdvisorFundsData.Items.Add(_tvBPCFolderNode);

                BindDataToGrid();

            }
        }

        private DataTable GetFileData(string filePath)
        {
            DataTable table = new DataTable();
            
            if(File.Exists(filePath))
            {
                using (StreamReader reader = new StreamReader(filePath))
                {
                    bool isHeader = true;
                    while(!reader.EndOfStream)
                    {
                        if (isHeader)
                        {
                            string header = reader.ReadLine();
                            string[] cols = header.Split(',');
                            if(cols.Length > 0)
                            {
                                List<string> duplCol = new List<string>();

                                foreach (var col in cols)
                                {
                                    if (table.Columns.Contains(col))
                                    {
                                        duplCol.Add(col);

                                        table.Columns.Add(col + duplCol.Where(x=>x== col).Count());
                                    }
                                    else
                                    {
                                        table.Columns.Add(col);
                                    }
                                }
                            }

                            isHeader = false;
                        }
                        else
                        {
                            string line = reader.ReadLine();
                            string[] lineData = line.Split(',');

                            DataRow newRow = table.NewRow();
                            int colIndx = 0;
                            foreach (var data in lineData)
                            {
                                if (table.Columns.Count > colIndx)
                                {
                                    newRow[colIndx] = data;
                                    colIndx++;
                                }
                            }

                            table.Rows.Add(newRow);
                        }
                    }
                }
            }

            return table;
        }

        private void BindDataToGrid()
        {
            
            if (_corpInfos != null && _corpInfos.Count > 0)
            {
                DataTable table = new DataTable();
                table.Columns.Add("ID");
                table.Columns.Add("Advisor");
                table.Columns.Add("Fund");
                table.Columns.Add("Query Type");
                
                List<string> selectedAdvisors = GetSelectedCheckBoxes(_tvBPCFolderNode);
                
                int maxCols = 0;
                foreach (var corp in _corpInfos)
                {

                    string advisor = corp.Advisor;
                    string file = corp.FileName;
                    
                    string fileName = CleanupFileName(file);

                    int cols = fileName.Split('_').Length;
                    if (cols > maxCols)
                    {
                        maxCols = cols;
                    }
                }

                if (maxCols > 0)
                {
                    for (int i = 2; i < maxCols; i++)
                    {
                        table.Columns.Add("Adhoc" + i);
                    }
                }

                table.Columns.Add("Total#");
                table.Columns.Add("Notes");

                foreach (var corp in _corpInfos)
                {

                    string advisor = corp.Advisor;
                    string file = corp.FileName;
                    if (!selectedAdvisors.Any(x => x == advisor))
                    {
                        continue;
                    }
                    string fileName = CleanupFileName(file);

                    List<string> cols = fileName.Split('_').ToList();

                    DataRow newRow = table.NewRow();
                    newRow["ID"] = corp.ID;
                    newRow["Advisor"] = advisor;
                    newRow["Fund"] = cols.ElementAtOrDefault(0);
                    newRow["Query Type"] = cols.ElementAtOrDefault(1);

                    for (int i = 2; i < maxCols; i++)
                    {
                        newRow["Adhoc" + i] = cols.ElementAtOrDefault(i);
                    }

                    newRow["Total#"] = corp.RecordCount;
                    newRow["Notes"] = corp.DataNotes;

                    table.Rows.Add(newRow);
                }

                dgData.ItemsSource = table.AsDataView();
                
            }
            
        }

        private static string CleanupFileName(string file)
        {
            string fileName = System.IO.Path.GetFileNameWithoutExtension(file);
            fileName = fileName.Replace("__", "_");
            fileName = fileName.Replace("E_", "");
            return fileName;
        }

        public TreeViewItem CreateTreeViewItem(string header, bool isExpanded, bool isSelected)
        {
            TreeViewItem tvItem = new TreeViewItem() {IsExpanded = isExpanded };
            
            CheckBox chkBox = new CheckBox();
            chkBox.Name = "chk";
            chkBox.Content = header;
            chkBox.IsChecked = isSelected;

            tvItem.Header = chkBox;

            return tvItem;
        }

        private void btnRefreshGrid_Click(object sender, RoutedEventArgs e)
        {
            BindDataToGrid();
        }

        List<string> selectedRowIds = new List<string>();
        private void btnImportSelected_Click(object sender, RoutedEventArgs e)
        {
            if(dgData.SelectedItems != null)
            {
                foreach (var item in dgData.SelectedItems)
                {
                    if(item is DataRowView)
                    {
                        DataRow row = ((DataRowView)item).Row;
                        selectedRowIds.Add(row["ID"].ToString());
                    }
                }
            }
        }

    }

    public class CorpInfo
    {
        public int ID { get; set; }

        public string Advisor { get; set; }

        public string FileName { get; set; }

        public int RecordCount { get; set; }

        public DataTable Data { get; set; }

        public string DataNotes { get; set; }
    }

}
