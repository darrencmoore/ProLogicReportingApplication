using System;
using System.Collections.Generic;
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
using System.Data.SqlClient;
using Nucleus;



namespace ProLogicReportingApplication
{    
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public List<string> ProLogic_zContractContacts = new List<string>();

        public MainWindow()
        {
            InitializeComponent();            
            //string[] args = Environment.GetCommandLineArgs();
            //MessageBox.Show(args[1]);
            // Call to Nucleus to get the data to populate the tree view
            LoadAccount_AccountContacts("00001");
            //TreeViewItem item = new TreeViewItem();
            //item.ItemsSource = ProLogic_zContractContacts; 
            //this.Loaded += new RoutedEventHandler(trvAccount_AccountContactsLoaded);            
        }
       

        public string LoadAccount_AccountContacts(string ID)
        {
            Nucleus.Agent _agent = new Nucleus.Agent();
            string _nucleusAgent_SelectRequest = ("SELECT * FROM ZContractContacts WHERE Contract = " + "'" + ID + "' ORDER BY AccountName" );            
            _agent.Select(_nucleusAgent_SelectRequest);            
            ProLogic_zContractContacts.AddRange(_agent.getProLogic_zContractContacts);
            return null;
        }

        private void trvAccount_AccountContactsLoaded(object sender, RoutedEventArgs e)
        {
            foreach (string itemNode in ProLogic_zContractContacts)
            {
                TreeViewItem ContactTreeItem = new TreeViewItem();
                var tree = sender as TreeView;
                if (itemNode.Contains("{ Header = Item Level 0 }"))
                {
                    ContactTreeItem.Header = itemNode.Replace("{ Header = Item Level 0 }", "");
                    tree.Items.Add(ContactTreeItem);
                }
                if (itemNode.Contains("{ Header = Item Level 1 }"))
                {
                    ContactTreeItem.ItemsSource = new string[] { itemNode.Replace("{ Header = Item Level 1 }", "") };
                    tree.Items.Add(ContactTreeItem);
                }
            }
        }
    }
}
