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
                       
        }
       
        /// <summary>
        /// Gets called on initialization takes the Contract passed from SYSPRO
        /// This method also calls on Nucleus.dll for the SELECT statement 
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        public string LoadAccount_AccountContacts(string ID)
        {
            Nucleus.Agent _agent = new Nucleus.Agent();
            string _nucleusAgent_SelectRequest = ("SELECT * FROM ZContractContacts WHERE Contract = " + "'" + ID + "' ORDER BY AccountName" );            
            _agent.Select(_nucleusAgent_SelectRequest);            
            ProLogic_zContractContacts.AddRange(_agent.getProLogic_zContractContacts);
            return null;
        }


        /// <summary>
        /// Tree View - loops over the list returned from Nucleus _agent SELECT statement
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void trvAccount_AccountContactsLoaded(object sender, RoutedEventArgs e)
        {            
            TreeViewItem accountItem = new TreeViewItem();
            TreeViewItem empItem = new TreeViewItem();
            var tree = sender as TreeView;
            for (int i = 0; i < ProLogic_zContractContacts.Count; i++)
            {               
                if(ProLogic_zContractContacts[i].Contains("{ Header = Item Level 0 }"))
                {
                    accountItem = new TreeViewItem();
                    accountItem.Header = ProLogic_zContractContacts[i].Replace("{ Header = Item Level 0 }", "");
                    tree.Items.Add(accountItem);
                }
                if(ProLogic_zContractContacts[i].Contains("{ Header = Item Level 1 }"))
                {
                    empItem = new TreeViewItem();
                    empItem.Header = ProLogic_zContractContacts[i].Replace("{ Header = Item Level 1 }", "");
                    accountItem.Items.Add(empItem);
                }
            }           
        }
    }
}
