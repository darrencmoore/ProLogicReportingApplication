using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Drawing;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.SqlClient;
using System.Web;
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
            //string _nucleusAgent_SelectRequest = ("SELECT * FROM ZContractContacts WHERE Contract = " + "'" + ID + "' ORDER BY AccountName" );            
            _agent.Select(ID);            
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
            TreeViewItem empEmailAddrItem = new TreeViewItem();

            var tree = sender as TreeView;
           

            for (int i = 0; i < ProLogic_zContractContacts.Count; i++)
            {
                //AccountName = Level 0               
                if (ProLogic_zContractContacts[i].Contains("{ Header = Item Level 0 }"))
                {
                    accountItem = new TreeViewItem();
                    accountItem.Tag = "{ Header = Item Level 0 }";
                    accountItem.Header = new CheckBox()
                    {
                        IsChecked = true,                        
                        Content = ProLogic_zContractContacts[i].Replace("{ Header = Item Level 0 }", "")
                    };
                    accountItem.PreviewMouseLeftButtonDown += test_click;
                    tree.Items.Add(accountItem);
                }
                //ContactFullName = Level 1
                if (ProLogic_zContractContacts[i].Contains("{ Header = Item Level 1 }"))
                {
                    empItem = new TreeViewItem();
                    empItem.Tag = "{ Header = Item Level 1 }";
                    empItem.Header = new CheckBox()
                    {
                        IsChecked = true,
                        Content = ProLogic_zContractContacts[i].Replace("{ Header = Item Level 1 }", "")
                    };
                    empItem.PreviewMouseLeftButtonDown += test_click;
                    accountItem.Items.Add(empItem);
                }
                //Contact Email Address = Level 2
                if (ProLogic_zContractContacts[i].Contains("{ Header = Item Level 2 }"))
                {
                    empEmailAddrItem = new TreeViewItem();
                    empEmailAddrItem.Tag = "{ Header = Item Level 2 }";
                    empEmailAddrItem.Header = new CheckBox()
                    {
                        IsEnabled = false,
                        IsChecked = true,
                        Content = ProLogic_zContractContacts[i].Replace("{ Header = Item Level 2 }", "")
                    };
                    empItem.Items.Add(empEmailAddrItem);
                }
            }
        }       

        private void test_click(object sender, RoutedEventArgs e)
        {
            NodeCheck(e.OriginalSource as DependencyObject);
            //TreeViewItem item = e.OriginalSource as TreeViewItem;
            //if (item != null)
            //{

            //}
        }

        private static TreeViewItem NodeCheck(DependencyObject source)
        {
            while(source != null && !(source is TreeViewItem)) source = System.Windows.Media.VisualTreeHelper.GetParent(source);
            Console.WriteLine("Source -> " + source.GetType());
            //Console.WriteLine("Source -> " + source.GetValue(source));
            return source as TreeViewItem;
        }

        /// <summary>
        /// This will handle Selected Item Changes
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void trvTree_Collapsed(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            string dadsd;
        }     
    }
}
