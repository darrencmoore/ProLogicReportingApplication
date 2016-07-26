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
using System.Timers;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using System.Net;
using System.Net.Mail;


/// <summary>
/// Created By Darren Moore
/// </summary>


namespace ProLogicReportingApplication
{    
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window 
    {        
        private List<String> proLogic_ContractContacts = new List<String>();        
        private ObservableCollection<String> proLogic_ContractContactsObservable = new ObservableCollection<String>();
        private Timer _mouseClickTimer = null;
        private bool clickHandled = false;       
        private int userClicks;

        public MainWindow()
        {
            InitializeComponent();
            //string[] args = Environment.GetCommandLineArgs();
            //MessageBox.Show(args[1]); pass args[1] to LoadContacts when live
            // Call to Nucleus to get the data to populate the tree view
            LoadContacts("00002");
            //this.ContactsGrid.ItemsSource = ProLogic_zContractContacts;
            //ContactsGrid.ItemsSource = ProLogic_zContractContacts.ToList();
        }

        /// <summary>
        /// Gets called on initialization takes the Contract passed from SYSPRO
        /// This method also calls on Nucleus.dll for the GetContacts() 
        /// </summary>
        /// <param name="contractId"></param>
        /// <returns></returns>
        public string LoadContacts(string contractId)
        {
            Nucleus.Agent _agent = new Nucleus.Agent();
            _agent.GetContacts(contractId);
            // Adds the list from Nucleus.Agent                      
            proLogic_ContractContacts.AddRange(_agent.Agent_ContractContactsResponse);
            // Copies ProLogic_zContractContacts to an observable collection
            // This is to be used for the treeview           
            proLogic_ContractContactsObservable = new ObservableCollection<String>(proLogic_ContractContacts);

            return null;
        }

        /// <summary>
        /// This will handle Selected Item Changes
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void trvTree_Collapsed(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            Console.WriteLine("TreeView Collapsed");
        }

        /// <summary>
        /// Bid Email Send
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_SendBid_Click(object sender, RoutedEventArgs e)
        {
            // Add smtp email stuff here
        }

        #region TreeView Load 
        /// <summary>
        /// Tree View - loops over the list returned from Nucleus _agent SELECT statement
        /// Puts each item in ProLogic_zContractContacts into a TreeViewItem
        /// Heirarchy goes Account => ContactFullName -> Contact Email Address. 
        /// I'm also using the account and for the report generation parameters needed are Account and Contract
        /// This gets called from MainWindow.xaml when the TreeView is intially loaded
        /// As there is not a OnClick event for TreeView nor the TreeViewItems I had to use PreviewMouseLeftButtonDown
        /// changed from using ProLogiz_zContractContacts list to zContractsContactsObservable
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void trvAccount_AccountContactsLoaded(object sender, RoutedEventArgs e)
        {
            TreeViewItem accountItem = new TreeViewItem();
            TreeViewItem empItem = new TreeViewItem();           
            TreeViewItem empEmailAddrItem = new TreeViewItem();
            var tree = sender as TreeView;

            for (int i = 0; i < proLogic_ContractContactsObservable.Count; i++)
            {
                //AccountName = Level 0                       
                if (proLogic_ContractContactsObservable[i].Contains("{ Header = Item Level 0 }"))
                {                 
                    accountItem = new TreeViewItem();
                    // Removing everything after the account ID for the Tag
                    accountItem.Tag = "{Parent} " + proLogic_ContractContactsObservable[i].Remove(5);
                    accountItem.IsExpanded = false;
                    accountItem.FontWeight = FontWeights.Black;                    
                    accountItem.Header = proLogic_ContractContactsObservable[i].Replace("{ Header = Item Level 0 }", "");
                    accountItem.Header = new CheckBox()
                    {
                        IsChecked = true,
                        IsEnabled = true,
                        Focusable = true,                      
                        Name = "ParentChkBox",
                        // Removing the account ID and just keeping the Account Name                                                
                        Content = proLogic_ContractContactsObservable[i].Remove(0,4).Replace("{ Header = Item Level 0 }", "")
                    };
                    tree.Items.Add(accountItem);
                }
                //ContactFullName = Level 1
                if (proLogic_ContractContactsObservable[i].Contains("{ Header = Item Level 1 }"))
                {                    
                    empItem = new TreeViewItem();
                    // Removing everything after the account ID for the Tag
                    empItem.Tag = "{Child} " + proLogic_ContractContactsObservable[i].Remove(5);
                    empItem.FontWeight = FontWeights.Black;
                    empItem.Header = proLogic_ContractContactsObservable[i].Replace("{ Header = Item Level 1 }", "");
                    empItem.Header = new CheckBox()
                    {
                        IsChecked = true,
                        IsEnabled = true,
                        Focusable = true,
                        Name = "ChildChkBox",
                        // Removing the account ID and just keeping the Contact Full Name
                        Content = proLogic_ContractContactsObservable[i].Remove(0,4).Replace("{ Header = Item Level 1 }", "")
                    };
                    accountItem.Items.Add(empItem);
                }
                //Contact Email Address = Level 2
                if (proLogic_ContractContactsObservable[i].Contains("{ Header = Item Level 2 }"))
                {                    
                    empEmailAddrItem = new TreeViewItem();
                    empEmailAddrItem.Tag = "{Email} " + proLogic_ContractContactsObservable[i].Replace("{ Header = Item Level 2 }", "");
                    empEmailAddrItem.FontWeight = FontWeights.Black;
                    empEmailAddrItem.Header = proLogic_ContractContactsObservable[i].Replace("{ Header = Item Level 2 }", "");
                    empEmailAddrItem.Header = new CheckBox()
                    {
                        IsChecked = true,
                        IsEnabled = false,
                        Content = proLogic_ContractContactsObservable[i].Replace("{ Header = Item Level 2 }", "")
                    };
                    empItem.Items.Add(empEmailAddrItem);
                }
            }
        }
        #endregion        

        #region Preview Mouse Click Events
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        [DllImport("user32.dll")]
        static extern uint GetDoubleClickTime();
        private void previewMouse_SingleClick(object sender, MouseButtonEventArgs e)
        {                     
            if (e.ChangedButton == MouseButton.Left)
            {
                userClicks++;
                _mouseClickTimer = new Timer(GetDoubleClickTime());
                _mouseClickTimer.AutoReset = false;
                _mouseClickTimer.Elapsed += new ElapsedEventHandler(MouseClickTimer);                    

                if (!_mouseClickTimer.Enabled)
                {                    
                    _mouseClickTimer.Start();
                    if(userClicks == 1)
                    {
                        NodeCheck(e.OriginalSource as DependencyObject, userClicks);
                        return;
                    }
                    if(userClicks >= 2)
                    {
                        Console.WriteLine("2 clicks just once");
                        e.Handled = true;
                        return;
                    }   
                }                            
            }           
        }
        #endregion

        #region Mouse Click Timer
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MouseClickTimer(object sender, ElapsedEventArgs e)
        {  
            _mouseClickTimer.Stop();
            Console.WriteLine("Mouse Timer Clicks -> " + userClicks);
            userClicks = 0;
            Console.WriteLine("Timer Stop -> " + userClicks);        
            //Console.WriteLine("timer stopped - Call Method");
        }
        #endregion

        #region Parent/Child Node Check
        /// <summary>
        ///  Handling parent and child clicks here
        ///  this is the primary UI function for checkbox manipulation
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        private static TreeViewItem NodeCheck(DependencyObject source, int numClicks)
        {
            while (source != null && !(source is TreeViewItem))
                source = System.Windows.Media.VisualTreeHelper.GetParent(source);
            TreeViewItem item = source as TreeViewItem;

            if (numClicks == 1)
            {
                try
                {
                    if (item != null)
                    {
                        for (int i = 0; i < item.ItemContainerGenerator.Items.Count; i++)
                        {
                            Console.WriteLine("Item Generator -> " + item.ItemContainerGenerator.Items[i].ToString());
                        }
                        Console.WriteLine("Item Name -> " + item);
                        Console.WriteLine("Item Parent -> " + item.Parent);
                        Console.WriteLine("Item Header -> " + item.Header);
                        Console.WriteLine("Item Tag -> " + item.Tag);

                        if (item.Tag.ToString().Contains("{Parent}"))
                        {
                            Console.WriteLine("Parent");
                            item.Focusable = true;
                            item.IsSelected = true;
                            ContentPresenter parentTreeItemContentPresenter = item.Template.FindName("PART_Header", item) as ContentPresenter;
                            CheckBox parentTreeItemChkBox = item.Header as CheckBox;
                            if (parentTreeItemContentPresenter != null && parentTreeItemChkBox.Name.ToString() == "ParentChkBox")
                            {
                                if (parentTreeItemChkBox.IsChecked == true)
                                {
                                    Console.WriteLine("Parent Check Box Unchecked  -> " + parentTreeItemChkBox);

                                }
                                else if (parentTreeItemChkBox.IsChecked == false)
                                {
                                    Console.WriteLine("Parent Check Box is Checked -> " + parentTreeItemChkBox);                                    
                                }
                            }
                        }
                        if (item.Tag.ToString().Contains("{Child}"))
                        {
                            Console.WriteLine("Child");
                            item.Focusable = true;
                            item.IsSelected = true;
                            ContentPresenter childTreeItemContentPresenter = item.Template.FindName("PART_Header", item) as ContentPresenter;
                            CheckBox childTreeItemChkBox = item.Header as CheckBox;
                            if (childTreeItemContentPresenter != null && childTreeItemChkBox.Name.ToString() == "ChildChkBox")
                            {

                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.ToString());
                }
            }            
            return item as TreeViewItem;            
        }
        #endregion
        
    }
}
