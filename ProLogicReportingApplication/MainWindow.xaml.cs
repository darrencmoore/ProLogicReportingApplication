using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
//using System.Drawing;
using System.Data;
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
using System.Runtime.Caching;
using System.Net;
using System.Net.Mail;
using CrystalDecisions.Shared;
using CrystalDecisions.CrystalReports.Engine;
using System.Reflection;
using SAPBusinessObjects.WPF.Viewer;

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
        private List<string> proLogic_ContractContacts = new List<string>();        
        private ObservableCollection<string> proLogic_ContractContactsObservable = new ObservableCollection<string>();
        private Timer _mouseClickTimer = null;
        private int userClicks;
        private static string contractId;        

        public MainWindow()
        {
            InitializeComponent();
            // Call to Nucleus to get the data to populate the tree view
            contractId = "00002";
            LoadContacts(contractId);            
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
            proLogic_ContractContactsObservable = new ObservableCollection<string>(proLogic_ContractContacts);

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

        #region Preview Mouse Click Event
        /// <summary>
        /// Get the PreviewLeftMouseButtonDown Event
        /// compares clicks against the users PC double click time
        /// Then calls the NodeCheck method 
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
                        NodeCheck(e.OriginalSource as DependencyObject);
                        return;
                    }
                    else
                    {
                        e.Handled = true;
                        return;
                    }   
                }                            
            }           
        }
        #endregion

        #region Mouse Click Timer
        /// <summary>
        /// Gets called when click timer expires
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

        private void CoalesceTreeView(TreeViewItem node, Boolean isChecked)
        {
            //Console.WriteLine("CoalesceTreeView " + node.);
            //Console.WriteLine("CoalesceTreeView " + node.ItemContainerGenerator.Items.ToString());
        }

        #region Parent/Child Node Check
        /// <summary>
        ///  Handling parent and child clicks here
        ///  this is the primary UI function for checkbox manipulation
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        private TreeViewItem NodeCheck(DependencyObject source)
        {
            while (source != null && !(source is TreeViewItem))
                source = System.Windows.Media.VisualTreeHelper.GetParent(source);
            TreeViewItem item = source as TreeViewItem;

            try
            {
                if (item != null)
                {
                    if (item.Tag.ToString().Contains("{Parent}"))
                    {
                        item.Focusable = true;
                        item.IsSelected = true;
                        ContentPresenter parentTreeItemContentPresenter = item.Template.FindName("PART_Header", item) as ContentPresenter;
                        CheckBox parentTreeItemChkBox = item.Header as CheckBox;
                        if (parentTreeItemContentPresenter != null && parentTreeItemChkBox.Name.ToString() == "ParentChkBox")
                        {
                            if (parentTreeItemChkBox.IsChecked == true)
                            {
                                Console.WriteLine("Parent Check Box Unchecked  -> " + parentTreeItemChkBox);
                                SetChildrenChecks(item, false);                                    
                                ContractBidReportPreview(contractId, item.Tag.ToString().Replace("{Parent}", "").Trim());                                
                            }
                            else if (parentTreeItemChkBox.IsChecked == false)
                            {
                                Console.WriteLine("Parent Check Box is Checked -> " + parentTreeItemChkBox);
                                SetChildrenChecks(item, true);
                                ContractBidReportPreview(contractId, item.Tag.ToString().Replace("{Parent}", "").Trim());
                            }
                        }
                    }
                    if (item.Tag.ToString().Contains("{Child}"))
                    {
                        item.Focusable = true;
                        item.IsSelected = true;
                        ContentPresenter childTreeItemContentPresenter = item.Template.FindName("PART_Header", item) as ContentPresenter;
                        CheckBox childTreeItemChkBox = item.Header as CheckBox;
                        if (childTreeItemContentPresenter != null && childTreeItemChkBox.Name.ToString() == "ChildChkBox")
                        {
                            if (childTreeItemChkBox.IsChecked == true)
                            {
                                Console.WriteLine("Child Check Box Unchecked  -> " + childTreeItemChkBox);
                                Console.WriteLine("Parent -> " + item.Parent.ToString());
                                SetParentChecks((TreeViewItem)item.Parent, true);
                                ContractBidReportPreview(contractId, item.Tag.ToString().Replace("{Child}", "").Trim());
                            }
                            else if (childTreeItemChkBox.IsChecked == false)
                            {
                                Console.WriteLine("Child Check Box is Checked -> " + childTreeItemChkBox);
                                Console.WriteLine("Parent -> " + item.Parent.ToString());
                                SetParentChecks((TreeViewItem)item.Parent, false);
                                ContractBidReportPreview(contractId, item.Tag.ToString().Replace("{Child}", "").Trim());
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            return item;            
        }
        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="item"></param>
        /// <param name="checkedState"></param>
        private void SetChildrenChecks(TreeViewItem item, bool checkedState)
        {
            foreach(TreeViewItem tv in item.Items)
            {
                if(checkedState == true)
                {
                    CheckBox parentHasChildItem = tv.Header as CheckBox;
                    parentHasChildItem.IsChecked = true;
                }
                else
                {
                    CheckBox parentHasChildItem = tv.Header as CheckBox;
                    parentHasChildItem.IsChecked = false;
                }               
            }
        }

        private void SetParentChecks(TreeViewItem item, bool checkedState)
        {

            if (checkedState == true)
            {
                CheckBox childsParentItem = item.Header as CheckBox;
                childsParentItem.IsChecked = true;
                //childsParentItem.IsEnabled = false;
            }
            else
            {
                CheckBox childsParentItem = item.Header as CheckBox;
                childsParentItem.IsChecked = false;
                childsParentItem.IsEnabled = true;
            }

        }


        #region Report Preview
        /// <summary>
        /// 
        /// </summary>
        /// <param name="contractId"></param>
        /// <param name="accountId"></param>
        private void ContractBidReportPreview(string contractId, string accountId)
        {
            try
            {
                Nucleus.Agent _agent = new Nucleus.Agent();
                DataTable reportPreviewTable = new DataTable();
                ObjectCache reportCache = MemoryCache.Default;
                reportPreviewTable = _agent.ReportPreview(contractId, accountId);
                ReportDocument contractBidReportPreview = new ReportDocument();
                var path = ("C:\\Users\\darrenm\\Desktop\\ProLogicReportingApplication\\ProLogicReportingApplication\\ContractBidReport.rpt");
                contractBidReportPreview.Load(path);
                contractBidReportPreview.SetDataSource(reportPreviewTable);
                bidContractReportPreview.Owner = Window.GetWindow(this);
                contractBidReportPreview.Refresh();                
                string reportPreviewCache = reportCache["contractBidReportPreview"] as String;                
                bidContractReportPreview.ViewerCore.ReportSource = contractBidReportPreview;
                _agent.ReportPreviewCache(reportPreviewCache, contractId, accountId);
            }
            catch (LogOnException e)
            {
                Console.WriteLine(e.ToString());
            }
            catch (DataSourceException e)
            {
                Console.WriteLine(e.ToString());
            } 
            catch (EngineException engEx)
            {
                Console.WriteLine(engEx.ToString());
            }           
        }
        #endregion

    }
}
