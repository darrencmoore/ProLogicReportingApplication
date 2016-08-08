using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Drawing;
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
using Encore;

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
        //private Timer _mouseClickTimer = null;
        //private int userClicks;
        private static string contractId;
        private static string ReportCacheDir = @"C:\AgentReportCache\";
        private static string SMTPserver = "smtp.office365.com";
        private static string currentReport;

        public MainWindow()
        {
            InitializeComponent();           
            //string[] args = Environment.GetCommandLineArgs();
            //MessageBox.Show(args[1]);
            // pass args[1] to LoadContacts 
            // Call to Nucleus to get the data to populate the tree view            
            contractId = "00002";
            LoadContacts(contractId);
        }

        #region Load Contacts
        /// <summary>
        /// Gets called on initialization takes the Contract passed from SYSPRO
        /// This method also calls on Nucleus.dll for the GetContacts() 
        /// Then loops over the proLogic_ContractContactsObservable and adds contract and account as a keypair
        /// </summary>
        /// <param name="contractId"></param>
        /// <returns></returns>
        public string LoadContacts(string contractId)
        {            
            List<KeyValuePair<string, string>> ToBeCachedReports = new List<KeyValuePair<string, string>>();
            Nucleus.Agent _agent = new Nucleus.Agent();
            _agent.GetContacts(contractId);
            // Adds the list from Nucleus.Agent                      
            proLogic_ContractContacts.AddRange(_agent.Agent_ContractContactsResponse);            
            // Copies ProLogic_zContractContacts to an observable collection
            // This is to be used for the treeview           
            proLogic_ContractContactsObservable = new ObservableCollection<string>(proLogic_ContractContacts);
            foreach(var item in proLogic_ContractContactsObservable)
            {
                if(item.Contains("{ Header = Item Level 0 }"))
                {                    
                    string accountId = item.Remove(4);
                    KeyValuePair<string, string> myItem = new KeyValuePair<string, string>(contractId, accountId);
                    ToBeCachedReports.Add(myItem);
                }
            }
            AgentReportCacheWorker(ToBeCachedReports);
            
            return null;
        }
        #endregion

        

        /// <summary>
        /// This will handle Selected Item Changes
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void trvTree_Collapsed(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            Console.WriteLine("TreeView Collapsed");
        }

        #region Email Send
        /// <summary>
        /// Bid Email Send
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_SendBid_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Encore.Utilities dll = new Encore.Utilities();
                //dll.Logon("ADMIN", " ", "TEST [TEST FOR 360 SHEET METAL LLC]", " ", Encore.Language.ENGLISH, 0, 0, " ");
                //good below here
                MailMessage msg = new MailMessage();
                msg.Subject = "Testing Email";
                msg.From = new MailAddress("darrenm@360sheetmetal.com");
                msg.To.Add(new MailAddress("darrenm@360sheetmetal.com"));
                msg.Body = "Email Sent from Bid Report Application";
                Attachment bidReport = new Attachment(currentReport + ".pdf"); // rename to bid proposal before sending email
                msg.Attachments.Add(bidReport);

                SmtpClient smtp = new SmtpClient(SMTPserver);
                //smtp.Host = "smtp.office365.com";
                smtp.Port = 587;
                smtp.EnableSsl = true;
                smtp.UseDefaultCredentials = false;
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.Credentials = new NetworkCredential("darrenm@360sheetmetal.com", "L14ei5g00d360");
                smtp.Send(msg);
                MessageBox.Show("Mail Sent");
                //good above here
                //// Add smtp email stuff here
                //SmtpClient MyServer = new SmtpClient();
                //MyServer.Host = "smtp.office365.com";
                //MyServer.Port = 587;
                //MyServer.EnableSsl = true;
                ////Server Credentials
                //NetworkCredential NC = new NetworkCredential();
                //NC.UserName = "";
                //NC.Password = "";
                ////assigned credetial details to server
                //MyServer.Credentials = NC;
                //MailAddress from = new MailAddress("darrenm@360sheetmetal.com", "Darren");
                //MailAddress receiver = new MailAddress("darrenm@360sheetmetal.com", "ME");
                //MailMessage Mymessage = new MailMessage(from, receiver);
                //Mymessage.Subject = "test";
                //Mymessage.Body = "test";
                ////sends the email
                //MyServer.Send(Mymessage);
            }
            catch (Exception es)
            {
                Console.WriteLine(es.ToString());
            }

        }
        #endregion

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
            TreeView tree = sender as TreeView;
            

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
                    CheckBox accountCheckBox = new CheckBox();
                    accountCheckBox.IsChecked = true;
                    accountCheckBox.IsEnabled = true;
                    accountCheckBox.Focusable = true;
                    accountCheckBox.IsThreeState = true;
                    accountCheckBox.Name = "ParentChkBox";
                    accountCheckBox.Content = proLogic_ContractContactsObservable[i].Remove(0, 4).Replace("{ Header = Item Level 0 }", "");
                    accountCheckBox.Click += mouseClickHandler;
                    accountItem.Header = accountCheckBox;
                    tree.Items.Add(accountItem);                                     
                }
                //ContactFullName = Level 1
                if (proLogic_ContractContactsObservable[i].Contains("{ Header = Item Level 1 }"))
                {                    
                    empItem = new TreeViewItem();
                    // Removing everything after the account ID for the Tag
                    empItem.Tag = "{Child} " + proLogic_ContractContactsObservable[i].Remove(5);
                    empItem.FontWeight = FontWeights.Black;
                    CheckBox empItemCheckBox = new CheckBox();
                    empItemCheckBox.IsChecked = true;
                    empItemCheckBox.IsEnabled = true;
                    empItemCheckBox.Focusable = true;
                    empItemCheckBox.Name = "ChildChkBox";
                    empItemCheckBox.Content = proLogic_ContractContactsObservable[i].Remove(0, 4).Replace("{ Header = Item Level 1 }", "");
                    empItemCheckBox.Click += mouseClickHandler;
                    empItem.Header = empItemCheckBox;                   
                    accountItem.Items.Add(empItem);
                }
                //Contact Email Address = Level 2
                if (proLogic_ContractContactsObservable[i].Contains("{ Header = Item Level 2 }"))
                {                    
                    empEmailAddrItem = new TreeViewItem();
                    empEmailAddrItem.Tag = "{Email} " + proLogic_ContractContactsObservable[i].Replace("{ Header = Item Level 2 }", "");
                    empEmailAddrItem.FontWeight = FontWeights.Black;
                    CheckBox empEmailAddrCheckBox = new CheckBox();
                    empEmailAddrCheckBox.IsChecked = true;
                    empEmailAddrCheckBox.IsEnabled = false;
                    empEmailAddrCheckBox.Content = proLogic_ContractContactsObservable[i].Replace("{ Header = Item Level 2 }", "");
                    empEmailAddrCheckBox.Click += mouseClickHandler;
                    empEmailAddrItem.Header = empEmailAddrCheckBox;
                    empItem.Items.Add(empEmailAddrItem);                    
                }                
            }                    
        }
        #endregion

        #region Mouse Click Handler
        private void mouseClickHandler(object sender, EventArgs e)
        {
            NodeCheck(sender as DependencyObject);           
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
                                //Console.WriteLine("Parent Check Box Checked  -> " + parentTreeItemChkBox);
                                Console.WriteLine("LIst Count" + proLogic_ContractContactsObservable.Count());
                                SetChildrenChecks(item, true);                                    
                                ContractBidReportPreview(contractId, item.Tag.ToString().Replace("{Parent}", "").Trim());                                
                            }
                            else if (parentTreeItemChkBox.IsChecked == false)
                            {
                                Console.WriteLine("Parent Check Box is Unchecked -> " + parentTreeItemChkBox);
                                //SetChildrenChecks(item, true);
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
                                Console.WriteLine("Child Check Box Checked  -> " + childTreeItemChkBox);
                                Console.WriteLine("Parent -> " + item.Parent.ToString());                                
                                SetParentChecks((TreeViewItem)item.Parent, true);
                                ContractBidReportPreview(contractId, item.Tag.ToString().Replace("{Child}", "").Trim());
                            }
                            else if (childTreeItemChkBox.IsChecked == false)
                            {
                                Console.WriteLine("Child Check Box is Unchecked -> " + childTreeItemChkBox);
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

        #region Parent/Child Checkbox Manipulation
        /// <summary>
        /// Item is a treeviewitem being passed in.
        /// </summary>
        /// <param name="item"></param>
        /// <param name="checkedState"></param>
        private void SetParentChecks(TreeViewItem item, bool checkedState)
        {
            if (checkedState == true)
            {
                Console.WriteLine("SetParentChecks STATE -> " + checkedState.ToString());
                CheckBox childsParentItem = item.Header as CheckBox;
                childsParentItem.IsChecked = true;
                childsParentItem.IsEnabled = true;
            }
            else
            {
                Console.WriteLine("SetParentChecks STATE -> " + checkedState.ToString());
                CheckBox childsParentItem = item.Header as CheckBox;
                childsParentItem.IsChecked = null;
                childsParentItem.IsEnabled = true;
            }
        }

        /// <summary>
        /// Item is a treeviewitem being passed in.
        /// </summary>
        /// <param name="item"></param>
        /// <param name="checkedState"></param>
        private void SetChildrenChecks(TreeViewItem item, bool checkedState)
        {
            foreach(TreeViewItem tv in item.Items)
            {
                if(checkedState == true)
                {
                    Console.WriteLine("SetChildrenChecks STATE -> " + checkedState.ToString());
                    CheckBox parentHasChildItem = tv.Header as CheckBox;
                    parentHasChildItem.IsChecked = true;
                }
                else
                {
                    Console.WriteLine("SetChildrenChecks STATE -> " + checkedState.ToString());
                    CheckBox parentHasChildItem = tv.Header as CheckBox;
                    parentHasChildItem.IsChecked = null;
                    //parentHasChildItem.IsEnabled = false;
                }               
            }
        }        
        #endregion

        #region Agent Report Cache Background Worker
        /// <summary>
        /// Worker for Caching reports. Calls reportPreviewCacheWorker to cache genereated reports
        /// </summary>
        /// <param name="reportsReadyForCache"></param>
        private void AgentReportCacheWorker(List<KeyValuePair<string, string>> reportsReadyForCache)
        {
            BackgroundWorker worker = new BackgroundWorker();            
            worker.DoWork += reportPreviewCacheWorker;
            //worker.RunWorkerCompleted += startGarbageCollector;
            worker.RunWorkerAsync(reportsReadyForCache);
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void reportPreviewCacheWorker(object sender, DoWorkEventArgs e)
        {
            Nucleus.Agent _agent = new Nucleus.Agent();
            List<KeyValuePair<string, string>> reportsReadyForCache = (List<KeyValuePair<string, string>>)e.Argument;
            foreach(KeyValuePair<string, string> combo_contractaccount in reportsReadyForCache)
            {
                ContractBidReportCache(combo_contractaccount.Key, combo_contractaccount.Value);
            }
        }        
        #endregion

        #region Application Report Cache
        /// <summary>
        /// Saves reports the get generated in background thread worker
        /// </summary>
        /// <param name="contractId"></param>
        /// <param name="accountId"></param>
        private void ContractBidReportCache(string contractId, string accountId)
        {
            Nucleus.Agent _agent = new Nucleus.Agent();
            DataTable reportPreviewCacheTable = new DataTable();            
            reportPreviewCacheTable = _agent.ReportPreview(contractId, accountId);
            ReportDocument contractBidReportPreviewCache = new ReportDocument();
            var path = ("C:\\Users\\darrenm\\Desktop\\ProLogicReportingApplication\\ProLogicReportingApplication\\ContractBidReport.rpt");
            contractBidReportPreviewCache.Load(path);
            contractBidReportPreviewCache.SetDataSource(reportPreviewCacheTable);
            contractBidReportPreviewCache.Refresh();            
            contractBidReportPreviewCache.ExportToDisk(ExportFormatType.CrystalReport, ReportCacheDir + contractId + accountId + ".rpt" );
            contractBidReportPreviewCache.ExportToDisk(ExportFormatType.PortableDocFormat, ReportCacheDir + "Bid Proposal" + "_" + contractId + accountId + ".pdf");
        }
        #endregion

        #region Application Report Preview
        /// <summary>
        /// Pulls the generated Report from C:\\AgentReportCache
        /// </summary>
        /// <param name="contractId"></param>
        /// <param name="accountId"></param>
        private void ContractBidReportPreview(string contractId, string accountId)
        {
            try
            {
                var cachedReportFile = Directory.GetFiles(ReportCacheDir, contractId + accountId + ".rpt");
                if(cachedReportFile != null)
                {
                    ReportDocument contractBidReportPreview = new ReportDocument();
                    string path = (ReportCacheDir + contractId + accountId + ".rpt");
                    currentReport = ReportCacheDir + contractId + accountId;
                    contractBidReportPreview.Load(path);                    
                    bidContractReportPreview.ViewerCore.ReportSource = contractBidReportPreview;
                }
                else
                {
                    //generate report here
                }                
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

        //private void CheckBox_Click(object sender, RoutedEventArgs e) { OnCheck(); }
        //public void OnCheck()
        //{
        //    Encore.Utilities dll = new Encore.Utilities();
        //}
    }
}
