using System;
using System.IO;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Data;
using System.Collections.ObjectModel;
using System.Net;
using System.Net.Mail;
using CrystalDecisions.Shared;
using CrystalDecisions.CrystalReports.Engine;
using System.Windows.Data;

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
        private List<string> proLogic_EmailRecipients = new List<string>();
        private static string contractId;
        private static string ReportCacheDir = @"C:\AgentReportCache\";
        private static string SmtpServer = "smtp.office365.com";
        private static string currentReport;
        private ReportDocument contractBidReportPreview = new ReportDocument();
        private string emailRecipient;


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
            try
            {
                List<KeyValuePair<string, string>> ToBeCachedReports = new List<KeyValuePair<string, string>>();
                Nucleus.Agent _agent = new Nucleus.Agent();
                _agent.GetContacts(contractId);
                // Adds the list from Nucleus.Agent                      
                proLogic_ContractContacts.AddRange(_agent.Agent_ContractContactsResponse);
                // Copies ProLogic_zContractContacts to an observable collection
                // This is to be used for the treeview           
                proLogic_ContractContactsObservable = new ObservableCollection<string>(proLogic_ContractContacts);
                foreach (var item in proLogic_ContractContactsObservable)
                {
                    if (item.Contains("{ Header = Item Level 0 }"))
                    {
                        string accountId = item.Substring(0, item.IndexOf("{"));
                        KeyValuePair<string, string> myItem = new KeyValuePair<string, string>(contractId, accountId);
                        ToBeCachedReports.Add(myItem);
                    }
                }
                AgentReportCacheWorker(ToBeCachedReports);
            }
            catch(Exception LoadContactsException)
            {
                MessageBox.Show(LoadContactsException.ToString());
            }
            
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
            string accountItemTag = null;
            string empItemTag = null;

            try
            {
                for (int i = 0; i < proLogic_ContractContactsObservable.Count; i++)
                {
                    //AccountName = Level 0                       
                    if (proLogic_ContractContactsObservable[i].Contains("{ Header = Item Level 0 }"))
                    {
                        accountItem = new TreeViewItem();
                        // Removing everything after the account ID for the Tag
                        accountItem.Tag = "{Parent} " + proLogic_ContractContactsObservable[i].Substring(0, proLogic_ContractContactsObservable[i].IndexOf("{"));
                        accountItemTag = accountItem.Tag.ToString().Replace(" ", string.Empty);
                        accountItem.IsExpanded = true;
                        accountItem.FontWeight = FontWeights.Black;
                        CheckBox accountCheckBox = new CheckBox();
                        accountCheckBox.IsChecked = true;
                        accountCheckBox.IsEnabled = true;
                        accountCheckBox.Focusable = true;                                                                   
                        //accountCheckBox.IsThreeState = true;
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
                        empItemTag = "{Child} " + proLogic_ContractContactsObservable[i].Remove(5).TrimEnd() + accountItemTag.Remove(0, 12);
                        empItem.Tag = empItemTag;
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
                        empEmailAddrCheckBox.Content = proLogic_ContractContactsObservable[i].Remove(0, 4).Replace("{ Header = Item Level 2 }", "");
                        empEmailAddrCheckBox.Click += mouseClickHandler;
                        empEmailAddrItem.Header = empEmailAddrCheckBox;
                        empItem.Items.Add(empEmailAddrItem);
                    }
                }
            }
            catch (Exception trvAccount_AccountContactsLoadedException)
            {
                MessageBox.Show(trvAccount_AccountContactsLoadedException.ToString());
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
                                ContractBidReportPreview(contractId, item.Tag.ToString().Replace("{Parent}", "").Trim().Replace(" ", string.Empty));                                
                            }
                            else if (parentTreeItemChkBox.IsChecked == false)
                            {
                                Console.WriteLine("Parent Check Box is Unchecked -> " + parentTreeItemChkBox);
                                SetChildrenChecks(item, false);
                                ContractBidReportPreview(contractId, item.Tag.ToString().Replace("{Parent}", "").Trim().Replace(" ", string.Empty));
                            }
                            else if (parentTreeItemChkBox.IsChecked == null)
                            {
                                ContractBidReportPreview(contractId, item.Tag.ToString().Replace("{Parent}", "").Trim().Replace(" ", string.Empty));
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
            catch (Exception NodeCheckException)
            {
                MessageBox.Show(NodeCheckException.ToString());
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
                List<CheckBox> childCheckBoxList = new List<CheckBox>();
                CheckBox childsParentChkbox = new CheckBox();
                CheckBox childsParentItem = item.Header as CheckBox;
                foreach (TreeViewItem c in item.Items)
                {
                    childsParentChkbox = c.Header as CheckBox;                    
                    childCheckBoxList.Add(childsParentChkbox);
                    TreeViewItem emailAddress = c.ItemContainerGenerator.Items[0] as TreeViewItem;
                    emailRecipient = emailAddress.Header.ToString();
                    if (emailRecipient.Contains("IsChecked:True"))
                    {
                        BackgroundWorker worker = new BackgroundWorker();
                        worker.DoWork += RemoveEmailRecipient;
                        worker.RunWorkerAsync(emailRecipient.Remove(0, 42).Replace("IsChecked:True", "").Trim());
                    }
                    if (emailRecipient.Contains("IsChecked:False"))
                    {
                        BackgroundWorker worker = new BackgroundWorker();
                        worker.DoWork += AddEmailRecipient;
                        worker.RunWorkerAsync(emailRecipient.Remove(0, 42).Replace("IsChecked:True", "").Trim());
                    }
                }
                if (childCheckBoxList.All(a => a.IsChecked == true))
                {
                    childsParentItem.IsChecked = true;
                    childsParentItem.IsEnabled = true;
                    childsParentItem.Focusable = true;                    
                }
                else
                {
                    childsParentItem.IsChecked = true;
                    childsParentItem.Focusable = true;
                }
            }
            else //checkedState == false
            {
                List<CheckBox> childCheckBoxList = new List<CheckBox>();
                CheckBox childsParentChkbox = new CheckBox();
                CheckBox childsParentItem = item.Header as CheckBox;
                foreach (TreeViewItem c in item.Items)
                {
                    childsParentChkbox = c.Header as CheckBox;
                    childCheckBoxList.Add(childsParentChkbox);
                    TreeViewItem emailAddress = c.ItemContainerGenerator.Items[0] as TreeViewItem;
                    emailRecipient = emailAddress.Header.ToString();
                    if(emailRecipient.Contains("IsChecked:True"))
                    {
                        BackgroundWorker worker = new BackgroundWorker();
                        worker.DoWork += RemoveEmailRecipient;
                        worker.RunWorkerAsync(emailRecipient.Remove(0, 42).Replace("IsChecked:True", "").Trim());
                    }
                    if (emailRecipient.Contains("IsChecked:False"))
                    {
                        BackgroundWorker worker = new BackgroundWorker();
                        worker.DoWork += AddEmailRecipient;
                        worker.RunWorkerAsync(emailRecipient.Remove(0, 42).Replace("IsChecked:True", "").Trim());
                    }
                           
                }
                if (childCheckBoxList.Any(a => a.IsChecked == false))
                {
                    childsParentItem.IsChecked = true;
                    childsParentItem.Focusable = true;
                }
                if(childCheckBoxList.All(a => a.IsChecked == false))
                {
                    childsParentItem.IsChecked = false;
                    childsParentItem.IsEnabled = true;
                    childsParentItem.Focusable = true;
                }
            }
        }

        /// <summary>
        /// Item is a treeviewitem being passed in.
        /// </summary>
        /// <param name="item"></param>
        /// <param name="checkedState"></param>
        private void SetChildrenChecks(TreeViewItem item, bool checkedState)
        {
            
            foreach (TreeViewItem tv in item.Items)
            {
                CheckBox parentHasChildItem = tv.Header as CheckBox;
                if (checkedState == true)
                {
                    if(tv.ItemContainerGenerator.Items.Count > 0)
                    {
                        TreeViewItem emailAddress = tv.ItemContainerGenerator.Items[0] as TreeViewItem;
                        emailRecipient = emailAddress.Header.ToString();
                        BackgroundWorker worker = new BackgroundWorker();
                        worker.DoWork += AddEmailRecipient;
                        worker.RunWorkerAsync(emailRecipient.Remove(0, 42).Replace("IsChecked:True", "").Trim());
                    }                    
                    parentHasChildItem.IsChecked = true;
                }
                else
                {
                    if(tv.ItemContainerGenerator.Items.Count > 0)
                    {
                        TreeViewItem emailAddress = tv.ItemContainerGenerator.Items[0] as TreeViewItem;
                        emailRecipient = emailAddress.Header.ToString();
                        BackgroundWorker worker = new BackgroundWorker();
                        worker.DoWork += RemoveEmailRecipient;
                        worker.RunWorkerAsync(emailRecipient.Remove(0, 42).Replace("IsChecked:True", "").Trim());
                    }                    
                    parentHasChildItem.IsChecked = false;
                }                            
            }
        }

        private void AddEmailRecipient(object sender, DoWorkEventArgs e)
        {
            if (!proLogic_EmailRecipients.Any(s => s.Equals(e.Argument)))
            {
                proLogic_EmailRecipients.Add(e.Argument.ToString());
            }
        }

        private void RemoveEmailRecipient(object sender, DoWorkEventArgs e)
        {
            if(proLogic_EmailRecipients.Any(s => s.Equals(e.Argument)))
            {
                proLogic_EmailRecipients.Remove(e.Argument.ToString());
            }
        }
        #endregion

        #region Agent Report Cache BackgroundWorker
        /// <summary>
        /// Worker for Caching reports. Calls reportPreviewCacheWorker to cache genereated reports
        /// </summary>
        /// <param name="reportsReadyForCache"></param>
        private void AgentReportCacheWorker(List<KeyValuePair<string, string>> reportsReadyForCache)
        {
            try
            {
                BackgroundWorker worker = new BackgroundWorker();
                worker.DoWork += reportPreviewCacheWorker;
                //worker.RunWorkerCompleted += startGarbageCollector;
                worker.RunWorkerAsync(reportsReadyForCache);
            }
            catch (Exception AgentReportCacheWorkerExeption)
            {
                MessageBox.Show(AgentReportCacheWorkerExeption.ToString());
            }            
        }

        /// <summary>
        /// Background worker event handler 
        /// Calls ContractBidReportCache to store the generated reports on load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void reportPreviewCacheWorker(object sender, DoWorkEventArgs e)
        {
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
            try
            {
                Nucleus.Agent _agent = new Nucleus.Agent();
                DataTable reportPreviewCacheTable = new DataTable();
                reportPreviewCacheTable = _agent.ReportPreview(contractId, accountId.Remove(5).Trim());
                ReportDocument contractBidReportPreviewCache = new ReportDocument();
                var path = (@"C:\Users\darrenm\Desktop\ProLogicReportingApplication\ProLogicReportingApplication\ContractBidReport.rpt");
                contractBidReportPreviewCache.Load(path);
                contractBidReportPreviewCache.SetDataSource(reportPreviewCacheTable);
                contractBidReportPreviewCache.Refresh();
                contractBidReportPreviewCache.ExportToDisk(ExportFormatType.CrystalReport, ReportCacheDir + contractId + accountId.Replace(" ", string.Empty) + ".rpt");
                contractBidReportPreviewCache.ExportToDisk(ExportFormatType.PortableDocFormat, ReportCacheDir + contractId + accountId.Replace(" ", string.Empty) + ".pdf");
            }
            catch (Exception ContractBidReportCacheException)
            {
                MessageBox.Show(ContractBidReportCacheException.ToString());
            }
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
                    string path = (ReportCacheDir + contractId + accountId + ".rpt");
                    currentReport = ReportCacheDir + contractId + accountId + ".pdf";
                    contractBidReportPreview.Load(path);
                    bidContractReportPreview.ViewerCore.ReportSource = contractBidReportPreview;
                }
                else
                {
                    //generate report here
                }               
            }
            catch (LogOnException ContractBidReportPreview_LogOnException)
            {
                MessageBox.Show(ContractBidReportPreview_LogOnException.ToString());
            }
            catch (DataSourceException ContractBidReportPreview_DataSourceException)
            {
                MessageBox.Show(ContractBidReportPreview_DataSourceException.ToString());
            } 
            catch (EngineException ContractBidReportPreview_EngineException)
            {
                MessageBox.Show(ContractBidReportPreview_EngineException.ToString());
            }
            catch (OutOfMemoryException ContractBidReportPreview_OutOfMemoryException)
            {
                MessageBox.Show(ContractBidReportPreview_OutOfMemoryException.ToString());
            }           
        }
        #endregion

        #region Send Bid Click
        /// <summary>
        /// Starts a background worker thread SendEmail  
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_SendBid_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                BackgroundWorker emailSendWorker = new BackgroundWorker();
                emailSendWorker.DoWork += SendEmail;
                emailSendWorker.RunWorkerCompleted += PostToSyspro;
                emailSendWorker.RunWorkerAsync();
            }
            catch (Exception SendBindClickException)
            {
                MessageBox.Show(SendBindClickException.ToString());
            }
        }
        #endregion 

        #region Email Send
        /// <summary>
        /// Send a email to a list of recipients
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SendEmail(object sender, DoWorkEventArgs e)
        {
            try
            {
                MailMessage msg = new MailMessage();
                msg.Subject = "Testing Email";
                msg.From = new MailAddress("darrenm@360sheetmetal.com");
                msg.To.Add(new MailAddress("darrenm@360sheetmetal.com"));
                msg.Body = "Email Sent from Bid Report Application";
                Attachment bidProposal = new Attachment(currentReport);
                bidProposal.Name = "Bid Proposal - Job Name: " + currentReport.Remove(0, 29);
                msg.Attachments.Add(bidProposal);

                SmtpClient smtp = new SmtpClient(SmtpServer);
                smtp.Port = 587;
                smtp.EnableSsl = true;
                smtp.UseDefaultCredentials = false;
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.Credentials = new NetworkCredential("darrenm@360sheetmetal.com", "L14ei5g00d360");
                smtp.Send(msg);
                MessageBox.Show("Mail Sent");
            }
            catch (Exception SendEmailException)
            {
                MessageBox.Show(SendEmailException.ToString());
            }
        }
        #endregion

        #region Post To Syspro
        private void PostToSyspro(object sender, RunWorkerCompletedEventArgs e)
        {
            Nucleus.Agent _agent = new Nucleus.Agent();
            _agent.PostXmlForSyspro();
        }
        #endregion       
    }
}
