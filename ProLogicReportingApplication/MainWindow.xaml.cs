using System;
using System.IO;
using System.Collections.Generic;
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
using System.Media;
using System.Threading;
using System.Windows.Input;

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
        private static string PATH_REPORTCACHEDIR = @"C:\AgentReportCache\";
        private static string SmtpServer = "smtp.office365.com";
        private static string contractId;
        private static string PATH_CONTRACTBIDREPORT = @"C:\Users\darrenm\Desktop\ProLogicReportingApplication\ProLogicReportingApplication\ContractBidReport.rpt";
        private static BackgroundWorker emailSendWorker = new BackgroundWorker();
        private ObservableCollection<string> proLogic_ContractContactsObservable = new ObservableCollection<string>();
        private ReportDocument contractBidReportPreview = new ReportDocument();
        private List<string> proLogic_ContractContacts = new List<string>();        
        private List<string> proLogic_EmailRecipients = new List<string>();
        private List<string> proLogic_StartActivities = new List<string>();
        private List<Guid> proLogic_ActivityGuids = new List<Guid>();
        private List<Attachment> proLogic_SentProposal = new List<Attachment>();
        private TreeViewItem accountItem = new TreeViewItem();
        private TreeViewItem empItem = new TreeViewItem();
        private TreeViewItem empEmailAddrItem = new TreeViewItem();
        private string emailRecipient;
        private string accountNumAndName;
        private string accountItemTag;
        private string empItemTag;
        

        public MainWindow()
        {
            InitializeComponent();            
            BackgroundWorker _createReportCacheDirWorker = new BackgroundWorker();
            _createReportCacheDirWorker.DoWork += CreateReportCacheDir;
            _createReportCacheDirWorker.RunWorkerAsync();

            //string[] args = Environment.GetCommandLineArgs();
            //MessageBox.Show(args[1]);
            // pass args[1] to LoadContacts 
            // Call to Nucleus to get the data to populate the tree view           
            contractId = "00002";
            LoadContacts(contractId);
        }

        #region Handlers for UI during Report Gen and Emailing
        /// <summary>
        /// Handler for Mouse Down during report gen
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MouseDownDuringReportGen(object sender, MouseEventArgs e)
        {
            MessageBox.Show("Billing Application is generating reports please wait.");
        }

        /// <summary>
        /// Handler for Mouse Up during report gen
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MouseUpDuringReportGen(object sender, MouseEventArgs e)
        {
            MessageBox.Show("Billing Application is generating reports please wait.");
        }

        /// <summary>
        /// Handler for Mouse Down during Emailing
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MouseDownDuringEmailSend(object sender, MouseEventArgs e)
        {
            MessageBox.Show("Billing Application is emailing proposals please wait.");
        }

        /// <summary>
        /// Handler for Mouse Up during Emailing
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MouseUpDuringEmailSend(object sender, MouseEventArgs e)
        {
            MessageBox.Show("Billing Application is emailing proposals please wait.");
        }
        #endregion

        #region Create Cache Dir
        /// <summary>
        /// This method creates the Directory to store generated reports 
        /// if it doesn't exist
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CreateReportCacheDir(object sender, DoWorkEventArgs e)
        {
            
            if (!Directory.Exists(PATH_REPORTCACHEDIR))
            {
                Directory.CreateDirectory(PATH_REPORTCACHEDIR);
            }
        }
        #endregion

        #region Delete Cache Dir Contents
        /// <summary>
        /// This method removes all the generated reports from 
        /// C:\AgentReportCache\ on closing 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ApplicationClosing(object sender, CancelEventArgs e)
        {
            DirectoryInfo cachedFiles = new DirectoryInfo(PATH_REPORTCACHEDIR);
            foreach(FileInfo cachedFile in cachedFiles.GetFiles())
            {
                cachedFile.Delete();
            }
        }
        #endregion      

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

        #region TreeView Load 
        /// <summary>
        /// Tree View - loops over the list returned from Nucleus _agent SELECT statement
        /// Puts each item in ProLogic_zContractContacts into a TreeViewItem
        /// Heirarchy goes Account => ContactFullName -> Contact Email Address. 
        /// I'm also using the account and for the report generation parameters needed are Account and Contract
        /// This gets called from MainWindow.xaml when the TreeView is intially loaded
        /// As there is not a OnClick event for TreeView nor the TreeViewItems I had to use PreviewMouseLeftButtonDown
        /// changed from using ProLogiz_zContractContacts list to zContractsContactsObservable
        /// Creating the email list on load of the contacts here.
        /// This list gets altered based on user interaction with the treeview
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void trvAccount_AccountContactsLoaded(object sender, RoutedEventArgs e)
        {        
            TreeView tree = sender as TreeView;
            int empTracker = 0; 

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
                        empTracker++;
                        empEmailAddrItem = new TreeViewItem();
                        empEmailAddrItem.Tag = proLogic_ContractContactsObservable[i].Replace("{ Header = Item Level 2 }", "").Trim() + "_" + accountItemTag.Replace("{Parent}", "").Trim(); //+ "_" + empTracker;
                        empEmailAddrItem.FontWeight = FontWeights.Black;
                        CheckBox empEmailAddrCheckBox = new CheckBox();
                        empEmailAddrCheckBox.IsChecked = true;
                        empEmailAddrCheckBox.IsEnabled = false;
                        empEmailAddrCheckBox.Content = proLogic_ContractContactsObservable[i].Remove(0, 42).Replace("{ Header = Item Level 2 }", "");                        
                        empEmailAddrCheckBox.Click += mouseClickHandler;
                        empEmailAddrItem.Header = empEmailAddrCheckBox;
                        empItem.Items.Add(empEmailAddrItem);                        
                        proLogic_EmailRecipients.Add(empEmailAddrItem.Tag.ToString());
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
        /// <summary>
        /// The Handler for the treeview Checkbox clicks
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void mouseClickHandler(object sender, EventArgs e)
        {
            NodeCheck(sender as DependencyObject);            
        }
        #endregion        

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
                                SetChildrenChecks(item, true);                                    
                                ContractBidReportPreview(contractId, item.Tag.ToString().Replace("{Parent}", "").Trim().Replace(" ", string.Empty));                                
                            }
                            else
                            {
                                SetChildrenChecks(item, false);
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
                                SetParentChecks((TreeViewItem)item.Parent, true);
                                ContractBidReportPreview(contractId, item.Tag.ToString().Replace("{Child}", "").Trim());
                            }
                            else if (childTreeItemChkBox.IsChecked == false)
                            {
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
                }

                TreeViewItem _emailAddress = item;
                TreeViewItem emailAddress = _emailAddress.ItemContainerGenerator.Items[0] as TreeViewItem;
                TreeViewItem _emailRecipient = emailAddress.ItemContainerGenerator.Items[0] as TreeViewItem;
                emailRecipient = _emailRecipient.Tag.ToString();
                AddEmailRecipient(emailRecipient.Trim());

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
                }

                TreeViewItem _emailAddress = item;
                TreeViewItem emailAddress = _emailAddress.ItemContainerGenerator.Items[0] as TreeViewItem;
                TreeViewItem _emailRecipient = emailAddress.ItemContainerGenerator.Items[0] as TreeViewItem;
                emailRecipient = _emailRecipient.Tag.ToString();
                RemoveEmailRecipient(emailRecipient.Trim());

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
                        emailRecipient = emailAddress.Tag.ToString();
                        AddEmailRecipient(emailRecipient.Trim());
                    }                    
                    parentHasChildItem.IsChecked = true;
                }
                else
                {
                    if(tv.ItemContainerGenerator.Items.Count > 0)
                    {
                        TreeViewItem emailAddress = tv.ItemContainerGenerator.Items[0] as TreeViewItem;
                        emailRecipient = emailAddress.Tag.ToString();                       
                        RemoveEmailRecipient(emailRecipient.Trim());
                    }                    
                    parentHasChildItem.IsChecked = false;
                }                            
            }
        }
        #endregion

        #region Add/Remove Email Recipients
        /// <summary>
        /// Adds recipients to email list
        /// </summary>
        /// <param name="emailAddress"></param>
        private void AddEmailRecipient(string emailAddress)
        {            
             proLogic_EmailRecipients.Add(emailAddress);            
        }

        /// <summary>
        /// Removes recipients from the email list
        /// </summary>
        /// <param name="emailAddress"></param>
        private void RemoveEmailRecipient(string emailAddress)
        {
            if (proLogic_EmailRecipients.Any(s => s.Equals(emailAddress)))
            {
                proLogic_EmailRecipients.Remove(emailAddress);                
            }
        }
        #endregion

        #region Agent Report Cache BackgroundWorker
        /// <summary>
        /// Worker for Caching reports. Calls reportPreviewCacheWorker to cache genereated reports
        /// Also adds handlers for PreviewMouseDown and PreviewMouseUp to stop user ineraction 
        /// until report gen has completed
        /// </summary>
        /// <param name="reportsReadyForCache"></param>
        private void AgentReportCacheWorker(List<KeyValuePair<string, string>> reportsReadyForCache)
        {
            try
            {
                WorkingSpinner.Visibility = Visibility.Visible;
                MainGrid.MouseLeftButtonDown += MouseDownDuringReportGen;
                MainGrid.MouseLeftButtonUp += MouseUpDuringReportGen;
                trvAccount_AccountContacts.PreviewMouseLeftButtonDown += MouseDownDuringReportGen;
                trvAccount_AccountContacts.PreviewMouseLeftButtonUp += MouseUpDuringReportGen;
                BackgroundWorker worker = new BackgroundWorker();
                worker.DoWork += reportPreviewCacheWorker;
                worker.RunWorkerCompleted += removeWorkingSpinner;
                worker.RunWorkerAsync(reportsReadyForCache);
            }
            catch (Exception AgentReportCacheWorkerExeption)
            {
                MessageBox.Show(AgentReportCacheWorkerExeption.ToString());
            }            
        }

        /// <summary>
        /// worker completed method removes the working spinner after reports have 
        /// completed generating.  Also removes the PreviewMouseDown and PreviewMouseUp Handlers 
        /// for the maingrid and treeview
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void removeWorkingSpinner(object sender, RunWorkerCompletedEventArgs e)
        {
            WorkingSpinner.Visibility = Visibility.Hidden;
            MainGrid.MouseLeftButtonDown -= MouseDownDuringReportGen;
            MainGrid.MouseLeftButtonUp -= MouseUpDuringReportGen;
            trvAccount_AccountContacts.PreviewMouseLeftButtonDown -= MouseDownDuringReportGen;
            trvAccount_AccountContacts.PreviewMouseLeftButtonUp -= MouseUpDuringReportGen;
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
                contractBidReportPreviewCache.Load(PATH_CONTRACTBIDREPORT);
                contractBidReportPreviewCache.SetDataSource(reportPreviewCacheTable);
                contractBidReportPreviewCache.Refresh();
                contractBidReportPreviewCache.ExportToDisk(ExportFormatType.CrystalReport, PATH_REPORTCACHEDIR + contractId + accountId.Replace(" ", string.Empty) + ".rpt");
                contractBidReportPreviewCache.ExportToDisk(ExportFormatType.PortableDocFormat, PATH_REPORTCACHEDIR + contractId + accountId.Replace(" ", string.Empty) + ".pdf");
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
                var cachedReportFile = Directory.GetFiles(PATH_REPORTCACHEDIR, contractId + accountId + ".rpt");
                if(cachedReportFile != null)
                {                    
                    string path = (PATH_REPORTCACHEDIR + contractId + accountId + ".rpt");
                    contractBidReportPreview.Load(path);
                    bidContractReportPreview.ViewerCore.ReportSource = contractBidReportPreview;
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
        /// Also starts the working spinner. Also adds handlers for PreviewMouseDown and PreviewMouseUp 
        /// to stop user ineraction until report gen has completed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_SendBid_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MessageBox.Show("Billing Application is Emaling Reports.  Please Wait.");
                WorkingSpinner.Visibility = Visibility.Visible;
                MainGrid.MouseLeftButtonDown += MouseDownDuringEmailSend;
                MainGrid.MouseLeftButtonUp += MouseUpDuringEmailSend;
                trvAccount_AccountContacts.PreviewMouseLeftButtonDown += MouseDownDuringEmailSend;
                trvAccount_AccountContacts.PreviewMouseLeftButtonUp += MouseUpDuringEmailSend;
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
            string _to;
            string _startActivity;

            try
            {                
                for (int i = 0; i < proLogic_EmailRecipients.Count; i++)
                {                    
                    accountNumAndName = proLogic_EmailRecipients[i].Substring(proLogic_EmailRecipients[i].LastIndexOf('_') + 1);

                    MailMessage msg = new MailMessage();
                    msg.Subject = "Bid Proposal";
                    msg.From = new MailAddress("bob@360sheetmetal.com");
                    _to = proLogic_EmailRecipients[i].Remove(0, 42);
                    msg.To.Add(new MailAddress(_to.Substring(0, _to.LastIndexOf("_"))));                    
                    msg.Body = "Email Sent from Bid Report Application";                    
                    Attachment bidProposal = new Attachment(PATH_REPORTCACHEDIR + contractId + accountNumAndName + ".pdf");
                    bidProposal.Name = "Bid Proposal - Job Name: " + accountNumAndName.Remove(0, 4) + ".pdf";
                    msg.Attachments.Add(bidProposal);

                    SmtpClient smtp = new SmtpClient(SmtpServer);
                    smtp.Port = 587;
                    smtp.EnableSsl = true;
                    smtp.UseDefaultCredentials = false;
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                    smtp.Credentials = new NetworkCredential("darrenm@360sheetmetal.com", "L14ei5g00d360");
                    smtp.Send(msg);
                    SystemSounds.Exclamation.Play();

                    _startActivity = proLogic_EmailRecipients[i].Remove(0, 5);
                    proLogic_StartActivities.Add(_startActivity.Substring(0, _startActivity.IndexOf("_")));
                    proLogic_SentProposal.Add(bidProposal);
                }                
            }
            catch (Exception SendEmailException)
            {
                MessageBox.Show(SendEmailException.ToString());
            }            
        }
        #endregion

        #region Post To Syspro
        /// <summary>
        /// Removes the working spinner
        /// passes the proLogic_ActivityGuids and proLogic_SentProposal to Nuclues.Agent
        /// for XML String Building
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PostToSyspro(object sender, RunWorkerCompletedEventArgs e)
        {
            WorkingSpinner.Visibility = Visibility.Hidden;
            MessageBox.Show("Emails Sent Successfully");
            MainGrid.MouseLeftButtonDown -= MouseDownDuringEmailSend;
            MainGrid.MouseLeftButtonUp -= MouseUpDuringEmailSend;
            trvAccount_AccountContacts.PreviewMouseLeftButtonDown -= MouseDownDuringEmailSend;
            trvAccount_AccountContacts.PreviewMouseLeftButtonUp -= MouseUpDuringEmailSend;
            Nucleus.Agent _agent = new Nucleus.Agent();
            proLogic_ActivityGuids = proLogic_StartActivities.ConvertAll<Guid>(x => new Guid(x));
            _agent.PostXmlForSyspro(proLogic_ActivityGuids, proLogic_SentProposal);
        }
        #endregion        
    }
}
