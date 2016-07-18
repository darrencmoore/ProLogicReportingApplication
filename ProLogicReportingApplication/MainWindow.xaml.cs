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

namespace ProLogicReportingApplication
{    
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window 
    {        
        public List<String> ProLogic_zContractContacts = new List<String>();        
        public ObservableCollection<String> zContractsContactsObservable = new ObservableCollection<String>();
        private Timer ClickTimer = null;      
        
        
        public MainWindow()
        {
            InitializeComponent();            
            //string[] args = Environment.GetCommandLineArgs();
            //MessageBox.Show(args[1]);
            // Call to Nucleus to get the data to populate the tree view
            LoadAccount_AccountContacts("00001");
            //this.ContactsGrid.ItemsSource = ProLogic_zContractContacts;
            //ContactsGrid.ItemsSource = ProLogic_zContractContacts.ToList();
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
            // Adds the list from Nucleus.Agent                      
            ProLogic_zContractContacts.AddRange(_agent.getProLogic_zContractContacts);
            // Copies ProLogic_zContractContacts to an observable collection
            // This is to be used for the treeview           
            zContractsContactsObservable = new ObservableCollection<String>(ProLogic_zContractContacts);

            return null;
        }

        #region TreeViewLoadMethod
        /// <summary>
        /// Tree View - loops over the list returned from Nucleus _agent SELECT statement
        /// Puts each item in ProLogic_zContractContacts into a TreeViewItem
        /// Heirarchy goes Account => ContactFullName -> Contact Email Address
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

            for (int i = 0; i < zContractsContactsObservable.Count; i++)
            {
                //AccountName = Level 0                       
                if (zContractsContactsObservable[i].Contains("{ Header = Item Level 0 }"))
                {                 
                    accountItem = new TreeViewItem(); 
                    accountItem.Tag = "{Parent} " + zContractsContactsObservable[i].Replace("{ Header = Item Level 0 }", "");
                    accountItem.IsExpanded = true;                    
                    accountItem.Header = new CheckBox() 
                    {
                        IsChecked = true,
                        IsEnabled = true,                                                                                                                            
                        Content = zContractsContactsObservable[i].Replace("{ Header = Item Level 0 }", "")
                    };                    
                    tree.Items.Add(accountItem);
                }
                //ContactFullName = Level 1
                if (zContractsContactsObservable[i].Contains("{ Header = Item Level 1 }"))
                {                    
                    empItem = new TreeViewItem();
                    empItem.Tag = "{Child} " + zContractsContactsObservable[i].Replace("{ Header = Item Level 1 }", "");
                    empItem.Header = new CheckBox()
                    {
                        IsChecked = true,
                        IsEnabled = true,
                        Content = zContractsContactsObservable[i].Replace("{ Header = Item Level 1 }", "")
                    };
                    accountItem.Items.Add(empItem);
                }
                //Contact Email Address = Level 2
                if (zContractsContactsObservable[i].Contains("{ Header = Item Level 2 }"))
                {                    
                    empEmailAddrItem = new TreeViewItem();
                    empEmailAddrItem.Tag = "{Email} " + zContractsContactsObservable[i].Replace("{ Header = Item Level 2 }", "");
                    empEmailAddrItem.Header = new CheckBox()
                    {
                        IsChecked = true,
                        IsEnabled = false,               
                        Content = zContractsContactsObservable[i].Replace("{ Header = Item Level 2 }", "")
                    };
                    empItem.Items.Add(empEmailAddrItem);
                }
            }
        }
        #endregion       

        #region Mouse Click Event Handlers                    
        private void trvMouse_SingleClick(object sender, RoutedEventArgs e) //RoutedEventArgs
        {
            e.Handled = true;        
            Console.WriteLine("you clicked once");
            Console.WriteLine("trvMouse_SingleClick =>  OriginalSource -> " + e.OriginalSource);
            Console.WriteLine("trvMouse_SingleClick => RoutedEvent -> " + e.RoutedEvent);
            Console.WriteLine("trvMouse_SingleClick => Source -> " + e.Source);
        }

        /// <summary>
        /// Handler for trvAccount_AccountContacts item click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void trvMouse_DoubleClick(object sender, RoutedEventArgs e) //RoutedEventArgs
        {            
            e.Handled = true;
            Console.WriteLine("you clicked twice");
            Console.WriteLine("trvMouse_DoubleClick =>  OriginalSource -> " + e.OriginalSource);
            Console.WriteLine("trvMouse_DoubleClick => RoutedEvent -> " + e.RoutedEvent);
            Console.WriteLine("trvMouse_DoubleClick => Source -> " + e.Source);
        }       

        private void trvRightMouseButton_Click(object sender, RoutedEventArgs e) //RoutedEventArgs
        {
            e.Handled = true;
            Console.WriteLine("trvRightMouseButton_Click =>  OriginalSource -> " + e.OriginalSource);
            Console.WriteLine("trvRightMouseButton_Click => RoutedEvent -> " + e.RoutedEvent);
            Console.WriteLine("trvRightMouseButton_Click => Source -> " + e.Source);
        }
        #endregion
        
        #region Preview Mouse Click Events
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void previewMouse_SingleClick(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                if(ClickTimer == null)
                {
                    ClickTimer = new Timer(300);                    
                    ClickTimer.Elapsed += new ElapsedEventHandler(singleClickTimer);                        
                }
                if(!ClickTimer.Enabled)
                {
                    //Clicked Once                    
                    ClickTimer.Start();
                    Console.WriteLine(e.ClickCount);
                    Console.WriteLine(e.RoutedEvent);
                    Console.WriteLine(e.OriginalSource);
                    
                                    
                    //NodeCheck(e.OriginalSource as DependencyObject);                                                                                          
                    //trvMouse_SingleClick(sender, e);
                    Console.WriteLine(e.ClickCount);
                }
                else 
                {
                    //Click Twice
                    ClickTimer.Stop();                   
                    Console.WriteLine("Double Click ->" + e.ClickCount);
                    trvMouse_DoubleClick(sender, e);                   
                }
                //IsHitTestVisible = true;                  
            }           
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void singleClickTimer(object sender, EventArgs e)
        {           
            ClickTimer.Stop();           
            Console.WriteLine("timer stopped");            
        }

        //private void previewMouse_DoubleClick(object sender, MouseButtonEventArgs e)
        //{
        //    e.Handled = true;
        //    Console.WriteLine("previewMouse_DoubleClick => " + e.OriginalSource);
        //    Console.WriteLine("previewMouse_DoubleClick => " + e.RoutedEvent);
        //    Console.WriteLine("previewMouse_DoubleClick => " + e.Source);
        //}

        //private void previewRightMouseButton_Click(object sender, MouseButtonEventArgs e)
        //{
        //    e.Handled = true;
        //    Console.WriteLine("previewRightMouseButton_Click => " + e.OriginalSource);
        //    Console.WriteLine("previewRightMouseButton_Click => " + e.RoutedEvent);
        //    Console.WriteLine("previewRightMouseButton_Click => " + e.Source);
        //}
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
        

        /// <summary>
        ///  
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        private static TreeViewItem NodeCheck(DependencyObject source)
        {
            while (source != null && !(source is TreeViewItem))
                source = System.Windows.Media.VisualTreeHelper.GetParent(source);
            TreeViewItem item = source as TreeViewItem;
                        
            Console.WriteLine("Source -> " + item.Parent.DependencyObjectType.Name);
            Console.WriteLine("Source -> " + item.ItemContainerGenerator.Items);
                                   
            for(int i = 0; i < item.ItemContainerGenerator.Items.Count; i++)
            {               
                Console.WriteLine("Item -> " + item.ItemContainerGenerator.Items[i].ToString());
            }           
            Console.WriteLine("Header -> " + item.Header);          
            Console.WriteLine("Tag -> " + item.Tag);
            return source as TreeViewItem; 
        }             
    }
}
