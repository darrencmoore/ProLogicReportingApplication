﻿using System;
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
        public List<String> ProLogic_zContractContacts = new List<String>();
        
        
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

        #region TreeViewLoadMethod
        /// <summary>
        /// Tree View - loops over the list returned from Nucleus _agent SELECT statement
        /// Puts each item in ProLogic_zContractContacts into a TreeViewItem
        /// Heirarchy goes Account => ContactFullName -> Contact Email Address
        /// This gets called from MainWindow.xaml when the TreeView is intially loaded
        /// As there isnot a OnClick event for TreeView nor the TreeViewItems I had to use PreviewMouseLeftButtonDown
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
                    accountItem.Tag = "{Parent} " + ProLogic_zContractContacts[i].Replace("{ Header = Item Level 0 }", "");
                    accountItem.IsExpanded = true;
                    accountItem.Header = new CheckBox()
                    {
                        IsChecked = true,
                        IsEnabled = true,
                        Content = ProLogic_zContractContacts[i].Replace("{ Header = Item Level 0 }", "")
                    };
                    //tree.MouseLeftButtonDown += trvSingle_click;
                    //tree.MouseDoubleClick += trvDouble_click;                    
                    // Account Item Right Mouse Click Handler
                    //tree.PreviewMouseRightButtonDown += trvRight_MouseButtonDown;
                    tree.Items.Add(accountItem);
                }
                //ContactFullName = Level 1
                if (ProLogic_zContractContacts[i].Contains("{ Header = Item Level 1 }"))
                {                    
                    empItem = new TreeViewItem();
                    empItem.Tag = "{Child} " + ProLogic_zContractContacts[i].Replace("{ Header = Item Level 1 }", "");
                    empItem.Header = new CheckBox()
                    {
                        IsChecked = true,
                        Content = ProLogic_zContractContacts[i].Replace("{ Header = Item Level 1 }", "")
                    };
                    //tree.PreviewMouseLeftButtonDown += trvSingle_click;
                    //tree.MouseDoubleClick += trvDouble_click;
                    // Employee Item Right Mouse Click Handler
                    //tree.PreviewMouseRightButtonDown += trvRight_MouseButtonDown;
                    accountItem.Items.Add(empItem);
                }
                //Contact Email Address = Level 2
                if (ProLogic_zContractContacts[i].Contains("{ Header = Item Level 2 }"))
                {                    
                    empEmailAddrItem = new TreeViewItem();
                    empEmailAddrItem.Tag = "{Email} " + ProLogic_zContractContacts[i].Replace("{ Header = Item Level 2 }", "");
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
        #endregion

        #region Mouse Click Event Handlers        
        private void trvMouse_SingleClick(object sender, MouseButtonEventArgs e) //RoutedEventArgs
        {
            Console.WriteLine("trvSingle_click -> " + e.ClickCount);
            if (e.RoutedEvent == UIElement.PreviewMouseLeftButtonDownEvent && e.ClickCount != 2)
            {
                Console.WriteLine("trvSingle_click -> " + e.ClickCount);
                e.Handled = true;
            }
        }
        /// <summary>
        /// Handler for trvAccount_AccountContacts item click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void trvMouse_DoubleClick(object sender, MouseButtonEventArgs e) //RoutedEventArgs
        {
            Console.WriteLine("trvDouble_click -> " + e.ClickCount);
            if (e.RoutedEvent == Control.MouseDoubleClickEvent && e.ClickCount != 1)
            {
                if(e.Source is TreeViewItem && (e.Source as TreeViewItem).IsSelected)
                {
                    Console.WriteLine("trvDouble_click is selected -> " + e.ClickCount);
                    NodeCheck(e.OriginalSource as DependencyObject);
                    e.Handled = true;
                }
            }
        }       

        private void trvRightMouseButton_Click(object sender, MouseButtonEventArgs e) //RoutedEventArgs
        {
            Console.WriteLine("trvRight_MouseButtonDown -> " + e.ClickCount);
            e.Handled = true;
        }
        #endregion

        #region Preview Mouse Click Events
        private void previewMouse_DoubleClick(object sender, RoutedEventArgs e)
        {

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
