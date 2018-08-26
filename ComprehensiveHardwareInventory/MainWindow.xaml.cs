﻿using System;
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

namespace ComprehensiveHardwareInventory
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region
        private List<RowObject> ItemList;
        #endregion
        public MainWindow()
        {
            InitializeComponent();
            ItemList = new List<RowObject>();
            for(int i = 0; i < 1000; i++)
            {
                ItemList.Add(new RowObject());
            }

            RowTable.ItemsSource = ItemList;
        }
    }

    public class RowObject
    {
        public string IOIndex;
        public string NameFunction;
        public string Anonym;
        public string Logic;
    }
}
