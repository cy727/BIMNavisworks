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
using System.Windows.Shapes;

namespace BIADBIMnavisworks
{
    /// <summary>
    /// WindowPDF.xaml 的交互逻辑
    /// </summary>
    public partial class WindowPDF : Window
    {
        public WindowPDF()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            viewPDF.LoadPDFfile("E:\\chris\\Revit\\WPF\\Drawings\\B-B01.pdf");
        }
    }
}
