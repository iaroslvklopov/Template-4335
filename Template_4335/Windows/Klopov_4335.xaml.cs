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

namespace Template_4335.Windows
{
    /// <summary>
    /// Логика взаимодействия для Klopov_4335.xaml
    /// </summary>
    public partial class Klopov_4335 : Window
    {
        public Klopov_4335()
        {
            InitializeComponent();
        }
        private void ExcelPageBtn_Click(object sender, RoutedEventArgs e)
        {
            MainFrame.Navigate(new Klopov4335.ExcelPage());
        }

        private void WordPageBtn_Click(object sender, RoutedEventArgs e)
        {
            //MainFrame.Navigate(new Klopov4335.WordPage());
        }
    }
}
