using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;

namespace NacCleaner {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {

        public static string LIFE = "NACOLAH Life";
        public static string ANNUITY = "NACOLAH Annuity";
        public static string filePath = "";
        public static string fileName = "";

        public MainWindow() {
            InitializeComponent();
            cbType.Items.Add(LIFE);
            cbType.Items.Add(ANNUITY);
            btnLoad.IsEnabled = false;
            btnGo.IsEnabled = false;
        }

        private void cbType_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            btnLoad.IsEnabled = true;
            lblFile.Content = "Choose file";
            lblStatus.Content = "...";
        }

        private void btnLoad_On_Click(object sender, RoutedEventArgs e) {
            System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog();
            ofd.InitialDirectory = "P:\\RALFG\\Common Files\\Commissions & Insurance\\Commission Statements\\2017\\";
            ofd.Filter = "PDF files (*.pdf)|*.pdf";
            ofd.FilterIndex = 1;
            ofd.RestoreDirectory = true;
            System.Windows.Forms.DialogResult result = ofd.ShowDialog();
            lblStatus.Content = "...";

            if (result == System.Windows.Forms.DialogResult.OK) {
                filePath = ofd.FileName;
                fileName = System.IO.Path.GetFileName(filePath);
                lblFile.Content = filePath;
                btnGo.IsEnabled = true;
            } else {
                lblFile.Content = "Please choose a file";
                btnGo.IsEnabled = false;
            }
        }

        private void btnGo_Click(object sender, RoutedEventArgs e) {
            if(fileName == "" || fileName == null){
                btnGo.IsEnabled = false;
                lblFile.Content = "Please choose a file";
            } else {
                if(cbType.SelectedValue.ToString() == LIFE) {
                    lblStatus.Content = "Starting NACOLAH Life clean";
                    new NacLife(filePath);
                    lblStatus.Content = lblStatus.Content + "...Done";
                } else {
                    lblStatus.Content = "Starting NACOLAH Annuity clean";
                    new NacAnn(filePath);
                    lblStatus.Content = lblStatus.Content + "...Done";
                }
                btnLoad.IsEnabled = true;
                lblFile.Content = "...";
                btnGo.IsEnabled = false;
                cbType.SelectedIndex = 0;
            }
        }
    }
}
