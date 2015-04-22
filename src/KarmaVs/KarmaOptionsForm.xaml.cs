using System;
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
using Microsoft.Win32;
using System.Xaml;
using devcoach.Tools.Properties;

namespace devcoach.Tools
{
    /// <summary>
    /// Interaction logic for KarmaOptionsForm.xaml
    /// </summary>
    public partial class KarmaOptionsForm : Window
    {
        public KarmaOptionsForm()
        {
            InitializeComponent();
            SetInitialConfig();
        }

        private void SetInitialConfig()
        {
            if (Settings.Default.KarmaConfigType == (int)KarmaVsStaticClass.KarmaConfigType.Default)
            {
                Default.IsChecked = true;
            }
            else
            {
                Custom.IsChecked = true;
                KarmaConfigFile.Text = Settings.Default.KarmaConfigLocation;
            }
        }

        private void selectButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog newOpenFileDialog = new OpenFileDialog();
            Nullable<bool> result = newOpenFileDialog.ShowDialog();
            if (result==true)
            {
                KarmaConfigFile.Text = newOpenFileDialog.FileName.ToString();
            }
        }

        private void Default_Checked(object sender, RoutedEventArgs e)
        {
            if (Default == null || Custom==null) return;
            Custom.IsChecked = !Default.IsChecked;
            KarmaConfigFile.IsEnabled = (bool)Custom.IsChecked;
            selectButton.IsEnabled = (bool)Custom.IsChecked;
        }

        private void Custom_Checked(object sender, RoutedEventArgs e)
        {
            if (Default == null || Custom == null) return;
            Custom.IsChecked = !Default.IsChecked;
            KarmaConfigFile.IsEnabled = (bool)Custom.IsChecked;
            selectButton.IsEnabled = (bool) Custom.IsChecked;
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            if (Custom.IsChecked==true)
            {
                Settings.Default.KarmaConfigType = (int)KarmaVsStaticClass.KarmaConfigType.Custom;
                Settings.Default.KarmaConfigLocation = KarmaConfigFile.Text;
                Properties.Settings.Default.Save();
                this.Close();
            }
            else
            {
                Settings.Default.KarmaConfigType = (int)KarmaVsStaticClass.KarmaConfigType.Default;
                Settings.Default.KarmaConfigLocation = "";
                Properties.Settings.Default.Save();
            }
        }

        private void Cancel_OnClick(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
