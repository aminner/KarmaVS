using System;
using System.Configuration;
using System.Windows;
using Microsoft.Win32;
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
            if ((int) Settings.Default.Properties[KarmaVsPackage.ProjectGuids].Attributes["ConfigType"] == (int) KarmaVsStaticClass.KarmaConfigType.Default)
            {
                Default.IsChecked = true;
                KarmaConfigFile.IsEnabled = false;
                selectButton.IsEnabled = false;
            }
            else
            {
                Custom.IsChecked = true;
                KarmaConfigFile.Text =
                    Settings.Default.Properties[KarmaVsPackage.ProjectGuids].Attributes["ConfigLocation"].ToString();
            }
        }

        private void selectButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog newOpenFileDialog = new OpenFileDialog();
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
                if (Settings.Default.Properties[KarmaVsPackage.ProjectGuids].Attributes["ConfigLocation"] == null)
                {
                    Settings.Default.Properties[KarmaVsPackage.ProjectGuids].Attributes.Add("ConfigLocation", KarmaConfigFile.Text);
                }
                else
                {
                    Settings.Default.Properties[KarmaVsPackage.ProjectGuids].Attributes["ConfigLocation"] = KarmaConfigFile.Text;
                }

                if (Settings.Default.Properties[KarmaVsPackage.ProjectGuids].Attributes["ConfigType"] == null)
                {
                    Settings.Default.Properties[KarmaVsPackage.ProjectGuids].Attributes.Add("ConfigType", KarmaVsStaticClass.KarmaConfigType.Custom);
                }
                else
                {
                    Settings.Default.Properties[KarmaVsPackage.ProjectGuids].Attributes["ConfigType"] = KarmaVsStaticClass.KarmaConfigType.Custom;
                }
            }
            else
            {
                if (Settings.Default.Properties[KarmaVsPackage.ProjectGuids].Attributes["ConfigLocation"] == null)
                {
                    Settings.Default.Properties[KarmaVsPackage.ProjectGuids].Attributes.Add("ConfigLocation", "");
                }
                else
                {
                    Settings.Default.Properties[KarmaVsPackage.ProjectGuids].Attributes["ConfigLocation"] = "";
                }

                if (Settings.Default.Properties[KarmaVsPackage.ProjectGuids].Attributes["ConfigType"] == null)
                {
                    Settings.Default.Properties[KarmaVsPackage.ProjectGuids].Attributes.Add("ConfigType", KarmaVsStaticClass.KarmaConfigType.Default);    
                }
                else
                {
                    Settings.Default.Properties[KarmaVsPackage.ProjectGuids].Attributes["ConfigType"] = KarmaVsStaticClass.KarmaConfigType.Default;    
                }
            }
            Settings.Default.Save();
            Close();
        }

        private void Cancel_OnClick(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
