using System;
using IO = System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using SW = System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;
using System.ComponentModel;

namespace PowerPointGenerator
{
    /// <summary>
    /// 
    /// Create/set the SourceFolder for PowerPoint file processing
    ///  ------ 
    /// Copy source images to the SourceFolder; each trust in it's own folder named with the trust ID
    /// Set a destination folder(Done for example)
    /// Requires the following support files.
    /// Base file:	list.txt <- list of names to be added
    /// PowerPoint template file:	MBRRACE02.potx
    

    /// </summary>
    public partial class MainWindow : Window
    {
        private string NameListFile;
        // for another day private BackgroundWorker backgroundWorker = new BackgroundWorker();

        public MainWindow()
        {
            InitializeComponent();

            InitialiseApplication();

        }

        private void InitialiseApplication()
        {
            ProcessedCount.IsReadOnly = true;
            ProcessInformation.IsReadOnly = true;

            NameListFile = Properties.Settings.Default.NameList;

            DestinationFolder.Text = Properties.Settings.Default.DestinationFolder;
            BaseSettingsFolder.Text = Properties.Settings.Default.BaseSettingsFolder;
        }


        #region "Screen controls"
        private void ExitForm_Click(object sender, RoutedEventArgs e)
        {
            SW.Application.Current.Shutdown();
        }

        private void SelectDestinationFolder_Click(object sender, RoutedEventArgs e)
        {
            DestinationFolder.Text = GetFolder();
        }

        private void SelectBaseSettings_Click(object sender, RoutedEventArgs e)
        {
            BaseSettingsFolder.Text = GetFolder();
        }

        private void StartFileProcess_Click(object sender, RoutedEventArgs e)
        {
            //CreatePPT(IO.Path.Combine(DestinationFolder.Text.Trim(), "test.pptx"));
            //  check working folders are set and existing

            if (!CheckFolderExists(DestinationFolder.Text.Trim()))
            {
                SW.MessageBox.Show("Please add a valid destination folder");
                return;
            }

            if (!CheckFolderExists(BaseSettingsFolder.Text.Trim()))
            {
                SW.MessageBox.Show("Please add a valid settings folder");
                return;
            }

            //  save folder settings
            Properties.Settings.Default.DestinationFolder = DestinationFolder.Text.Trim();
            Properties.Settings.Default.BaseSettingsFolder = BaseSettingsFolder.Text.Trim();

            //  set base presentation 

            PowerPointReports powerPointReports = new PowerPointReports();

            powerPointReports.TemplateFile = Properties.Settings.Default.Template;
            powerPointReports.NameListFile = NameListFile;
            powerPointReports.BaseSettingsFolder = BaseSettingsFolder.Text.Trim();

            powerPointReports.DestinationFolder = DestinationFolder.Text.Trim();
            powerPointReports.ProcessTextBox = ProcessInformation;
            //powerPointReports.ProcessedCount.Text = powerPointReports.CreatePresentations();

            powerPointReports.CreatePresentations();
            

            Properties.Settings.Default.Save();

            //powerPointReports.ClearDestinationFolder();

            powerPointReports.GetBaseCounts();

            //  clear up
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

        }

        #endregion        

        #region "Helper methods"

        private static string GetFolder()
        {
            var dialog = new FolderBrowserDialog();
            var result = dialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                return dialog.SelectedPath;
            }

            return string.Empty;
        }

        private static bool CheckFolderExists(string folderToCheck)
        {
            if (folderToCheck == string.Empty)
            {
                return false;
            }

            if (!IO.Directory.Exists(folderToCheck))
            {
                return false;
            }

            return true;
        }        
    }
    #endregion


}

