using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MahApps.Metro.Controls;
using Word_To_Markup_Converter.Module;
using Microsoft.Win32;
using System.IO;

namespace Word_To_Markup_Converter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        private MarkupGenerator generator;
        OpenFileDialog openOriginalFile;
        SaveFileDialog saveNewFile;
        public MainWindow()
        {
            InitializeComponent();
            openOriginalFile = new OpenFileDialog();
            saveNewFile = new SaveFileDialog();
            openOriginalFile.FileOk += new System.ComponentModel.CancelEventHandler(openOriginalFile_FileOk);
            saveNewFile.FileOk += new System.ComponentModel.CancelEventHandler(saveNewFile_FileOk);
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

        }

        private void btnGenerateMarkup_Click(object sender, RoutedEventArgs e)
        {
            if (txtDocumentName.Text != String.Empty)
            {
                StreamWriter writer = new StreamWriter(txtSavePath.Text);
                if ((bool)rbtnHTML.IsChecked)
                {
                    generator = new HTMLGenerator();                    
                    writer.Write(((HTMLGenerator)generator).generateMarkup(txtDocumentName.Text, txtHeaderTextPath.Text, txtFooterTextPath.Text, txtDocumentTitle.Text));                 
                    System.Diagnostics.Process.Start(txtSavePath.Text);                    
                }
                else if ((bool)rbtnMarkDown.IsChecked)
                {
                    generator = new MarkdownGenerator();
                    writer.Write(generator.generateMarkup(txtDocumentName.Text));
                }
                writer.Close();
                MessageBox.Show("File Generation Complete");
            }
        }

        private void txtDocumentName_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {

        }

        private void btnOpenOriginalFile_Click(object sender, RoutedEventArgs e)
        {            
            openOriginalFile.Filter ="Word Documents (*.doc, *docx)|*.docx;*.doc";
            openOriginalFile.ShowDialog();
        }

        

        private void btnOpenSaveFile_Click(object sender, RoutedEventArgs e)
        {
            if ((bool)rbtnHTML.IsChecked)
            {
                saveNewFile.Filter = "HTML Document (*.html, *.htm)|*.html;*htm";
            }
            else if ((bool)rbtnMarkDown.IsChecked)
            {
                saveNewFile.Filter = "HTML Document (*.markdown, *.md)|*.markdown;*.md";
            }
            saveNewFile.ShowDialog();
        }

        private void btnOpenFooter_Click(object sender, RoutedEventArgs e)
        {

        }

        void openOriginalFile_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            txtDocumentName.Text = openOriginalFile.FileName;
        }

        void saveNewFile_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            txtSavePath.Text = saveNewFile.FileName;
        }
    }
}
