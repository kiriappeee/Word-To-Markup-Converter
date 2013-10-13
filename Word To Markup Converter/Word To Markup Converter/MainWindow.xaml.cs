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
using System.Text.RegularExpressions;

namespace Word_To_Markup_Converter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        private MarkupGenerator generator;
        OpenFileDialog openFile;
        OpenFileDialog openHeaderFooter;
        SaveFileDialog saveNewFile;
        private string openFileParent;

        public MainWindow()
        {
            InitializeComponent();
            openFile = new OpenFileDialog();
            saveNewFile = new SaveFileDialog();
            openFile.FileOk += new System.ComponentModel.CancelEventHandler(openOriginalFile_FileOk);
            saveNewFile.FileOk += new System.ComponentModel.CancelEventHandler(saveNewFile_FileOk);
        }

        private void btnGenerateMarkup_Click(object sender, RoutedEventArgs e)
        {
            if (txtDocumentName.Text != String.Empty)
            {
                StreamWriter writer = new StreamWriter(txtSavePath.Text);
                if ((bool)rbtnHTML.IsChecked)
                {
                    generator = new HTMLGenerator();
                    if (txtHeaderTextPath.Text != "" && txtFooterTextPath.Text != "")
                    {
                        ((HTMLGenerator)generator).generateMarkup(txtDocumentName.Text, txtHeaderTextPath.Text, txtFooterTextPath.Text, txtDocumentTitle.Text);
                    }
                    else
                    {

                        ((HTMLGenerator)generator).generateMarkup(txtDocumentName.Text, txtDocumentTitle.Text);

                    }
                     
                    System.Diagnostics.Process.Start(txtSavePath.Text);                    
                }
                else if ((bool)rbtnMarkDown.IsChecked)
                {
                    generator = new MarkdownGenerator();
                    generator.generateMarkup(txtDocumentName.Text);                    
                }
                writer.Write(generator.docText.ToString());
                writer.Close();
                MessageBox.Show("File Generation Complete");
            }
        }

        private void txtDocumentName_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {

        }

        private void btnOpenOriginalFile_Click(object sender, RoutedEventArgs e)
        {
            openFileParent = "ORIGINAL";
            openFile.Filter ="Word Documents (*.doc, *docx)|*.docx;*.doc";
            openFile.ShowDialog();
        }

        private void btnOpenFooter_Click(object sender, RoutedEventArgs e)
        {
            openFileParent = "FOOTER";
            openFile.Filter = "HTML Documents (*.html, *.htm)|*.html;*htm";
            openFile.ShowDialog();
        }

        private void btnOpenHeader_Click(object sender, RoutedEventArgs e)
        {
            openFileParent = "HEADER";
            openFile.Filter = "HTML Documents (*.html, *.htm)|*.html;*htm";
            openFile.ShowDialog();
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

        void openOriginalFile_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            switch (openFileParent)
            {
                case "ORIGINAL":
                    txtDocumentName.Text = openFile.FileName;    
                    break;
                case "HEADER":
                    txtHeaderTextPath.Text = openFile.FileName;
                    break;
                case "FOOTER":
                    txtFooterTextPath.Text = openFile.FileName;
                    break;
            }
            
        }

        void saveNewFile_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            txtSavePath.Text = saveNewFile.FileName;
        }
    }
}
