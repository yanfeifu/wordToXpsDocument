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

using System.IO;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.Windows.Xps.Packaging;

namespace WordToXpsDocument
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Window1 : System.Windows.Window
    {
        public Window1()
        {
            InitializeComponent();

        }

        /// <summary>
        /// This method takes a Word document full path and new XPS document full path and name
        /// and returns the new XpsDocument
        /// </summary>
        /// <param name="wordDocName"></param>
        /// <param name="xpsDocName"></param>
        /// <returns></returns>
        private XpsDocument ConvertWordDocToXPSDoc(string wordDocName, string xpsDocName)
        {
            // Create a WordApplication and add Document to it
            Microsoft.Office.Interop.Word.Application
                wordApplication = new Microsoft.Office.Interop.Word.Application();
            wordApplication.Documents.Add(wordDocName);

           
            Document doc = wordApplication.ActiveDocument;
            // You must make sure you have Microsoft.Office.Interop.Word.Dll version 12.
            // Version 11 or previous versions do not have WdSaveFormat.wdFormatXPS option
            try
            {
                doc.SaveAs(xpsDocName, WdSaveFormat.wdFormatXPS);
                wordApplication.Quit();

                XpsDocument xpsDoc = new XpsDocument(xpsDocName, System.IO.FileAccess.Read);
                return xpsDoc;
            }
            catch (Exception exp)
            {
                string str = exp.Message;
            }
            return null;            
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension
            dlg.DefaultExt = ".doc";
            dlg.Filter = "Word documents (.doc)|*.doc|" + "All files (*.*)|*.*";

            // Display OpenFileDialog by calling ShowDialog method
            Nullable<bool> result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox
            if (result == true)            
            {
                if (dlg.FileName.Length > 0)
                {
                    SelectedFileTextBox.Text = dlg.FileName;
                    string newXPSDocumentName = String.Concat(System.IO.Path.GetDirectoryName(dlg.FileName), "\\",
                                   System.IO.Path.GetFileNameWithoutExtension(dlg.FileName), ".xps");

                    // Set DocumentViewer.Document to XPS document
                    documentViewer1.Document =
                        ConvertWordDocToXPSDoc(dlg.FileName, newXPSDocumentName).GetFixedDocumentSequence();
                }                
            }
        }        
    }
}
