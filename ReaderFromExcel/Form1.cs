using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReaderFromExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void browser_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fileDialog=new OpenFileDialog())
            {
                fileDialog.RestoreDirectory = true;
                fileDialog.Filter = "Excel files(*.xlsx;*.xls)|*.xlsx;*.xls";


                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    excelName.Text = fileDialog.FileName;
                }

                ExcelWorkBookReader reader = new ExcelWorkBookReader();
                reader.ReadExcelWorkBook(excelName.Text, reader.NamesOfSheets(excelName.Text));
            }
        }

        //private void convertXML_Click(object sender, EventArgs e)
        //{
        //    //txbXmlView.Clear();
        //    string excelfileName = excelName.Text;


        //    if (string.IsNullOrEmpty(excelfileName) || !File.Exists(excelfileName))
        //    {
        //        MessageBox.Show("The Excel file is invalid! Please select a valid file.");
        //        return;
        //    }


        //    try
        //    {
        //        string xmlFormatstring = new ConvertToXml().GetXML(excelfileName);
        //        if (string.IsNullOrEmpty(xmlFormatstring))
        //        {
        //            MessageBox.Show("The content of Excel file is Empty!");
        //            return;
        //        }


        //        xmlView.Text = xmlFormatstring;


        //        // If txbXmlView has text, set btnSaveAs button to be enable 
        //        saveXML.Enabled = true;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("There is error occurs! The error message is: " + ex.Message);
        //    }
        //}

        private void saveXML_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog savefiledialog = new SaveFileDialog())
            {
                savefiledialog.RestoreDirectory = true;
                savefiledialog.DefaultExt = "xml";
                savefiledialog.Filter = "All Files(*.xml)|*.xml";
                if (savefiledialog.ShowDialog() == DialogResult.OK)
                {
                    Stream filestream = savefiledialog.OpenFile();
                    StreamWriter streamwriter = new StreamWriter(filestream);
                    streamwriter.Write("<?xml version='1.0'?>" +
                        Environment.NewLine + xmlView.Text);
                    streamwriter.Close();
                }
            }
        }
    }
}
