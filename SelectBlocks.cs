using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Schema;

namespace BlocksInsertionAndPositioning
{
    public partial class SelectBlocks : Form
    {
        public SelectBlocks()
        {
            InitializeComponent();
            xmain.Text = "0";
            ymain.Text = "0";
        }
        
        private void mainBlock_TextChanged(object sender, EventArgs e)
        {

        }
        

        private void Browse_Click(object sender, EventArgs e)
        {
            using(OpenFileDialog fileSelect = new OpenFileDialog())
            {
                fileSelect.Filter = "Drawing Files | *.dwg";
                fileSelect.Title = "Select the block Drawing";

                if(fileSelect.ShowDialog()==DialogResult.OK)
                {
                    
                    mainBlock.Text = fileSelect.FileName;
                }

            }
        }

        private void Block2_TextChanged(object sender, EventArgs e)
        {

        }

        private void Browse2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fileSelect = new OpenFileDialog())
            {
                fileSelect.Filter = "Drawing Files | *.dwg";
                fileSelect.Title = "Select the block Drawing";

                if (fileSelect.ShowDialog() == DialogResult.OK)
                {
                    Block2.Text = fileSelect.FileName;
                }

            }
        }

        private void Open_Click(object sender, EventArgs e)
        {
            string file1 = mainBlock.Text;
            string file2 = Block2.Text;

            if (string.IsNullOrEmpty(file1) || string.IsNullOrEmpty(file2))
            {
                MessageBox.Show("Please select the drawing files");
            }
            if(!double.TryParse(xmain.Text, out double xMain))

            {
                MessageBox.Show("Please Enter a valid x Value");
            }
            if(!double.TryParse(ymain.Text, out double yMain))
            {
                MessageBox.Show("Please Enter a valid x Value");
            }
            try
            {
                InsertionBlocks.Placing(file1, file2, xMain, yMain);
                MessageBox.Show("Both drawings have been successfully placed.");
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void Close_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void xmain_TextChanged(object sender, EventArgs e)
        {

        }

        private void ymain_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
