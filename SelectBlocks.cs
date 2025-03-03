using Ganss.Excel;
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


//This is the form for the initial pair or triplet.
namespace BlocksPlacingContinuousLoop
{
    public partial class SelectBlocks : Form
    {
        private List<BlocksPlaced> blocks = new List<BlocksPlaced>();
        private double distance = 0;
       
        //Initializer of form x and y initially set to 0.
        public SelectBlocks()
        {
            InitializeComponent();
            X.Text = "0";
            Y.Text = "0";
        }

        public void SelectMain_Load(object sender, EventArgs e)
        {

        }


        public void MainBlock_TextChanged(object sender, EventArgs e)
        {

        }

        public void X_TextChanged(object sender, EventArgs e)
        {

        }

        public void Y_TextChanged(object sender, EventArgs e)
        {

        }

        public void LeftBlock_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        public void RightBlock_TextChanged(object sender, EventArgs e)
        {

        }
        
        //main block browsing will populate the mainblock text box 
        public void Browse_Main_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fileSelect = new OpenFileDialog())
            {
                fileSelect.Filter = "Drawing Files | *.dwg";
                fileSelect.Title = "Select the Block Drawing";

                if (fileSelect.ShowDialog() == DialogResult.OK)
                {
                    MainBlock.Text = fileSelect.FileName;

                }
            }


        }

        //left block browsing will populate the leftblock text box 
        public void Browse_Left_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fileSelect = new OpenFileDialog())
            {
                fileSelect.Filter = "Drawing Files | *.dwg";
                fileSelect.Title = "Select the Block Drawing";

                if (fileSelect.ShowDialog() == DialogResult.OK)
                {
                    LeftBlock.Text = fileSelect.FileName;

                }
            }

        }

        //right block browsing will populate the rightblock text box 
        public void Browse_Right_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fileSelect = new OpenFileDialog())
            {
                fileSelect.Filter = "Drawing Files | *.dwg";
                fileSelect.Title = "Select the Block Drawing";

                if (fileSelect.ShowDialog() == DialogResult.OK)
                {
                    RightBlock.Text = fileSelect.FileName;

                }
            }

        }

        //open click (text = run) will simply call the Insertblocks class's placing method with all user inputs to insert the blocks
        public void Open_Click(object sender, EventArgs e)
        {
            List<string> files = new List<string>();
            files.Add(MainBlock.Text); files.Add(LeftBlock.Text); files.Add(RightBlock.Text);
            //main block can't be empty
            if(string.IsNullOrEmpty(MainBlock.Text))
            {
                MessageBox.Show("There is no main block given!");
                return;
            }

            //if (files.Any(file => string.IsNullOrEmpty(file)))
            //{
            //    MessageBox.Show("Please select the drawing files");
            //}
            if (!double.TryParse(X.Text, out double xMain))

            {
                MessageBox.Show("Please Enter a valid x Value");
            }
            if (!double.TryParse(Y.Text, out double yMain))
            {
                MessageBox.Show("Please Enter a valid y Value");
            }
            try
            {  
                BlocksPlaced bp = new BlocksPlaced
                {
                    MainBlock = Path.GetFileNameWithoutExtension(files[0]),
                    LeftBlock = Path.GetFileNameWithoutExtension(files[1]),
                    RightBlock = Path.GetFileNameWithoutExtension(files[2]),
                    X = xMain,
                    Y = yMain
                };
                blocks.Add(bp);
                ExcelMapper export = new ExcelMapper();
                var newFile = $@"C:\Users\{Environment.UserName}\Downloads\BlocksPlaced.xlsx";
                export.Save(newFile, blocks, "Blocks", true);
                InsertBlocks.Placing(files, xMain, yMain);
                MessageBox.Show("All Blocks have been successfully placed.");
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }

        }


        //next will load the next form - selectcontinuousblocks
        public void Next_Click(object sender, EventArgs e)
        {
            List<string> files = new List<string>();
            files.Add(MainBlock.Text); files.Add(LeftBlock.Text); files.Add(RightBlock.Text);
            if (string.IsNullOrEmpty(MainBlock.Text))
            {
                MessageBox.Show("There is no main block given!");
                return;
            }

            //if (files.Any(file => string.IsNullOrEmpty(file)))
            //{
            //    MessageBox.Show("Please select the drawing files");
            //}
            if (!double.TryParse(X.Text, out double xMain))

            {
                MessageBox.Show("Please Enter a valid x Value");
            }
            if (!double.TryParse(Y.Text, out double yMain))
            {
                MessageBox.Show("Please Enter a valid y Value");
            }
            try
            {
               
               
                InsertBlocks.Placing(files, xMain, yMain);
                MessageBox.Show("All Blocks have been successfully placed.");
              
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            Hide();
            Close();

            SelectContinuousBlocks continuousBlocksForm = new SelectContinuousBlocks();
            BlocksPlaced bp = new BlocksPlaced
            {
                MainBlock = Path.GetFileNameWithoutExtension(files[0]),
                LeftBlock = Path.GetFileNameWithoutExtension(files[1]),
                RightBlock = Path.GetFileNameWithoutExtension(files[2]),
                X = xMain,
                Y = yMain
            };
            //initial position of x and y of main are kept as xmain and ymain here and horizontal and
            //vertical distances will be added to them.
            continuousBlocksForm.SetInitialPosition(xMain, yMain);
            continuousBlocksForm.FirstBp(bp);
            continuousBlocksForm.ShowDialog();
        }

        //click close to close the form and terminate the plugin.
        public void Close_Click(object sender, EventArgs e)
        {
            Close();

        }

    }
}