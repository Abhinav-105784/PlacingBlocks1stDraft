using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
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
using Autodesk.AutoCAD.ApplicationServices.Core;

namespace BlocksPlacingContinuousLoop
{
    public partial class SelectContinuousBlocks : Form
    {
        private double distance = 0;
        private double vertical_Distance = 0;
        private List<(double, double)> insertedPositions = new List<(double, double)>();
        private List<BlocksPlaced> blocks = new List<BlocksPlaced>();
        public SelectContinuousBlocks()
        {
            InitializeComponent();
            Next_Distance.Text = "0";
            Vertical_Distance.Text = "0";
            Up_Check.Checked = true;
            RightCheck.Checked = true;

        }

        private void Main_Block_TextChanged(object sender, EventArgs e)
        {

        }

        private void Left_Block_TextChanged(object sender, EventArgs e)
        {

        }

        private void Right_Block_TextChanged(object sender, EventArgs e)
        {

        }

        private void Browse_Main_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fileSelect = new OpenFileDialog())
            {
                fileSelect.Filter = "Drawing Files | *.dwg";
                fileSelect.Title = "Select the Block Drawing";

                if (fileSelect.ShowDialog() == DialogResult.OK)
                {
                    Main_Block.Text = fileSelect.FileName;

                }
            }
        }

        private void Browse_Left_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fileSelect = new OpenFileDialog())
            {
                fileSelect.Filter = "Drawing Files | *.dwg";
                fileSelect.Title = "Select the Block Drawing";

                if (fileSelect.ShowDialog() == DialogResult.OK)
                {
                    Left_Block.Text = fileSelect.FileName;

                }
            }

        }

        private void Browse_right_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fileSelect = new OpenFileDialog())
            {
                fileSelect.Filter = "Drawing Files | *.dwg";
                fileSelect.Title = "Select the Block Drawing";

                if (fileSelect.ShowDialog() == DialogResult.OK)
                {
                    Right_Block.Text = fileSelect.FileName;

                }
            }

        }

        private void Open_Click(object sender, EventArgs e)
        {
            if (!double.TryParse(Next_Distance.Text, out distance))
            {
                MessageBox.Show("Please enter a valid distance.");
                return;
            }
            if (!double.TryParse(Vertical_Distance.Text, out vertical_Distance))
            {
                MessageBox.Show("Please enter a valid distance");
                return;
            }
            List<string> files = new List<string>();
            files.Add(Main_Block.Text); files.Add(Left_Block.Text); files.Add(Right_Block.Text);
            if (string.IsNullOrEmpty(Main_Block.Text))
            {
                MessageBox.Show("No Main block given");
                return;
            }
            if (insertedPositions.Count == 0)
            {
                MessageBox.Show("No initial position set.");
                return;
            }

            double currentX = insertedPositions.Last().Item1;
            double currentY = insertedPositions.Last().Item2;
            if (LeftCheck.Checked)
            {
                currentX -= distance;
                RightCheck.Checked = false;
            }
            else if (RightCheck.Checked)
            {
                currentX += distance;
                LeftCheck.Checked = false;
            }
            if (Up_Check.Checked)
            {
                currentY += vertical_Distance;
                Down_Check.Checked = false;
            }
            else if (Down_Check.Checked)
            {
                currentY -= vertical_Distance;
                Up_Check.Checked = false;
            }
            insertedPositions.Add((currentX, currentY));
            try
            {
                InsertBlocks.Placing(files, insertedPositions.Last().Item1, insertedPositions.Last().Item2);
                MessageBox.Show($"All blocks have been inserted successfully at x: {currentX} and Y: {currentY}.");
                BlocksPlaced bp = new BlocksPlaced
                {
                    MainBlock = Path.GetFileNameWithoutExtension(files[0]),
                    LeftBlock = Path.GetFileNameWithoutExtension(files[1]),
                    RightBlock = Path.GetFileNameWithoutExtension(files[2]),
                    X = currentX,
                    Y = currentY
                };
                blocks.Add(bp);
                //trying for dimensions
                var pair = insertedPositions[0];
                double minY = pair.Item2;
                for (int i =1; i< insertedPositions.Count;i++)
                {
                    pair = insertedPositions[i];
                    if (minY > pair.Item2)
                    {
                        minY = pair.Item2;
                    }
                }
                double lastX = currentX > 0 ? currentX - distance : currentX + distance;
                using (Transaction tr = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    using (Dimension dim = new AlignedDimension(new Autodesk.AutoCAD.Geometry.Point3d(lastX, minY, 0),
                        new Autodesk.AutoCAD.Geometry.Point3d(currentX, minY, 0),
                        new Autodesk.AutoCAD.Geometry.Point3d(currentX, minY, 0),
                        "",
                        ObjectId.Null))
                    {
                        dim.DimensionStyle = default;
                        dim.Annotative = AnnotativeStates.True;
                        btr.AppendEntity(dim);
                        tr.AddNewlyCreatedDBObject(dim, true);
                    }
                    tr.Commit();
                }
                //exporting excel.
                ExcelMapper export = new ExcelMapper();
                var newFile = $@"C:\Users\{Environment.UserName}\Downloads\BlocksPlaced.xlsx";
                export.Save(newFile, blocks, "Blocks", true);

                Close();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"{ex.Message}");
            }


        }
        public void FirstBp(BlocksPlaced bp)
        {
            blocks.Insert(0,bp);

        }
        public void SetInitialPosition(double x, double y)
        {
            insertedPositions.Add((x, y));
        }

        private void Next_Click(object sender, EventArgs e)
        {
            if (!double.TryParse(Next_Distance.Text, out distance))
            {
                MessageBox.Show("Please enter a valid distance.");
                return;
            }

            if (!double.TryParse(Vertical_Distance.Text, out vertical_Distance))
            {
                MessageBox.Show("Please enter a valid distance");
                return;
            }
            List<string> files = new List<string>();
            files.Add(Main_Block.Text); files.Add(Left_Block.Text); files.Add(Right_Block.Text);
            if (insertedPositions.Count == 0)
            {
                MessageBox.Show("No initial position set.");
                return;
            }
            double currentX = insertedPositions.Last().Item1;
            double currentY = insertedPositions.Last().Item2;
            if (LeftCheck.Checked)
            {
                currentX -= distance;
                RightCheck.Checked = false;
            }
            else if (RightCheck.Checked)
            {
                currentX += distance;
                LeftCheck.Checked = false;
            }
            if (Up_Check.Checked)
            {
                currentY += vertical_Distance;
                Down_Check.Checked = false;
            }
            else if (Down_Check.Checked)
            {
                currentY -= vertical_Distance;
                Up_Check.Checked = false;
            }
            insertedPositions.Add((currentX, currentY));
            try
            {
                BlocksPlaced bp = new BlocksPlaced
                {
                    MainBlock = Path.GetFileNameWithoutExtension(files[0]),
                    LeftBlock = Path.GetFileNameWithoutExtension(files[1]),
                    RightBlock = Path.GetFileNameWithoutExtension(files[2]),
                    X = currentX,
                    Y = currentY
                };
                blocks.Add(bp);

                InsertBlocks.Placing(files, currentX, currentY);
                MessageBox.Show($"Block inserted at X: {currentX}, Y: {currentY}");

                //trying for dimensions
                var pair = insertedPositions[0];
                double minY = pair.Item2;
                for (int i = 1; i < insertedPositions.Count; i++)
                {
                    pair = insertedPositions[i];
                    if (minY > pair.Item2)
                    {
                        minY = pair.Item2;
                    }
                }
                double lastX = currentX > 0 ? currentX - distance : currentX + distance;
                using (Transaction tr = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    using (Dimension dim = new AlignedDimension(new Autodesk.AutoCAD.Geometry.Point3d(lastX, minY, 0),
                        new Autodesk.AutoCAD.Geometry.Point3d(currentX, minY, 0),
                        new Autodesk.AutoCAD.Geometry.Point3d(currentX, minY, 0),
                        "",
                        ObjectId.Null))
                    {
                        dim.DimensionStyle = default;
                        dim.Annotative = AnnotativeStates.True;
                        btr.AppendEntity(dim);
                        tr.AddNewlyCreatedDBObject(dim, true);
                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"{ex.Message}");
            }
            Main_Block.Clear();
            Left_Block.Clear();
            Right_Block.Clear();
        }

        private void Next_Distance_TextChanged(object sender, EventArgs e)
        {

        }

        private void Close_Click(object sender, EventArgs e)
        {
            ExcelMapper export = new ExcelMapper();
            var newFile = $@"C:\Users\{Environment.UserName}\Downloads\BlocksPlaced.xlsx";
            export.Save(newFile, blocks, "Blocks", true);
            Close();
        }

        private void LeftCheck_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void RightCheck_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void Vertical_Distance_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
