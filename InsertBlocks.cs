using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Ganss.Excel;
using NPOI.POIFS.Storage;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
[assembly: CommandClass(typeof(BlocksPlacingContinuousLoop.InsertBlocks))]

namespace BlocksPlacingContinuousLoop
{
    public class InsertBlocks
    {
        //Command name
        [CommandMethod("AssembleBlocks")]

        //Method to load the forms through command
        public void LoadForm()
        {
            SelectBlocks form = new SelectBlocks();
            form.ShowDialog();
        }

        // Here we will take inputs of files as a list of string and x and y for the main blocks, the same x and y will be
        //used to add distances in next pair or triplet.
        public static void Placing(List<string> files, double xMain, double yMain)
        {
            Document document = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Database database = document.Database;
            Editor editor = document.Editor;
           

            try
            {
                List<Database> dbFiles = new List<Database>();
                List<ObjectId> blockIds = new List<ObjectId>();
                //transaction for main drawing
                using (Transaction transaction = database.TransactionManager.StartTransaction())
                {
                    //adding files as a database
                    for (int i = 0; i < files.Count; i++)
                    {
                        if (string.IsNullOrEmpty(files[i]))
                        {
                            dbFiles.Add(null);
                            continue;
                        }
                        Database dbFile = new Database(false, true);
                        dbFile.ReadDwgFile(files[i], FileOpenMode.OpenForReadAndAllShare, false, null);
                        dbFiles.Add(dbFile);
                    }
                    //using files database to get blocks objectids
                    for (int i = 0; i < dbFiles.Count; i++)
                    {
                        if (dbFiles[i] == null)
                        {
                            blockIds.Add(ObjectId.Null);
                            continue;
                        }
                        //transactions for block drawings using block ids to be added to current drawing
                        using (Transaction dbTransaction = dbFiles[i].TransactionManager.StartTransaction())
                        {
                            BlockTable btMain = (BlockTable)dbTransaction.GetObject(dbFiles[i].BlockTableId, OpenMode.ForRead);
                            ObjectId blockId = database.Insert(System.IO.Path.GetFileNameWithoutExtension(files[i]), dbFiles[i].Wblock(btMain[System.IO.Path.GetFileNameWithoutExtension(files[i])]), true);
                            blockIds.Add(blockId);
                            dbTransaction.Commit();
                        }
                    }                    

                    //opening blocktablerecord to append main and side blocks
                    BlockTableRecord currentSpace = transaction.GetObject(database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    //adding main block
                    BlockReference mainBr = new BlockReference(new Point3d(xMain, yMain, 0), blockIds[0])
                    {
                        ScaleFactors = new Scale3d(1, 1, 1)
                    };
                    currentSpace.AppendEntity(mainBr);
                    transaction.AddNewlyCreatedDBObject(mainBr, true);
                    BlockTableRecord btr = (BlockTableRecord)transaction.GetObject(blockIds[0], OpenMode.ForRead);
                    RXClass attDefClass = RXObject.GetClass(typeof(AttributeDefinition));
                  
                    foreach (ObjectId id in btr)
                    {
                        //reading main blocks attributes
                        if (id.ObjectClass == attDefClass)
                        {
                            AttributeDefinition attDef = (AttributeDefinition)transaction.GetObject(id, OpenMode.ForRead);
                            AttributeReference attRef = new AttributeReference();
                            attRef.SetAttributeFromBlock(attDef, mainBr.BlockTransform);
                            mainBr.AttributeCollection.AppendAttribute(attRef);
                            transaction.AddNewlyCreatedDBObject(attRef, true);
                        }
                    }
                    //x and y to insert left and right blocks
                    double insert_x1 = 0;
                    double insert_y1 = 0;
                    double insert_x2 = 0;
                    double insert_y2 = 0;

                    foreach (ObjectId attId in mainBr.AttributeCollection)
                    {
                        //using main block attributes to insert left and right clamp blocks
                        AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForRead);

                        //adding left right block
                        if (attRef.Tag == "PHASE2_FIXING_POSITION")
                        {
                            string[] coordinates = attRef.TextString.Split(',');
                            if (double.TryParse(coordinates[0], out insert_x1) &&
                              double.TryParse(coordinates[1], out insert_y1))
                            {
                                if (blockIds.Count > 2 && blockIds[2] != ObjectId.Null)
                                {
                                    BlockReference brRight = new BlockReference(new Point3d(xMain + insert_x1, yMain + insert_y1, 0), blockIds[2])
                                    {
                                        ScaleFactors = new Scale3d(1, 1, 1)

                                    };
                                    currentSpace.AppendEntity(brRight);
                                    transaction.AddNewlyCreatedDBObject(brRight, true);
                                    
                                }
                            }
                        }

                        //adding left block
                        if (attRef.Tag == "PHASE1_FIXING_POSITION")
                        {
                            string[] coordinates = attRef.TextString.Split(',');
                            if (double.TryParse(coordinates[0], out insert_x2) && double.TryParse(coordinates[1], out insert_y2))
                            {
                                if (blockIds.Count > 1 && blockIds[1] != ObjectId.Null)
                                {
                                    BlockReference brLeft = new BlockReference(new Point3d(xMain + insert_x2, yMain + insert_y2, 0), blockIds[1])
                                    {
                                        ScaleFactors = new Scale3d(-1, 1, 1)
                                    };
                                    currentSpace.AppendEntity(brLeft);
                                    transaction.AddNewlyCreatedDBObject(brLeft, true);
                                }
                            }
                        }

                    }
                    //commiting main drawing transitions.
                    transaction.Commit();
                }
                //ExcelMapper export = new ExcelMapper();
                //var newFile = @"C:\Users\goswamia0490\Downloads\BlocksPlaced.xlsx";
                //export.Save(newFile, blocks, "Blocks", true);

            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"{ex}");
            } //using try and catch just to bypass any major error that could cause fatal error in civil3d and close the application.
            //catch will give us the error.
        }
    }

    public class BlocksPlaced
    {
        public string LeftBlock { set; get; }
        public string MainBlock { set; get; }
        public string RightBlock { set; get; }
        public double X { set; get; }
        public double Y { set; get; }

        public BlocksPlaced()
        {

        }
    }
}
