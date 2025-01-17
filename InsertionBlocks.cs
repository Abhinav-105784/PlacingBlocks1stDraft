using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
[assembly: CommandClass(typeof(BlocksInsertionAndPositioning.InsertionBlocks))]
namespace BlocksInsertionAndPositioning
{
    public class InsertionBlocks
    {
        [CommandMethod("BlocksPlaced")]
        public void LoadForm()
        {
            SelectBlocks st = new SelectBlocks();
            st.ShowDialog();
        }
        public static void Placing(string file1, string file2, double xMain, double yMain)
        {
            var document = Application.DocumentManager.MdiActiveDocument;
            Database database = document.Database;
            Editor editor = document.Editor;

            try
            {
                Database mainDb = new Database(false, true);
                Database secondDb = new Database(false, true);
                ObjectId MainBrId;
                ObjectId secondBrId;
                using (Transaction tr = database.TransactionManager.StartTransaction())
                {
                    try
                    {
                        mainDb.ReadDwgFile(file1, FileOpenMode.OpenForReadAndAllShare, false, null);
                        using (Transaction trMain = mainDb.TransactionManager.StartTransaction())
                        {
                            BlockTable btMain = (BlockTable)trMain.GetObject(mainDb.BlockTableId, OpenMode.ForRead);
                            MainBrId = database.Insert(Path.GetFileNameWithoutExtension(file1), mainDb.Wblock(btMain[Path.GetFileNameWithoutExtension(file1)]), true);
                            trMain.Commit();
                        }
                        editor.WriteMessage($"We added {Path.GetFileNameWithoutExtension(file1)}");// to check if block is actually inserted
                        BlockTableRecord currentSpace = tr.GetObject(database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        BlockReference mainBr = new BlockReference(new Point3d(xMain, yMain, 0), MainBrId)
                        {
                            ScaleFactors = new Scale3d(1, 1, 1)
                        };
                        currentSpace.AppendEntity(mainBr);
                        tr.AddNewlyCreatedDBObject(mainBr, true);
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(MainBrId, OpenMode.ForRead);
                        RXClass attDefClass = RXObject.GetClass(typeof(AttributeDefinition));

                        foreach (ObjectId id in btr)
                        {
                            if (id.ObjectClass == attDefClass)
                            {
                                AttributeDefinition attDef = (AttributeDefinition)tr.GetObject(id, OpenMode.ForRead);
                                AttributeReference attRef = new AttributeReference();
                                attRef.SetAttributeFromBlock(attDef, mainBr.BlockTransform);
                                mainBr.AttributeCollection.AppendAttribute(attRef);
                                tr.AddNewlyCreatedDBObject(attRef, true);
                            }
                        }
                        secondDb.ReadDwgFile(file2, FileOpenMode.OpenForReadAndAllShare, false, null);
                        using (Transaction trSecond = secondDb.TransactionManager.StartTransaction())
                        {
                            BlockTable btSecond = (BlockTable)trSecond.GetObject(secondDb.BlockTableId, OpenMode.ForRead);
                            secondBrId = database.Insert(Path.GetFileNameWithoutExtension(file2), secondDb.Wblock(btSecond[Path.GetFileNameWithoutExtension(file2)]), true);
                            trSecond.Commit();
                        }
                        double insert_x1 = 0;
                        double insert_y1 = 0;
                        double insert_x2 = 0;
                        double insert_y2 = 0;

                        foreach (ObjectId attId in mainBr.AttributeCollection)
                        {
                            AttributeReference attRef = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                            if (attRef.Tag == "PHASE2_FIXING_POSITION")
                            {
                                string[] coordinates = attRef.TextString.Split(',');
                                if (double.TryParse(coordinates[0], out insert_x1) &&
                                  double.TryParse(coordinates[1], out insert_y1))
                                {
                                    BlockReference secondBr1 = new BlockReference(new Point3d(xMain - (-insert_x1), yMain + insert_y1, 0), secondBrId)
                                    {
                                        ScaleFactors = new Scale3d(1, 1, 1)
                                    };
                                    currentSpace.AppendEntity(secondBr1);
                                    tr.AddNewlyCreatedDBObject(secondBr1, true);
                                }
                            }
                            if (attRef.Tag == "PHASE1_FIXING_POSITION")
                            {
                                string[] coordinates = attRef.TextString.Split(',');
                                if (double.TryParse(coordinates[0], out insert_x2) && double.TryParse(coordinates[1], out insert_y2))
                                {
                                    BlockReference secondBr2 = new BlockReference(new Point3d(xMain + insert_x2, yMain + insert_y2, 0), secondBrId)
                                    {
                                        ScaleFactors = new Scale3d(-1, 1, 1)
                                    };
                                    currentSpace.AppendEntity(secondBr2);
                                    tr.AddNewlyCreatedDBObject(secondBr2, true);
                                }
                            }

                        }
                    }
                    catch (System.Exception ex)
                    {
                        editor.WriteMessage($"{ex}");
                    }
                    tr.Commit();
                }
                editor.WriteMessage("\nBoth drawings have been successfully loaded.");
            }
            catch (System.Exception ex)
            {

                editor.WriteMessage($"\nError: {ex.Message}");
            }
        }

    }
}

