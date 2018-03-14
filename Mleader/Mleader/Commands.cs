using System;
#region Autodesk
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
#endregion

namespace Mleader
{
    public class Commands
    {
        [CommandMethod("netTextMLeader")]
        public static void netTextMLeader()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                BlockTable table = Tx.GetObject(
                db.BlockTableId,
                OpenMode.ForRead)
                as BlockTable;

                BlockTableRecord model = Tx.GetObject(
                table[BlockTableRecord.ModelSpace],
                OpenMode.ForWrite)
                as BlockTableRecord;

                MLeader leader = new MLeader();
                leader.SetDatabaseDefaults();

                leader.ContentType = ContentType.MTextContent;

                MText mText = new MText();
                mText.SetDatabaseDefaults();
                //mText.Width = 100;
                //mText.Height = 50;
                mText.SetContentsRtf("MLeader");
                mText.Location = new Point3d(1, 1, 0);

                leader.MText = mText;

                int idx = leader.AddLeaderLine(new Point3d(0, 0, 0));
                leader.AddFirstVertex(idx, new Point3d(0, 0, 0));

                model.AppendEntity(leader);
                Tx.AddNewlyCreatedDBObject(leader, true);

                Tx.Commit();
            }
        }

        [CommandMethod("creaMl")]
        public void CreateMleader()
        {
            Point3d pt = new Point3d(10, 10, 10);
            Mleader.Create(pt, "Test");
        }
    }
}
