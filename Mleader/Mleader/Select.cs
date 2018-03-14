using System;
using System.Collections.Generic;
#region Autodesk
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
#endregion

namespace Mleader
{
    class Select
    {
        public static void Line()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            TypedValue[] filterlist = new TypedValue[1];
            filterlist[0] = new TypedValue(0, "LINE");
            SelectionFilter filter = new SelectionFilter(filterlist);

            PromptSelectionResult selRes = ed.SelectAll(filter);

            if (selRes.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nerror in getting the selectAll");
                return;
            }
            ObjectId[] ids = selRes.Value.GetObjectIds();

            ed.WriteMessage("No entity found: " + ids.Length.ToString() + "\n");



        }
    }
}
