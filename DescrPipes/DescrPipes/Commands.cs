using System;
using System.Collections.Generic;

#region Autodesk
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#endregion


namespace AcadNet
{

    public class Commands
    {
        [CommandMethod("Command1")]
        public void DescLayerLines()
        {
            var obj = AcApp.AcadApplication;
            Type type = obj.GetType();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            List<Datas> list = Select.Pipes();
            if(list != null)
            {
                foreach (Datas x in list)
                    Mleader.Create(x.Position, x.Content);

                ed.WriteMessage("\n" + list.Count + " Mleaders created.");
            }
            else
            {
                ed.WriteMessage("\nNo Pipes found!");
            }
        }
    }
}
