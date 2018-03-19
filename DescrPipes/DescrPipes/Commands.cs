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
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            string version = Acversion.FindAutoCAD();
            if (version != null)
                if (version.Contains("Plant 3D"))
                {
                    ed.WriteMessage("\nFunction Mleader pipes loaded. Type->Command1");
                }
                else
                {
                    ed.WriteMessage("\nFunction Mleader runs only in Autocad Plant3D");
                    return;
                }

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

    public class AcadNet : IExtensionApplication
    {
        public void Initialize()

        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            string version = Acversion.FindAutoCAD();
            if (version != null)
                if (version.Contains("Plant 3D"))
                {
                    ed.WriteMessage("\nFunction Mleader pipes loaded. Type->Command1");
                }
                else
                {
                    ed.WriteMessage("\nFunction Mleader runs only in Autocad Plant3D");
                }
        }
    



        public void Terminate()

        {

        }

    }
}
