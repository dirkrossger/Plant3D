using System;
using System.Collections.Generic;

#region Autocad
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
#endregion
using PnP3d = Autodesk.ProcessPower.PnP3dObjects;

using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.ProjectManager;
using Autodesk.ProcessPower.P3dProjectParts;
using Autodesk.ProcessPower.PartsRepository;
using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.DataObjects;

#region Autocad Plant3D

#endregion

namespace AcadNet
{
    class Select
    {
        public static List<Datas> Lines()
        {
            List<Datas> list = new List<Datas>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database acadDb = doc.Database;
            Editor ed = doc.Editor;

            TypedValue[] filterlist = new TypedValue[1];
            filterlist[0] = new TypedValue(0, "LINE");
            SelectionFilter filter = new SelectionFilter(filterlist);

            PromptSelectionResult psr = ed.SelectAll(filter);
            SelectionSet sset = null;
            //Entity ent = null;
            Line line = null;

            if (psr.Status == PromptStatus.OK)
            {
                sset = psr.Value;
            }

            using (Transaction tr = acadDb.TransactionManager.StartTransaction())
            {
                if (sset != null)
                {
                    for (int i = 0; i < sset.Count; i++)
                    {
                        try
                        {
                            line = tr.GetObject(sset[i].ObjectId, OpenMode.ForWrite) as Line;
                            if (line != null)
                            {
                                list.Add(new Datas { Position = line.StartPoint, Content = line.Layer });
                            }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage("\nËrror: " + ex);
                        }

                    }

                }
                tr.Commit();


                return list;
            }
        }

        public static List<Datas> Pipes()
        {
            // Get the AutoCAD editor
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acadDb = doc.Database;
            Editor ed = doc.Editor;
            List<Datas> list = new List<Datas>();

            //http://adndevblog.typepad.com/autocad/wayne-brill/
            Project currentProject = PlantApplication.CurrentProject.ProjectParts["PnId"];
            DataLinksManager dlm = currentProject.DataLinksManager;
            PnPDatabase pnpdb = dlm.GetPnPDatabase();

            TypedValue[] filterlist = new TypedValue[1];
            filterlist[0] = new TypedValue(0, "ACPPPIPE");
            SelectionFilter filter = new SelectionFilter(filterlist);

            PromptSelectionResult psr = ed.SelectAll(filter);
            SelectionSet sset = null;
            //Entity ent = null;
            //Line line = null;

            if (psr.Status == PromptStatus.OK)
            {
                sset = psr.Value;
            }

            using (Transaction trans = acadDb.TransactionManager.StartTransaction())
            {
                if (sset != null)
                {
                    for (int i = 0; i < sset.Count; i++)
                    {
                        try
                        {
                            PnP3d.Pipe supP3DPipe = trans.GetObject(sset[i].ObjectId, OpenMode.ForRead) as PnP3d.Pipe;
                            if (supP3DPipe != null)
                            {
                                Point3d start = supP3DPipe.CenterOfGravity;
                                Point3d end = supP3DPipe.EndPoint;
                                double len = supP3DPipe.Length;
                                string lay = supP3DPipe.Layer;

                                list.Add(new Datas { Position = start, Content = lay });

                                //ed.WriteMessage(
                                //    "\nProperties of selected Pipe: " +
                                //    "\nStartpoint: " + start +
                                //    "\nEndpoint: " + end +
                                //    "\nLength: " + len +
                                //    "\nLayer: " + lay
                                //    );

                            }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage("\nËrror: " + ex);
                        }

                    }

                }
                trans.Commit();


            }
            return list;
        }

        public static List<Datas> Insert()
        {
            List<Datas> list = new List<Datas>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database acadDb = doc.Database;
            Editor ed = doc.Editor;

            TypedValue[] filterlist = new TypedValue[1];
            filterlist[0] = new TypedValue(0, "INSERT");
            SelectionFilter filter = new SelectionFilter(filterlist);

            PromptSelectionResult psr = ed.SelectAll(filter);
            SelectionSet sset = null;
            //Entity ent = null;
            BlockReference insert = null;

            if (psr.Status == PromptStatus.OK)
            {
                sset = psr.Value;
            }

            using (Transaction tr = acadDb.TransactionManager.StartTransaction())
            {
                if (sset != null)
                {
                    for (int i = 0; i < sset.Count; i++)
                    {
                        try
                        {
                            insert = tr.GetObject(sset[i].ObjectId, OpenMode.ForWrite) as BlockReference;
                            if (insert != null)
                            {
                                list.Add(new Datas { Position = insert.Position, Content = insert.Layer });
                            }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage("\nËrror: " + ex);
                        }

                    }

                }
                tr.Commit();


                return list;
            }
        }
    }
}
