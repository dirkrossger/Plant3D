
// Plant 3D Assemblies that need to be referenced - ..\PlantSDK 2011\inc-win32
// (Make Copy Local property False for these dlls)
// PnP3dDataLinksManager.dll
// PnP3dObjectsMgd.dll
// PnP3dPartsRepository.dll
// PnP3dProjectPartsMgd.dll
// PnPDataLinks.dl
// PnPDataObjects.dll
// PnPProjectManagerMgd.dll
// AutoCAD dlls - (Make Copy Local property False for these dlls)
// AcDbMgd.dll 
// AcMgd.dll

// namespaces in use
using System;
using System.IO;
using System.Collections.Specialized;

#region Autodesk
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
#endregion

// Doing this because the class SpecPart is defined in here and in PartsRepository.Specification
// Then refer specify PnP3d.SpecPart or PartRep.SpecPart when using those classes.
//using Autodesk.ProcessPower.PnP3dObjects;

using PnP3d = Autodesk.ProcessPower.PnP3dObjects;

using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.ProjectManager;
using Autodesk.ProcessPower.P3dProjectParts;
using Autodesk.ProcessPower.PartsRepository;

// Doing this because the class SpecPart is defined in here and in Pnp3dObjects
// Then refer specify PnP3d.SpecPart or PartRep.SpecPart when using those classes.
//using Autodesk.ProcessPower.PartsRepository.Specification;

using PartRep = Autodesk.ProcessPower.PartsRepository.Specification;
using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.DataObjects;


namespace Plant3d_Create_a_Pipe
{
    public class Class1
    {
        [CommandMethod("NewPipe")]
        public static void NewPipe()
        {
            string strSpec = "CS300";

            PlantProject proj = PlantApplication.CurrentProject;
            PipingProject prjpart = (PipingProject)proj.ProjectParts["Piping"];

            // Current state of API does not include a .NET version of the spec 
            // manager so the lower part's repository API must be used.
            // Use project to get spec folder, for now use CS300 spec
           
            PartRep.SpecPart specPart = null;
            string strSpecFileName = Path.Combine(prjpart.SpecSheetsDirectory, strSpec + ".pspx");


            // Open the spec
            using (PartRep.PipePartSpecification spec = (PartRep.PipePartSpecification)PartRep.PipePartSpecification.OpenSpecification(strSpecFileName))
            {
                NominalDiameter nd = new NominalDiameter("in", 10.0);
                using (PartQueryResults results = spec.SelectParts("Pipe", nd))
                {
                    if (results != null)
                    {
                        // use first row in spec for a 10inch pipe
                        specPart = (PartRep.SpecPart)results.NextPart();
                    }
                }
  
                // Create the pipe and return the ObjectId 
                ObjectId newPipeId = CreatePipe();

               
               
                //try
                //{
                //    int cacheId = AddPartToProject(prjpart, strSpec, newPipeId, specPart);

                //    // if cacheId == -1 then AddPartToProject failed
                //    if (cacheId == -1)
                //    {
                //        // Failed to add part to project 
                //        Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                //        ed.WriteMessage("Error creating part in the project erasing the pipe entity");
                      
                //        // Erase the pipe entity
                //        Database db = HostApplicationServices.WorkingDatabase;
                //        using (Transaction trans = db.TransactionManager.StartTransaction())
                //        {
                //            PnP3d.Pipe p = trans.GetObject(newPipeId, OpenMode.ForWrite) as PnP3d.Pipe;
                //            if (p != null)
                //            {
                //                p.Erase();
                //                trans.Commit();
                //                return;
                //            }
                //        }
                //        return;
                //     }
                     
                //    setPartGeometry(prjpart, newPipeId, cacheId);
                  
                //    int groupId = createUnassignedLineGroup(prjpart);
                    
                //    assignToLineGroup(prjpart, newPipeId, cacheId, groupId);
                     
                //}
                //catch(Autodesk.AutoCAD.Runtime.Exception ex)
                //{
                //    // Some other problem occurred 
                //    Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                //    ed.WriteMessage("Error - erasing the pipe entity");
                //    ed.WriteMessage(ex.ToString());
                    
                //    // Erase the pipe entity
                //    Database db = HostApplicationServices.WorkingDatabase;
                //    using (Transaction trans = db.TransactionManager.StartTransaction())
                //    {
                //        PnP3d.Pipe p = trans.GetObject(newPipeId, OpenMode.ForWrite) as PnP3d.Pipe;
                //        if (p != null)
                //        {
                //            p.Erase();
                //            trans.Commit();
                //        }
                //    }

                //}

                spec.Close();
            }
        }

        private static ObjectId CreatePipe()
        {
            // Create a pipe
            PnP3d.Pipe p = new PnP3d.Pipe();
            p.StartPoint = new Autodesk.AutoCAD.Geometry.Point3d(0, 0, 0);
            p.EndPoint = new Autodesk.AutoCAD.Geometry.Point3d(100, 0, 0);
            p.OuterDiameter = 10;

            ObjectId newId = ObjectId.Null;

            // Add the pipe to the current space, either model or paper space
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                if (bt != null)
                {
                    BlockTableRecord btr = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    if (btr != null)
                    {
                        newId = btr.AppendEntity(p);
                        trans.AddNewlyCreatedDBObject(p, true);
                        trans.Commit();
                    }
                }
            }

            return newId;
        }

         private static int AddPartToProject(PipingProject prjpart, string strSpec, ObjectId partId, PartRep.SpecPart specPart)
         {
           Autodesk.ProcessPower.DataObjects.PnPDatabase db = prjpart.DataLinksManager.GetPnPDatabase();
           PartsRepository rep = PartsRepository.AttachedRepository(db, false);

            // Create new part in project
            Autodesk.ProcessPower.PartsRepository.Part part = rep.NewPart(specPart.PartType);
            rep.AutoAccept = false;

            // Assign property values in project
            StringCollection props = specPart.PropertyNames;
            for (int i = 0; i < props.Count; ++i)
            {
                PartProperty prop = rep.GetPartProperty(specPart.PartType, props[i], false);
                if (prop == null || prop.IsExpression)
                {
                    continue;   // can't be assigned
                }
                try
                {
                    part[props[i]] = specPart[props[i]];
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    // display exception on the command line
                    Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                    ed.WriteMessage(ex.ToString());
                }
            }

            // assign special spec property
            //
            try
            {
                part["Spec"] = strSpec;
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                // display exception on the command line
                Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage(ex.ToString());
            }

            // add reference to spec record only if it is not set yet
            //
            try
            {
                if (part["SpecRecordId"] == System.DBNull.Value)
                {
                    part["SpecRecordId"] = specPart.PartId;
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                // display exception on the command line
                Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage(ex.ToString());
            }


            // Ok now deal with the ports
            //
            Autodesk.ProcessPower.PartsRepository.PortCollection ports = specPart.Ports;
            Autodesk.ProcessPower.PartsRepository.Port principal_port = ports[0];

            foreach (Autodesk.ProcessPower.PartsRepository.Port port in ports)
            {
                System.Guid sizeRecId = System.Guid.Empty;
                if (string.Compare(port.Name, principal_port.Name) != 0)
                {
                    sizeRecId = (System.Guid)port["SizeRecordId"];
                }

                Autodesk.ProcessPower.PartsRepository.Port newPort = null;
                bool bNew = true;
                bool bNeedAccept = false;

                // Principal port is embedded.
                //
                if (sizeRecId != System.Guid.Empty)
                {
                    newPort = part.NewPortBySizeRecordId(port.Name, sizeRecId.ToString(), out bNew);
                    bNeedAccept = true;
                }
                else
                {
                    newPort = part.NewPort(port.Name);
                }

                if (bNew)
                {
                    foreach (string prop in port.PropertyNames)
                    {
                        if (string.Compare(prop, "PortName", true) == 0)
                        {
                            continue;   // dont copy port name
                        }
                        try
                        {
                            newPort[prop] = port[prop];
                        }
                        catch(Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            // display exception on the command line
                            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                            ed.WriteMessage(ex.ToString()); 
                        }
                    }

                    if (bNeedAccept)
                    {
                        try
                        {
                            rep.CommitPort(newPort);
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            // display exception on the command line
                            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                            ed.WriteMessage(ex.ToString());
                        }
                    }
                }

                part.Ports.Add(newPort);
            }

            // Add new part to the project database
            rep.AddPart(part);

            // Transform properties sentitive to current project's unit settings
            switch (prjpart.ProjectUnitsType)
            {
                case ProjectUnitsType.eMetric:
                    part.TransformPropertiesToUnits(Units.Mm, Units.Mm);
                    break;
                case ProjectUnitsType.eMixedMetric:
                    part.TransformPropertiesToUnits(Units.Mm, Units.Inch);
                    break;
                case ProjectUnitsType.eImperial:
                    part.TransformPropertiesToUnits(Units.Inch, Units.Inch);
                    break;
            }

            int cacheId = part.PartId;

            // Finally now we can link entity to row in project
            if (cacheId != -1)
            {
                Autodesk.ProcessPower.DataLinks.DataLinksManager dlm = prjpart.DataLinksManager;
                try
                {
                    dlm.Link(partId, cacheId);
                }
                catch
                {
                    cacheId = -1; ;
                }
            }

            return cacheId; // -1 for error
        }
        private static void setPartGeometry(PipingProject prjpart, ObjectId objId, int rowId)
        {
            Database db = objId.Database;
            DataLinksManager dlm = prjpart.DataLinksManager;
            StringCollection props = new StringCollection();
            StringCollection vals = new StringCollection();
            
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {

                    PnP3d.Part part = trans.GetObject(objId, OpenMode.ForRead) as PnP3d.Part;

                    //if (part is Pipe)
                    if (part is PnP3d.Pipe)
                    {
                        PnP3d.Pipe pipe = (PnP3d.Pipe)part;

                        props.Add("Length");
                        vals.Add(pipe.Length.ToString());

                        props.Add("CutLength");
                        vals.Add(pipe.Length.ToString());

                        props.Add("Position X");
                        vals.Add(pipe.StartPoint.X.ToString());

                        props.Add("Position Y");
                        vals.Add(pipe.StartPoint.Y.ToString());

                        props.Add("Position Z");
                        vals.Add(pipe.StartPoint.Z.ToString());
                    }
                    // else if ()    // Handle other Plant3D entity types
                    // {
                    // }
                    // else
                    // {
                    // }

                    dlm.SetProperties(rowId, props, vals);
                    trans.Commit();
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    // display exception on the command line
                    Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                    ed.WriteMessage(ex.ToString());
                    trans.Abort();
                }
            }
        }
       
        private static int createUnassignedLineGroup(PipingProject prjpart)
        {
            // Create a new line group in project database
            DataLinksManager dlm = prjpart.DataLinksManager;
            PnPDatabase db = dlm.GetPnPDatabase();

            PnPTable tbl = db.Tables["P3dLineGroup"];

            PnPRow row = tbl.NewRow();
            tbl.Rows.Add(row);

            return row.RowId;
        }

        private static void assignToLineGroup(PipingProject prjpart, ObjectId partId, int cacheId, int groupId)
        {
            // Assign an entity's row to line group
            DataLinksManager dlm = prjpart.DataLinksManager;
            PnPDatabase db = dlm.GetPnPDatabase();

            // Relate objects to line group
            dlm.Relate("P3dLineGroupPartRelationship",
                       "LineGroup", groupId,
                       "Part", cacheId);

            // Also relate drawing to line group
            Database acdb = partId.Database;
            int dwgId = dlm.GetDrawingId(acdb);
            if (dwgId != -1)
            {
                dlm.Relate("P3dDrawingLineGroupRelationship",
                            "Drawing", dwgId,
                            "LineGroup", groupId);
            }
        }

    }
}
