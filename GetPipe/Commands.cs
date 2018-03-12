using System;
using System.Collections.Generic;

#region Autodesk
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
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
using Autodesk.AutoCAD.Geometry;

namespace WalkPipes
{
    public class Commands
    {
        [CommandMethod("PipeWalk")]
        public void pipeWalk()
        {
            // 4.1 Declare a variable as a PlantProject. Instantiate it using
            // the CurrentProject of PlantApplication
            PlantProject mainPrj = PlantApplication.CurrentProject;

            // 4.2 Declare a Project and instantiate it using  
            // ProjectParts[] of the PlantProject from step 4.1
            // use "Piping" for the name. This will get the Piping project 
            Project prj = mainPrj.ProjectParts["Piping"];

            // 4.3 Declare a variable as a DataLinksManager. Instantiate it using
            // the DataLinksManager property of the Project from 4.2.
            DataLinksManager dlm = prj.DataLinksManager;

            //  PipingProject pipingPrj = (PipingProject) mainPrj.ProjectParts["Piping"];
            //  DataLinksManager dlm = pipingPrj.DataLinksManager;


            // Get the TransactionManager
            // Autodesk.AutoCAD.DatabaseServices.TransactionManager tm =
            //AcadApp.DocumentManager.MdiActiveDocument.Database.TransactionManager;

            // Get the AutoCAD editor
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;


            // Prompt the user to select a pipe entity
            PromptEntityOptions pmtEntOpts = new PromptEntityOptions("Select a Pipe : ");
            PromptEntityResult pmtEntRes = ed.GetEntity(pmtEntOpts);
            if (pmtEntRes.Status == PromptStatus.OK)
            {
                // Get the ObjectId of the selected entity
                ObjectId entId = pmtEntRes.ObjectId;

                // Use the using statement and start a transaction
                // Use the transactionManager created above (tm)
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    //#region transaction
                    //try
                    //{

                    //    // 4.4 Declare a variable as a Part. Instantiate it using
                    //    // the GetObject Method of the Transaction created above (tr)
                    //    // for the ObjectId argument use the ObjectId from above (entId)
                    //    // Open it for read. (need to cast it to Part)
                    //Part pPart = tr.GetObject(entId, OpenMode.ForRead);

                    //    // 4.5 Declare a variable as a PortCollection. Instantiate it 
                    //    // using the GetPorts method of the Part from step 4.4.
                    //    // use PortType.All for the PortType.
                    //    PortCollection portCol = pPart.GetPorts(PortType.All); // (PortType.Both);

                    //    // 4.6 Use the WriteMessage function of the Editor created above (ed)
                    //    // print the Count property of the PortCollection from step 4.5.
                    //    // use a string similar to this: "\n port collection count = "
                    //    ed.WriteMessage("\n port collection count = " + portCol.Count);

                    //    // 4.7 Declare a ConnectionManager variable. 
                    //    // (Autodesk.ProcessPower.PnP3dObjects.ConnectionManager)
                    //    // Instantiate the ConnectionManager variable by making it
                    //    // equal to a new Autodesk.ProcessPower.PnP3dObjects.ConnectionManager();
                    //    ConnectionManager conMgr = new Autodesk.ProcessPower.PnP3dObjects.ConnectionManager();

                    //    // 4.8 Declare a bool variable named bPartIsConnected and make it false
                    //    bool bPartIsConnected = false;


                    //    // 4.9 Use a foreach loop and iterate through all of the Port in
                    //    // the PortCollection from step 4.5.
                    //    // Note: Put the closing curly brace below step 4.18
                    //    foreach (Port pPort in portCol)
                    //    {
                    //        // 4.10 Use the WriteMessage function of the Editor created above (ed)
                    //        // print the Name property of the Port (looping through the ports) 
                    //        // use a string similar to this: "\nName of this Port = " +
                    //        ed.WriteMessage("\nName of this Port = " + pPort.Name);

                    //        // 4.11 Use the WriteMessage function of the Editor created above (ed)
                    //        // print the X property of the Position from the Port
                    //        // use a string similar to this: "\nX of this Port = " + 
                    //        ed.WriteMessage("\nX of this Port = " + pPort.Position.X.ToString());

                    //        // 4.12 Declare a variable as a Pair and make it equal to a 
                    //        // new Pair().
                    //        Pair pair1 = new Pair();

                    //        // 4.13 Make the ObjectId property of the Pair created in step 4.10 
                    //        // equal to the ObjectId of the selected Part (entId)
                    //        pair1.ObjectId = entId;

                    //        // 4.14 Make the Port property of the Pair created in step 4.10
                    //        // equal to the port from the foreach cycle (step 4.7)
                    //        pair1.Port = pPort;


                    //        // 4.15 Use an if else and the IsConnected method of the ConnectionManager
                    //        // from step 4.7. Pass in the Pair from step 4.12
                    //        // Note: Put the else statement below step 4.17 and the closing curly
                    //        // brace for the else below step 4.18
                    //        if (conMgr.IsConnected(pair1))
                    //        {
                    //            // 4.16 Use the WriteMessage function of the Editor (ed)
                    //            // and put this on the command line:
                    //            // "\n Pair is connected "
                    //            ed.WriteMessage("\n Pair is connected ");

                    //            // 4.17 Make the bool from step 4.8 equal to true.
                    //            // This is used in an if statement in step 4.19.
                    //            bPartIsConnected = true;
                    //        }
                    //        else
                    //        {
                    //            // 4.18 Use the WriteMessage function of the Editor (ed)
                    //            // and put this on the command line:
                    //            // "\n Pair is NOT connected "
                    //            ed.WriteMessage("\n Pair is NOT connected ");
                    //        }

                    //    }


                    //    // 4.19 Use an If statement and the bool from step 4.8. This will be 
                    //    // true if one of the pairs tested in loop above loop was connected. 
                    //    // Note: Put the closing curly brace after step 4.26
                    //    if (bPartIsConnected)
                    //    {

                    //        // 4.20 Declare an ObjectId named curObjID make it 
                    //        // equal to ObjectId.Null
                    //        ObjectId curObjId = ObjectId.Null;


                    //        // 4.21 Declare an int name it rowId 
                    //        int rowId;

                    //        // 4.22 Declare a variable as a  ConnectionIterator instantiate it
                    //        // using the NewIterator method of ConnectionIterator (Autodesk.ProcessPower.PnP3dObjects.)
                    //        // Use the ObjectId property of the Part from step 4.4 
                    //        ConnectionIterator connectIter = ConnectionIterator.NewIterator(pPart.ObjectId);                       //need PnP3dObjectsMgd.dll

                    //        // You could Also use this, need to ensure that pPort is connected
                    //        // Use the ConnectionManager and a Pair as in the example above.
                    //        // conIter = ConnectionIterator.NewIterator(pPart.ObjectId, pPort);

                    //        // 4.23 Use a for loop and loop through the connections in the 
                    //        // ConnectionIterator from step 4.22. The initializer can be empty.
                    //        // Use !.Done for the condition. use .Next for the iterator.  
                    //        // Note: Put the closing curly brace after step 4.26
                    //        for (; !connectIter.Done(); connectIter.Next())
                    //        {

                    //            // 4.24 Make the ObjectId from step 4.20 equal to the ObjectId
                    //            // property of the ConnectionIterator
                    //            curObjId = connectIter.ObjectId;

                    //            // 4.25 Make the integer from step 4.21 equal to the return from 
                    //            // FindAcPpRowId method of the DataLinksManager from step 4.3.
                    //            // pass in the ObjectId from step 4.24
                    //            rowId = dlm.FindAcPpRowId(curObjId);

                    //            //4.26 Use the WriteMessage function of the Editor (ed)
                    //            // and pring the integer from step 4.25. Use a string similar to this
                    //            // this on the command line:
                    //            // "\n PnId = " +
                    //            ed.WriteMessage("\n PnId = " + rowId);

                    //        }

                    //    }
                    //}
                    //catch (System.Exception ex)
                    //{

                    //    ed.WriteMessage(ex.ToString());
                    //}
                    //#endregion
                }
            }
        }

        [CommandMethod("PipeGet")]
        public void pipeGet()
        {
            // Get the AutoCAD editor
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            //http://adndevblog.typepad.com/autocad/wayne-brill/
            Project currentProject = PlantApplication.CurrentProject.ProjectParts["PnId"];
            DataLinksManager dlm = currentProject.DataLinksManager;
            PnPDatabase pnpdb = dlm.GetPnPDatabase();

            // Prompt the user to select a pipe entity
            PromptEntityOptions pmtEntOpts = new PromptEntityOptions("Select a Pipe : ");
            PromptEntityResult pmtEntRes = ed.GetEntity(pmtEntOpts);
            if (pmtEntRes.Status == PromptStatus.OK)
            {
                // Get the ObjectId of the selected entity
                ObjectId entId = pmtEntRes.ObjectId;

                // Use the using statement and start a transaction
                // Use the transactionManager created above (tm)
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    PnP3d.Pipe supP3DPipe = trans.GetObject(entId, OpenMode.ForRead) as PnP3d.Pipe;
                    if (supP3DPipe == null)
                    {
                        ed.WriteMessage("\n   Selected entity is not a Plant 3D Pipe");
                    }
                    else
                    {
                        Point3d start = supP3DPipe.StartPoint;
                        Point3d end = supP3DPipe.EndPoint;
                        double len = supP3DPipe.Length;
                        string lay = supP3DPipe.Layer;

                        ed.WriteMessage(
                            "\nProperties of selected Pipe: " +
                            "\nStartpoint: " + start +
                            "\nEndpoint: " + end +
                            "\nLength: " + len +
                            "\nLayer: " + lay
                            );

                    }
                    trans.Commit();
                }
            }
        }
    }
}
