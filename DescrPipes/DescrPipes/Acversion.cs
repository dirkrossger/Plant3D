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
    public class Acversion
    {

        public static string FindAutoCAD()
        {
            string acadpath;

            try
            {
                RegistryKey ukey = Registry.CurrentUser;
                RegistryKey acad = ukey.OpenSubKey("SOFTWARE\\Autodesk\\AutoCAD");
                string curver = acad.GetValue("CurVer") as string;
                RegistryKey version = ukey.OpenSubKey("SOFTWARE\\Autodesk\\AutoCAD\\" + curver);
                string key = version.GetValue("CurVer") as string;

                //- We cannot read HKEY_LOCAL_MACHINE on Vista
                RegistryKey lkey = Registry.LocalMachine;
                RegistryKey acad2 = lkey.OpenSubKey("SOFTWARE\\Autodesk\\AutoCAD\\" + curver + "\\" + key);
                acadpath = acad2.GetValue("ProductName") as string;
              }
            catch
            {
                acadpath = "...not found!";
            }
            return (acadpath);
        }

    }
}
