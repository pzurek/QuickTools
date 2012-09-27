#region Imported Namespaces

//.NET common used namespaces
using System;
using System.Collections.Generic;
using System.Windows.Forms;

//Revit.NET common used namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

using View = Autodesk.Revit.DB.View;
using Application = Autodesk.Revit.ApplicationServices.Application;

using QuickTools;

#endregion

namespace QuickTools
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class PurgeUnnamedRefPlanes : IExternalCommand
    {
        UIApplication uiApp;
        UIDocument uiDoc;
        Application app;
        Document doc;

        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements)
        {
            try
            {
                uiApp = commandData.Application;
                uiDoc = uiApp.ActiveUIDocument;
                app = uiApp.Application;
                doc = uiDoc.Document;

                FilteredElementCollector refPlaneCollector = new FilteredElementCollector(doc);
                IList<Element> refPlanes = refPlaneCollector.OfClass(typeof(ReferencePlane)).ToElements();

                Transaction transaction = new Transaction(doc);
                transaction.Start("Purging unnamed reference planes");

                foreach (Element refPlane in refPlanes)
                {
                    if (refPlane.Name == "Reference Plane")
                        doc.Delete(refPlane);
                }

                transaction.Commit();

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                TaskDialog.Show("Revit Quick Tools", e.Message);
                return Result.Failed;
            }
        }
    }
}
