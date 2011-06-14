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
    public class HideCropBoxes : IExternalCommand
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

                CropBoxVisibility.setCropBoxVisibility(doc, false);
            }
            catch (Exception e)
            {
                TaskDialog.Show("Revit Quick Tools", e.Message);
                return Result.Failed;
            }

            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ShowCropBoxes : IExternalCommand
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

                CropBoxVisibility.setCropBoxVisibility(doc, true);
            }
            catch (Exception e)
            {
                TaskDialog.Show("Revit Quick Tools", e.Message);
                return Result.Failed;
            }

            return Result.Succeeded;
        }
    }

    public class CropBoxVisibility
    {
        public static void setCropBoxVisibility(Document doc, bool visible)
        {
            FilteredElementCollector viewportCollector = new FilteredElementCollector(doc);
            IList<Element> viewports = viewportCollector.OfCategory(BuiltInCategory.OST_Viewports).ToElements();

            Transaction transaction = new Transaction(doc);
            transaction.Start(visible ? "Showing crop boxes" : "Hiding crop boxes");

            foreach (Element element in viewports)
            {
                if (element.get_Parameter("Crop Region Visible") != null)
                    element.get_Parameter("Crop Region Visible").Set(visible ? 1 : 0);
            }

            transaction.Commit();

            System.Media.SystemSounds.Asterisk.Play();
        }
    }
}
