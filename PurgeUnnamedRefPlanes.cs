#region Imported Namespaces

//.NET common used namespaces
using System;
using System.Collections.Generic;

//Revit.NET common used namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using Application = Autodesk.Revit.ApplicationServices.Application;

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

                ElementId nameParamterId = new ElementId(BuiltInParameter.DATUM_TEXT);
                ParameterValueProvider pvp = new ParameterValueProvider(nameParamterId);
                FilterStringRuleEvaluator evaluator = new FilterStringEquals();

                FilterRule rule = new FilterStringRule(pvp, evaluator, "", false);
                ElementFilter filter = new ElementParameterFilter(rule);

                FilteredElementCollector unnamedRefPlaneCollector = new FilteredElementCollector(doc);
                unnamedRefPlaneCollector.OfClass(typeof(ReferencePlane));
                unnamedRefPlaneCollector.WherePasses(filter);
                IList<Element> unnamedRefPlanes = unnamedRefPlaneCollector.ToElements();

                Transaction transaction = new Transaction(doc);
                transaction.Start("Purging unnamed reference planes");

                foreach (Element refPlane in unnamedRefPlanes)
                     doc.Delete(refPlane.Id);
 
                transaction.Commit();

                TaskDialog.Show("Purge Unnamed Ref Planes",
                                String.Format("{0} reference planes were found in the project.\n{1} unnamed reference planes were deleted.",
                                               refPlanes.Count.ToString(), unnamedRefPlanes.Count.ToString()));

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
