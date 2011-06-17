#region Imported Namespaces

//.NET common used namespaces
using System;
using System.Collections.Generic;
using System.Linq;
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
    public class QuickWipe : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            FilteredElementCollector viewCollector = new FilteredElementCollector(doc);
            IList<Element> views = viewCollector.OfClass(typeof(View)).ToElements();

            IEnumerable<View> realViews = views.Cast<View>();

            Stack<View> allViews = new Stack<View>();
            foreach (View view in realViews)
            {
                if (view.ViewName.StartsWith("{3D}"))
                    continue;
                else
                    allViews.Push(view);
            }

            Transaction transaction = new Transaction(doc);
            transaction.Start("Wiping clean");

            int counter = 0;

            while (allViews.Count > 0)
            {
                try
                {
                    doc.Delete(allViews.Pop());
                }
                catch (Exception e)
                {
                    TaskDialog.Show("Problem during deletion", string.Format("Object {1} no.{0} could not be deleted",
                        counter.ToString(),
                        e.Source.ToString()));
                }
            }

            transaction.Commit();

            return Result.Succeeded;
        }
    }
}
