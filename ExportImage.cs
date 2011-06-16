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
    public class ExportImage : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc.Document;

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Title = "Save view as file";
                saveFileDialog.Filter = "Png files (*.png)|*.png";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ImageExportOptions imageExportOptions = new ImageExportOptions();
                    imageExportOptions.ExportRange = ExportRange.VisibleRegionOfCurrentView;
                    imageExportOptions.FilePath = saveFileDialog.FileName;
                    imageExportOptions.HLRandWFViewsFileType = ImageFileType.PNG;
                    imageExportOptions.ImageResolution = ImageResolution.DPI_72;
                    imageExportOptions.ShadowViewsFileType = ImageFileType.PNG;
                    imageExportOptions.ZoomType = ZoomFitType.Zoom;
                    imageExportOptions.Zoom = 100;

                    doc.ExportImage(imageExportOptions);
                }
            }
            catch (Exception e)
            {
                TaskDialog.Show("Revit Quick Tools", e.Message);
                return Result.Failed;
            }

            return Result.Succeeded;
        }
    }
}
