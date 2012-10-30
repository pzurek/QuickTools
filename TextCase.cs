#region Imported Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

using View = Autodesk.Revit.DB.View;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Document = Autodesk.Revit.DB.Document;

using QuickTools;

#endregion

namespace QuickTools
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class LowerCaseText : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            Selection sel = uiDoc.Selection;

            return TextCase.SetTextCase(doc, sel, "lower");
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class UpperCaseText : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            Selection sel = uiDoc.Selection;

            return TextCase.SetTextCase(doc, sel, "upper");
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class TitleCaseText : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            Selection sel = uiDoc.Selection;

            return TextCase.SetTextCase(doc, sel, "title");
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SentenceCaseText : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements)
        {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc.Document;
                Selection sel = uiDoc.Selection;

                return TextCase.SetTextCase(doc, sel, "sentence");
        }
    }

    public class TextCase
    {
        public static Result SetTextCase(Document doc, Selection selection, string textCase)
        {
            try
            {
                List<Element> textElements;

                if (selection.Elements.IsEmpty)
                {
                    TaskDialog selectAllDialog = new TaskDialog("Select all?");
                    selectAllDialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
                    selectAllDialog.MainInstruction = "No elements have been selected. Do you want to select all text elements in the project?";
                    selectAllDialog.MainContent = string.Format("Press YES if you want to change all text elements in the project to {0} case or press NO if you want to cancel.", textCase);

                    if (selectAllDialog.Show() == TaskDialogResult.Yes)
                    {
                        FilteredElementCollector textSelector = new FilteredElementCollector(doc);
                        textElements = textSelector.OfClass(typeof(TextElement)).ToElements().ToList<Element>();
                        if (textElements == null || textElements.ToList<Element>().Count == 0)
                            return Result.Cancelled;
                    }
                    else
                    {
                        return Result.Cancelled;
                    }
                }
                else
                {
                    textElements = new List<Element>();
                    foreach (Element element in selection.Elements)
                        textElements.Add(element);
                }

                Transaction transaction = new Transaction(doc);
                transaction.Start(string.Format("Setting text to {0} case", textCase));

                foreach (Element element in textElements)
                {
                    TextElement item = element as TextElement;

                    if (item == null)
                        continue;
                    else
                    {
                        switch (textCase)
                        {
                            case ("upper"):
                                item.Text = item.Text.ToUpper();
                                break;
                            case ("lower"):
                                item.Text = item.Text.ToLower();
                                break;
                            case ("title"):
                                item.Text = ToTitleCase(item.Text);
                                break;
                            default:
                                item.Text = ToSentenceCase(item.Text);
                                break;
                        }
                    }
                }
                transaction.Commit();

                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }
        }

        public static string ToTitleCase(string text)
        {
            return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(text.ToLower());
        }

        public static string ToSentenceCase(string text)
        {
            var lower = text.ToLower();
            bool isNewSentence = true;
            string result = "";

            for (int i = 0; i < lower.Length; i++)
            {
                switch (lower[i])
                {
                    case '.':
                        isNewSentence = true;
                        result = result + lower[i];
                        break;
                    case ':':
                        isNewSentence = true;
                        result = result + lower[i];
                        break;
                    case ' ':
                        result = result + lower[i];
                        break;
                    case '\r':
                        result = result + lower[i];
                        break;
                    case '\n':
                        result = result + lower[i];
                        break;
                    default:
                        if (isNewSentence)
                        {
                            result = result + lower[i].ToString().ToUpper();
                            isNewSentence = false;
                        }
                        else
                        {
                            result = result + lower[i];
                        }
                        break;
                }
            }
            return result;
        }
    }
}
