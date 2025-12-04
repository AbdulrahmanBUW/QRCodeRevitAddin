using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using QRCodeRevitAddin.Views;

namespace QRCodeRevitAddin.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class InsertQrFromSheetCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                UIDocument uiDoc = commandData.Application.ActiveUIDocument;
                
                if (uiDoc == null)
                {
                    TaskDialog.Show("Error", "No active Revit document found.\n\nPlease open a project first.");
                    return Result.Failed;
                }

                Document doc = uiDoc.Document;

                // Check if current view is a sheet
                ViewSheet sheet = doc.ActiveView as ViewSheet;
                
                if (sheet == null)
                {
                    TaskDialog.Show("Not a Sheet View",
                        "Please switch to a Sheet view before using this command.\n\n" +
                        "Quick Insert from Sheet only works when you have a sheet view active.\n\n" +
                        "To use this command:\n" +
                        "1. Open a sheet view in Revit\n" +
                        "2. Click this button again\n" +
                        "3. The window will open with sheet data pre-filled",
                        TaskDialogCommonButtons.Ok);
                    return Result.Cancelled;
                }

                QrWindow.Show(uiDoc, autoFillFromSheet: true);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Failed to insert QR from sheet: {ex.Message}";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }
        }
    }
}
