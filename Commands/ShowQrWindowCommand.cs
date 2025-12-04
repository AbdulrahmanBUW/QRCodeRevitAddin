using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using QRCodeRevitAddin.Views;

namespace QRCodeRevitAddin.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ShowQrWindowCommand : IExternalCommand
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

                QrWindow.Show(uiDoc, autoFillFromSheet: false);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Failed to open QR Code window: {ex.Message}";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }
        }
    }
}
