using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using QRCodeRevitAddin.Views;
using QRCodeRevitAddin.Utils;

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
                Logger.LogInfo("ShowQrWindowCommand executed");

                UIDocument uiDoc = commandData.Application.ActiveUIDocument;

                if (uiDoc == null)
                {
                    Logger.LogError("No active UIDocument");
                    TaskDialog.Show("Error", "No active Revit document found.\n\nPlease open a project first.");
                    return Result.Failed;
                }

                QrWindow.Show(uiDoc, autoFillFromSheet: false);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                Logger.LogError("ShowQrWindowCommand failed", ex);
                message = $"Failed to open QR Code window: {ex.Message}";
                TaskDialog.Show("Error", message);
                return Result.Failed;
            }
        }
    }
}