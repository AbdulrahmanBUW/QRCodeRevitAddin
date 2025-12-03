using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using QRCodeRevitAddin.Views;

namespace QRCodeRevitAddin.Commands
{
    /// <summary>
    /// External command to open the QR Code generator window.
    /// This is the main entry point for the manual QR generation workflow.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ShowQrWindowCommand : IExternalCommand
    {
        /// <summary>
        /// Executes the command to show the QR window.
        /// </summary>
        /// <param name="commandData">Command data from Revit</param>
        /// <param name="message">Error message to return to Revit</param>
        /// <param name="elements">Elements to highlight in case of failure</param>
        /// <returns>Result indicating success or failure</returns>
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                // Get UI document
                UIDocument uiDoc = commandData.Application.ActiveUIDocument;
                
                if (uiDoc == null)
                {
                    TaskDialog.Show("Error", "No active Revit document found.\n\nPlease open a project first.");
                    return Result.Failed;
                }

                // Show the QR window without auto-fill
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
