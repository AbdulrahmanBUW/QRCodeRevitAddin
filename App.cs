using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;

namespace QRCodeRevitAddin
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                string tabName = "QR Tools";
                try
                {
                    application.CreateRibbonTab(tabName);
                }
                catch (Exception)
                {
                }
                RibbonPanel panel = application.CreateRibbonPanel(tabName, "QR Code Operations");

                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                string assemblyDir = Path.GetDirectoryName(assemblyPath);

                PushButtonData buttonData1 = new PushButtonData(
                    "GenerateQRCode",
                    "Generate\nQR Code",
                    assemblyPath,
                    "QRCodeRevitAddin.Commands.ShowQrWindowCommand");

                buttonData1.ToolTip = "Open the QR Code generator window to create and insert QR codes into sheets";
                buttonData1.LongDescription = "Opens a window where you can manually enter information or automatically " +
                    "extract data from the current sheet to generate a QR code. The QR code can be previewed, saved as " +
                    "a PNG file, or inserted directly into the active sheet.";

                SetButtonIcon(buttonData1, assemblyDir, "qr-icon-32.png");

                PushButton button1 = panel.AddItem(buttonData1) as PushButton;

                panel.AddSeparator();

                PushButtonData buttonData2 = new PushButtonData(
                    "QuickInsertQR",
                    "Quick Insert\nfrom Sheet",
                    assemblyPath,
                    "QRCodeRevitAddin.Commands.InsertQrFromSheetCommand");

                buttonData2.ToolTip = "Quickly generate and insert a QR code using data from the current sheet";
                buttonData2.LongDescription = "Automatically extracts sheet number, sheet name, revision, and current date " +
                    "from the selected sheet, generates a QR code, and opens the window pre-filled with this data. " +
                    "Perfect for rapid QR code insertion on multiple sheets.";

                SetButtonIcon(buttonData2, assemblyDir, "qr-icon-32.png");

                PushButton button2 = panel.AddItem(buttonData2) as PushButton;

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error",
                    $"Failed to initialize QR Code Add-in:\n\n{ex.Message}\n\n{ex.StackTrace}");
                return Result.Failed;
            }
        }

        private void SetButtonIcon(PushButtonData buttonData, string assemblyDir, string iconFileName)
        {
            string[] possiblePaths = new string[]
            {
                Path.Combine(assemblyDir, "Resources", iconFileName),
                Path.Combine(assemblyDir, iconFileName),
                Path.Combine(assemblyDir, "..", "Resources", iconFileName)
            };

            foreach (string iconPath in possiblePaths)
            {
                try
                {
                    if (File.Exists(iconPath))
                    {
                        Uri iconUri = new Uri(iconPath, UriKind.Absolute);
                        BitmapImage icon = new BitmapImage(iconUri);
                        buttonData.LargeImage = icon;
                        return; // Success, exit
                    }
                }
                catch
                {
                    continue;
                }
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            // Cleanup if needed
            return Result.Succeeded;
        }
    }
}