using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;

namespace QRCodeRevitAddin
{
    /// <summary>
    /// Main application class that implements IExternalApplication.
    /// Handles add-in initialization and creates the ribbon UI.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class App : IExternalApplication
    {
        /// <summary>
        /// Called when Revit starts up. Initializes the add-in and creates ribbon UI.
        /// </summary>
        /// <param name="application">The Revit application object</param>
        /// <returns>Success or failure status</returns>
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Create custom ribbon tab
                string tabName = "QR Tools";
                try
                {
                    application.CreateRibbonTab(tabName);
                }
                catch (Exception)
                {
                    // Tab might already exist, continue
                }

                // Create ribbon panel
                RibbonPanel panel = application.CreateRibbonPanel(tabName, "QR Code Operations");

                // Get assembly path for button images and commands
                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                string assemblyDir = Path.GetDirectoryName(assemblyPath);

                // Create "Generate QR Code" button
                PushButtonData buttonData1 = new PushButtonData(
                    "GenerateQRCode",
                    "Generate\nQR Code",
                    assemblyPath,
                    "QRCodeRevitAddin.Commands.ShowQrWindowCommand");

                buttonData1.ToolTip = "Open the QR Code generator window to create and insert QR codes into sheets";
                buttonData1.LongDescription = "Opens a window where you can manually enter information or automatically " +
                    "extract data from the current sheet to generate a QR code. The QR code can be previewed, saved as " +
                    "a PNG file, or inserted directly into the active sheet.";

                // Set button image (try to load custom icon, fallback to no icon)
                try
                {
                    string iconPath = Path.Combine(assemblyDir, "Resources", "qr-icon-32.png");
                    if (File.Exists(iconPath))
                    {
                        Uri iconUri = new Uri(iconPath);
                        BitmapImage icon = new BitmapImage(iconUri);
                        buttonData1.LargeImage = icon;
                    }
                }
                catch
                {
                    // Icon loading failed, continue without icon
                }

                PushButton button1 = panel.AddItem(buttonData1) as PushButton;

                // Add separator
                panel.AddSeparator();

                // Create "Quick Insert from Sheet" button
                PushButtonData buttonData2 = new PushButtonData(
                    "QuickInsertQR",
                    "Quick Insert\nfrom Sheet",
                    assemblyPath,
                    "QRCodeRevitAddin.Commands.InsertQrFromSheetCommand");

                buttonData2.ToolTip = "Quickly generate and insert a QR code using data from the current sheet";
                buttonData2.LongDescription = "Automatically extracts sheet number, sheet name, revision, and current date " +
                    "from the selected sheet, generates a QR code, and opens the window pre-filled with this data. " +
                    "Perfect for rapid QR code insertion on multiple sheets.";

                // Set button image
                try
                {
                    string iconPath = Path.Combine(assemblyDir, "Resources", "qr-icon-32.png");
                    if (File.Exists(iconPath))
                    {
                        Uri iconUri = new Uri(iconPath);
                        BitmapImage icon = new BitmapImage(iconUri);
                        buttonData2.LargeImage = icon;
                    }
                }
                catch
                {
                    // Icon loading failed, continue without icon
                }

                PushButton button2 = panel.AddItem(buttonData2) as PushButton;

                // Log successful initialization
                TaskDialog.Show("QR Code Add-in", 
                    "QR Code Add-in loaded successfully!\n\n" +
                    "Look for the 'QR Tools' tab in the ribbon.",
                    TaskDialogCommonButtons.Ok);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", 
                    $"Failed to initialize QR Code Add-in:\n\n{ex.Message}\n\n{ex.StackTrace}");
                return Result.Failed;
            }
        }

        /// <summary>
        /// Called when Revit shuts down. Performs cleanup.
        /// </summary>
        /// <param name="application">The Revit application object</param>
        /// <returns>Success or failure status</returns>
        public Result OnShutdown(UIControlledApplication application)
        {
            // Cleanup if needed
            return Result.Succeeded;
        }
    }
}
