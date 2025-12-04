using System;
using System.Windows;
using Autodesk.Revit.UI;
using QRCodeRevitAddin.ViewModels;
using QRCodeRevitAddin.Utils;

namespace QRCodeRevitAddin.Views
{
    public partial class QrWindow : Window
    {
        public QrWindow(QrWindowViewModel viewModel)
        {
            try
            {
                InitializeComponent();
                DataContext = viewModel;
                Logger.LogInfo("QR Window initialized");
            }
            catch (Exception ex)
            {
                Logger.LogError("Failed to initialize QR Window", ex);
                throw;
            }
        }

        public static void Show(UIDocument uiDoc, bool autoFillFromSheet = false)
        {
            if (uiDoc == null)
            {
                Logger.LogError("UIDocument is null");
                TaskDialog.Show("Error", "No active Revit document found.");
                return;
            }

            try
            {
                Logger.LogInfo($"Opening QR Window (autoFill: {autoFillFromSheet})");
                QrWindowViewModel viewModel = new QrWindowViewModel(uiDoc, autoFillFromSheet);
                QrWindow window = new QrWindow(viewModel);
                window.Show();
            }
            catch (Exception ex)
            {
                Logger.LogError("Failed to open QR Code window", ex);
                TaskDialog.Show("Error", $"Failed to open QR Code window:\n\n{ex.Message}");
            }
        }
    }
}