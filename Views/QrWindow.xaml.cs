using System.Windows;
using Autodesk.Revit.UI;
using QRCodeRevitAddin.ViewModels;

namespace QRCodeRevitAddin.Views
{
    public partial class QrWindow : Window
    {
        public QrWindow(QrWindowViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }

        public static void Show(UIDocument uiDoc, bool autoFillFromSheet = false)
        {
            if (uiDoc == null)
            {
                TaskDialog.Show("Error", "No active Revit document found.");
                return;
            }

            try
            {
                QrWindowViewModel viewModel = new QrWindowViewModel(uiDoc, autoFillFromSheet);
                QrWindow window = new QrWindow(viewModel);
                window.Show();
            }
            catch (System.Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to open QR Code window:\n\n{ex.Message}");
            }
        }
    }
}