using System.Windows;
using Autodesk.Revit.UI;
using QRCodeRevitAddin.ViewModels;

namespace QRCodeRevitAddin.Views
{
    /// <summary>
    /// Code-behind for QrWindow.xaml.
    /// Interaction logic for the QR Code Generator window.
    /// </summary>
    public partial class QrWindow : Window
    {
        /// <summary>
        /// Constructor that takes a ViewModel.
        /// </summary>
        /// <param name="viewModel">The ViewModel for this window</param>
        public QrWindow(QrWindowViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }

        /// <summary>
        /// Static method to show the window with proper Revit context.
        /// </summary>
        /// <param name="uiDoc">The active UI document</param>
        /// <param name="autoFillFromSheet">If true, automatically fills data from current sheet</param>
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
                window.Show(); // Changed from ShowDialog() to Show() - now modeless!
            }
            catch (System.Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to open QR Code window:\n\n{ex.Message}");
            }
        }
    }
}