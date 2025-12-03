using System;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using QRCodeRevitAddin.Domain;
using QRCodeRevitAddin.Models;
using Microsoft.Win32;

namespace QRCodeRevitAddin.ViewModels
{
    /// <summary>
    /// ViewModel for the QR Code window.
    /// Implements MVVM pattern with commands for UI interactions.
    /// Uses ExternalEvent for operations that require Revit API transactions.
    /// </summary>
    public class QrWindowViewModel : INotifyPropertyChanged
    {
        private readonly UIDocument _uiDoc;
        private readonly Document _doc;
        private readonly QrCodeDomainService _service;
        private readonly Commands.InsertQrEventHandler _insertEventHandler;
        private readonly ExternalEvent _insertEvent;
        private DocumentInfo _documentInfo;
        private BitmapImage _qrPreviewImage;
        private byte[] _currentQrBytes;
        private bool _isQrGenerated;
        private ViewSheet _currentSheet;

        #region Properties

        /// <summary>
        /// Document information model bound to UI fields.
        /// </summary>
        public DocumentInfo DocumentInfo
        {
            get => _documentInfo;
            set
            {
                if (_documentInfo != value)
                {
                    _documentInfo = value;
                    OnPropertyChanged(nameof(DocumentInfo));
                }
            }
        }

        /// <summary>
        /// QR code preview image displayed in UI.
        /// </summary>
        public BitmapImage QrPreviewImage
        {
            get => _qrPreviewImage;
            set
            {
                if (_qrPreviewImage != value)
                {
                    _qrPreviewImage = value;
                    OnPropertyChanged(nameof(QrPreviewImage));
                }
            }
        }

        /// <summary>
        /// Indicates whether a QR code has been generated.
        /// </summary>
        public bool IsQrGenerated
        {
            get => _isQrGenerated;
            set
            {
                if (_isQrGenerated != value)
                {
                    _isQrGenerated = value;
                    OnPropertyChanged(nameof(IsQrGenerated));
                    OnPropertyChanged(nameof(CanInsert));
                    OnPropertyChanged(nameof(CanSave));
                }
            }
        }

        /// <summary>
        /// Indicates whether insert operations are available.
        /// </summary>
        public bool CanInsert => IsQrGenerated && _currentSheet != null;

        /// <summary>
        /// Indicates whether save operation is available.
        /// </summary>
        public bool CanSave => IsQrGenerated;

        /// <summary>
        /// Indicates whether sheet data can be extracted.
        /// </summary>
        public bool CanUseSheetData => _currentSheet != null;

        #endregion

        #region Commands

        public ICommand GenerateQrCommand { get; }
        public ICommand SaveQrCommand { get; }
        public ICommand InsertQrCommand { get; }
        public ICommand QuickInsertCommand { get; }
        public ICommand UseSheetDataCommand { get; }
        public ICommand OpenInViewerCommand { get; }

        #endregion

        /// <summary>
        /// Constructor. Initializes the ViewModel with Revit context.
        /// </summary>
        /// <param name="uiDoc">The active UI document</param>
        /// <param name="autoFillFromSheet">If true, automatically fills data from current sheet</param>
        public QrWindowViewModel(UIDocument uiDoc, bool autoFillFromSheet = false)
        {
            _uiDoc = uiDoc ?? throw new ArgumentNullException(nameof(uiDoc));
            _doc = _uiDoc.Document;
            _service = new QrCodeDomainService();
            _documentInfo = new DocumentInfo();
            _isQrGenerated = false;

            // Create external event for insert operations
            _insertEventHandler = new Commands.InsertQrEventHandler();
            _insertEvent = ExternalEvent.Create(_insertEventHandler);

            // Try to get current sheet
            _currentSheet = _doc.ActiveView as ViewSheet;

            // Initialize commands
            GenerateQrCommand = new RelayCommand(GenerateQr, CanExecuteGenerateQr);
            SaveQrCommand = new RelayCommand(SaveQr, () => CanSave);
            InsertQrCommand = new RelayCommand(InsertQr, () => CanInsert);
            QuickInsertCommand = new RelayCommand(QuickInsert, () => CanInsert);
            UseSheetDataCommand = new RelayCommand(UseSheetData, () => CanUseSheetData);
            OpenInViewerCommand = new RelayCommand(OpenInViewer, () => CanSave);

            // Auto-fill from sheet if requested
            if (autoFillFromSheet && _currentSheet != null)
            {
                UseSheetData();
            }

            // Notify about current sheet availability
            OnPropertyChanged(nameof(CanUseSheetData));
            OnPropertyChanged(nameof(CanInsert));
        }

        #region Command Implementations

        private bool CanExecuteGenerateQr()
        {
            return _documentInfo != null && _documentInfo.IsValid();
        }

        private void GenerateQr()
        {
            try
            {
                if (!_documentInfo.IsValid())
                {
                    MessageBox.Show("Please fill in all fields before generating QR code.",
                        "Incomplete Data", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Generate QR code bytes
                _currentQrBytes = _service.GenerateQrCodeBytes(_documentInfo.CombinedText);

                // Convert to BitmapImage for preview
                QrPreviewImage = BytesToBitmapImage(_currentQrBytes);

                IsQrGenerated = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to generate QR code:\n\n{ex.Message}",
                    "Generation Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveQr()
        {
            try
            {
                if (_currentQrBytes == null)
                {
                    MessageBox.Show("Please generate a QR code first.",
                        "No QR Code", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Show save file dialog
                SaveFileDialog saveDialog = new SaveFileDialog
                {
                    Filter = "PNG Image|*.png",
                    Title = "Save QR Code",
                    FileName = $"QRCode_{DateTime.Now:yyyyMMdd_HHmmss}.png"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    // Create temp file first, then copy to destination
                    string tempPath = _service.CreateTempQrFile(_currentQrBytes);
                    File.Copy(tempPath, saveDialog.FileName, true);

                    // Clean up temp file
                    try { File.Delete(tempPath); } catch { }

                    MessageBox.Show($"QR code saved successfully to:\n{saveDialog.FileName}",
                        "Save Successful", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save QR code:\n\n{ex.Message}",
                    "Save Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void InsertQr()
        {
            try
            {
                if (_currentQrBytes == null)
                {
                    MessageBox.Show("Please generate a QR code first.",
                        "No QR Code", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (_currentSheet == null)
                {
                    MessageBox.Show("Please open a sheet view to insert the QR code.",
                        "No Sheet View", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Use a default location instead of prompting
                XYZ insertionPoint = new XYZ(1.0, 1.0, 0); // 1 foot from left, 1 foot from bottom

                // Set data for external event and raise it
                _insertEventHandler.SetInsertData(_currentQrBytes, _currentSheet, insertionPoint, false);
                _insertEvent.Raise();

                // Give a moment for the event to execute
                System.Threading.Thread.Sleep(100);

                // Show result
                if (_insertEventHandler.Success)
                {
                    MessageBox.Show("QR code inserted successfully onto the sheet!",
                        "Insert Successful", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show($"Insert operation failed.\n\n{_insertEventHandler.ResultMessage}",
                        "Insert Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to insert QR code:\n\n{ex.Message}",
                    "Insert Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void QuickInsert()
        {
            try
            {
                if (_currentQrBytes == null)
                {
                    MessageBox.Show("Please generate a QR code first.",
                        "No QR Code", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (_currentSheet == null)
                {
                    MessageBox.Show("Please open a sheet view to insert the QR code.",
                        "No Sheet View", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Set data for external event (quick insert will generate random location)
                _insertEventHandler.SetInsertData(_currentQrBytes, _currentSheet, XYZ.Zero, true);
                _insertEvent.Raise();

                // Give a moment for the event to execute
                System.Threading.Thread.Sleep(100);

                // Show result
                if (_insertEventHandler.Success)
                {
                    MessageBox.Show("QR code inserted successfully at a random location on the sheet.",
                        "Quick Insert Successful", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show($"Quick insert operation failed.\n\n{_insertEventHandler.ResultMessage}",
                        "Quick Insert Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to quick insert QR code:\n\n{ex.Message}",
                    "Quick Insert Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UseSheetData()
        {
            try
            {
                if (_currentSheet == null)
                {
                    MessageBox.Show("Please select a sheet view first.\n\nOpen a sheet in Revit, then click this button.",
                        "No Sheet Selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Extract data from sheet
                DocumentInfo = _service.ExtractSheetData(_currentSheet);

                MessageBox.Show("Sheet data loaded successfully.\n\nClick 'Generate QR' to create the QR code.",
                    "Data Loaded", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to extract sheet data:\n\n{ex.Message}",
                    "Extraction Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenInViewer()
        {
            try
            {
                if (_currentQrBytes == null)
                {
                    MessageBox.Show("Please generate a QR code first.",
                        "No QR Code", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Save to temp file and open
                string tempPath = _service.CreateTempQrFile(_currentQrBytes);
                _service.OpenQrInViewer(tempPath);

                // Note: temp file will be cleaned up by OS or when Revit closes
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to open QR code in viewer:\n\n{ex.Message}",
                    "Viewer Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Helper Methods

        /// <summary>
        /// Prompts user to pick a point on the sheet for QR insertion.
        /// </summary>
        /// <returns>XYZ point in sheet coordinates, or null if cancelled</returns>
        private XYZ PromptForInsertionPoint()
        {
            try
            {
                // Prompt user to pick point on sheet
                TaskDialog.Show("Pick Insertion Point",
                    "Click on the sheet to select where to place the QR code.\n\n" +
                    "The QR code will be placed at the clicked point.");

                Selection selection = _uiDoc.Selection;
                XYZ pickedPoint = selection.PickPoint(ObjectSnapTypes.None, "Click to place QR code");

                return pickedPoint;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // User cancelled
                return null;
            }
            catch (Exception)
            {
                // If pick point fails, use a default location
                return new XYZ(1.0, 1.0, 0);
            }
        }

        /// <summary>
        /// Converts byte array to BitmapImage for WPF display.
        /// </summary>
        /// <param name="bytes">Image bytes</param>
        /// <returns>BitmapImage for WPF</returns>
        private BitmapImage BytesToBitmapImage(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
                return null;

            BitmapImage image = new BitmapImage();
            using (MemoryStream stream = new MemoryStream(bytes))
            {
                stream.Position = 0;
                image.BeginInit();
                image.CreateOptions = BitmapCreateOptions.PreservePixelFormat;
                image.CacheOption = BitmapCacheOption.OnLoad;
                image.StreamSource = stream;
                image.EndInit();
                image.Freeze();
            }
            return image;
        }

        #endregion

        #region INotifyPropertyChanged Implementation

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion
    }

    /// <summary>
    /// Simple RelayCommand implementation for MVVM commanding.
    /// </summary>
    public class RelayCommand : ICommand
    {
        private readonly Action _execute;
        private readonly Func<bool> _canExecute;

        public RelayCommand(Action execute, Func<bool> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object parameter)
        {
            return _canExecute == null || _canExecute();
        }

        public void Execute(object parameter)
        {
            _execute();
        }
    }
}