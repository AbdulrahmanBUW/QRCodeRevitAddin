using System;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using QRCodeRevitAddin.Domain;
using QRCodeRevitAddin.Models;
using QRCodeRevitAddin.Utils;
using Microsoft.Win32;

namespace QRCodeRevitAddin.ViewModels
{
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

        public bool CanInsert => IsQrGenerated && _currentSheet != null;
        public bool CanSave => IsQrGenerated;
        public bool CanUseSheetData => _currentSheet != null;

        #endregion

        #region Commands

        public ICommand GenerateQrCommand { get; }
        public ICommand SaveQrCommand { get; }
        public ICommand InsertQrCommand { get; }
        public ICommand UseSheetDataCommand { get; }
        public ICommand OpenInViewerCommand { get; }

        #endregion

        public QrWindowViewModel(UIDocument uiDoc, bool autoFillFromSheet = false)
        {
            try
            {
                Logger.LogInfo("Initializing QR Window ViewModel");

                _uiDoc = uiDoc ?? throw new ArgumentNullException(nameof(uiDoc));
                _doc = _uiDoc.Document;
                _service = new QrCodeDomainService();
                _documentInfo = new DocumentInfo();
                _isQrGenerated = false;

                _insertEventHandler = new Commands.InsertQrEventHandler();
                _insertEvent = ExternalEvent.Create(_insertEventHandler);

                _currentSheet = _doc.ActiveView as ViewSheet;

                GenerateQrCommand = new RelayCommand(GenerateQr, CanExecuteGenerateQr);
                SaveQrCommand = new RelayCommand(SaveQr, () => CanSave);
                InsertQrCommand = new RelayCommand(InsertQr, () => CanInsert);
                UseSheetDataCommand = new RelayCommand(UseSheetData, () => CanUseSheetData);
                OpenInViewerCommand = new RelayCommand(OpenInViewer, () => CanSave);

                if (autoFillFromSheet && _currentSheet != null)
                {
                    UseSheetData();
                }

                OnPropertyChanged(nameof(CanUseSheetData));
                OnPropertyChanged(nameof(CanInsert));

                Logger.LogInfo("QR Window ViewModel initialized successfully");
            }
            catch (Exception ex)
            {
                Logger.LogError("Failed to initialize QR Window ViewModel", ex);
                throw;
            }
        }

        #region Command Implementations

        private bool CanExecuteGenerateQr()
        {
            if (_documentInfo == null)
                return false;

            string errorMessage;
            return _documentInfo.IsValid(out errorMessage);
        }

        private void GenerateQr()
        {
            try
            {
                Logger.LogInfo("Generate QR command executed");

                string errorMessage;
                if (!_documentInfo.IsValid(out errorMessage))
                {
                    Logger.LogWarning($"Validation failed: {errorMessage}");
                    MessageBox.Show(errorMessage, "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                _currentQrBytes = _service.GenerateQrCodeBytes(_documentInfo.CombinedText);

                QrPreviewImage = BytesToBitmapImage(_currentQrBytes);

                IsQrGenerated = true;

                Logger.LogInfo("QR code generated and preview updated");
            }
            catch (Exception ex)
            {
                Logger.LogError("Failed to generate QR code", ex);
                MessageBox.Show($"Failed to generate QR code:\n\n{ex.Message}", "Generation Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveQr()
        {
            try
            {
                Logger.LogInfo("Save QR command executed");

                if (_currentQrBytes == null)
                {
                    Logger.LogWarning("No QR code to save");
                    MessageBox.Show("Please generate a QR code first.", "No QR Code", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                SaveFileDialog saveDialog = new SaveFileDialog
                {
                    Filter = "PNG Image|*.png",
                    Title = "Save QR Code",
                    FileName = $"QRCode_{DateTime.Now:yyyyMMdd_HHmmss}.png"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    string tempPath = _service.CreateTempQrFile(_currentQrBytes);
                    File.Copy(tempPath, saveDialog.FileName, true);

                    try { File.Delete(tempPath); } catch { }

                    Logger.LogInfo($"QR code saved to: {saveDialog.FileName}");
                    MessageBox.Show($"QR code saved successfully to:\n{saveDialog.FileName}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                Logger.LogError("Failed to save QR code", ex);
                MessageBox.Show($"Failed to save QR code:\n\n{ex.Message}", "Save Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void InsertQr()
        {
            try
            {
                Logger.LogInfo("Insert QR command executed");

                if (_currentQrBytes == null)
                {
                    Logger.LogWarning("No QR code to insert");
                    MessageBox.Show("Please generate a QR code first.", "No QR Code", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (_currentSheet == null)
                {
                    Logger.LogWarning("No sheet view active");
                    MessageBox.Show("Please open a sheet view to insert the QR code.", "No Sheet View", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                BoundingBoxUV outline = _currentSheet.Outline;
                double x = outline.Min.U + (outline.Max.U - outline.Min.U) * 0.25;
                double y = outline.Min.V + (outline.Max.V - outline.Min.V) * 0.75;
                XYZ insertionPoint = new XYZ(x, y, 0);

                _insertEventHandler.SetInsertData(_currentQrBytes, _currentSheet, insertionPoint);
                _insertEvent.Raise();

                Logger.LogInfo("QR code insertion triggered");
            }
            catch (Exception ex)
            {
                Logger.LogError("Failed to insert QR code", ex);
                MessageBox.Show($"Failed to insert QR code:\n\n{ex.Message}", "Insert Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UseSheetData()
        {
            try
            {
                Logger.LogInfo("Use Sheet Data command executed");

                if (_currentSheet == null)
                {
                    Logger.LogWarning("No sheet selected");
                    MessageBox.Show("Please select a sheet view first.\n\nOpen a sheet in Revit, then click this button.", "No Sheet Selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                DocumentInfo = _service.ExtractSheetData(_currentSheet);
                Logger.LogInfo("Sheet data extracted and populated");
            }
            catch (Exception ex)
            {
                Logger.LogError("Failed to extract sheet data", ex);
                MessageBox.Show($"Failed to extract sheet data:\n\n{ex.Message}", "Extraction Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenInViewer()
        {
            try
            {
                Logger.LogInfo("Open in Viewer command executed");

                if (_currentQrBytes == null)
                {
                    Logger.LogWarning("No QR code to view");
                    MessageBox.Show("Please generate a QR code first.", "No QR Code", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string tempPath = _service.CreateTempQrFile(_currentQrBytes);
                _service.OpenQrInViewer(tempPath);

                Logger.LogInfo("QR code opened in viewer");
            }
            catch (Exception ex)
            {
                Logger.LogError("Failed to open QR code in viewer", ex);
                MessageBox.Show($"Failed to open QR code in viewer:\n\n{ex.Message}", "Viewer Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Helper Methods

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