using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace QRCodeRevitAddin.Commands
{
    /// <summary>
    /// External event handler for inserting QR codes from modeless dialogs.
    /// Executes the insert operation within Revit's API context.
    /// </summary>
    public class InsertQrEventHandler : IExternalEventHandler
    {
        private byte[] _qrBytes;
        private ViewSheet _sheet;
        private XYZ _insertionPoint;
        private bool _isQuickInsert;
        private string _resultMessage;
        private bool _success;

        /// <summary>
        /// Gets the result message after execution.
        /// </summary>
        public string ResultMessage => _resultMessage;

        /// <summary>
        /// Gets whether the operation was successful.
        /// </summary>
        public bool Success => _success;

        /// <summary>
        /// Sets the data for the next insert operation.
        /// </summary>
        public void SetInsertData(byte[] qrBytes, ViewSheet sheet, XYZ insertionPoint, bool isQuickInsert)
        {
            _qrBytes = qrBytes;
            _sheet = sheet;
            _insertionPoint = insertionPoint;
            _isQuickInsert = isQuickInsert;
            _success = false;
            _resultMessage = string.Empty;
        }

        /// <summary>
        /// Executes the insert operation within Revit's API context.
        /// </summary>
        public void Execute(UIApplication app)
        {
            try
            {
                if (_qrBytes == null || _sheet == null)
                {
                    _resultMessage = "Invalid insert data";
                    _success = false;
                    return;
                }

                Document doc = app.ActiveUIDocument.Document;
                Domain.QrCodeDomainService service = new Domain.QrCodeDomainService();

                Element result;
                if (_isQuickInsert)
                {
                    result = service.QuickInsertQrIntoSheet(doc, _sheet, _qrBytes);
                }
                else
                {
                    result = service.InsertQrIntoSheet(doc, _sheet, _qrBytes, _insertionPoint);
                }

                if (result != null)
                {
                    _success = true;
                    _resultMessage = "QR code inserted successfully";
                }
                else
                {
                    _success = false;
                    _resultMessage = "Failed to create image element";
                }
            }
            catch (Exception ex)
            {
                _success = false;
                _resultMessage = $"Insert failed: {ex.Message}";
            }
        }

        /// <summary>
        /// Returns the name of this external event handler.
        /// </summary>
        public string GetName()
        {
            return "Insert QR Code Event";
        }
    }
}