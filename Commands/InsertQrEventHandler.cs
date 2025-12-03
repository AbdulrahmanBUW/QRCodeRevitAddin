using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace QRCodeRevitAddin.Commands
{
    public class InsertQrEventHandler : IExternalEventHandler
    {
        private byte[] _qrBytes;
        private ViewSheet _sheet;
        private XYZ _insertionPoint;
        private bool _isQuickInsert;
        private string _resultMessage;
        private bool _success;

        public string ResultMessage => _resultMessage;
        public bool Success => _success;

        public void SetInsertData(byte[] qrBytes, ViewSheet sheet, XYZ insertionPoint, bool isQuickInsert)
        {
            _qrBytes = qrBytes;
            _sheet = sheet;
            _insertionPoint = insertionPoint;
            _isQuickInsert = isQuickInsert;
            _success = false;
            _resultMessage = string.Empty;
        }

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
                    _resultMessage = "QR code inserted successfully onto the sheet!";

                    // Select the inserted QR code
                    app.ActiveUIDocument.Selection.SetElementIds(new ElementId[] { result.Id });
                }
                else
                {
                    _success = false;
                    _resultMessage = "Failed to insert image element";
                }
            }
            catch (Exception ex)
            {
                _success = false;
                _resultMessage = $"Insert failed: {ex.Message}";
            }
        }

        public string GetName()
        {
            return "Insert QR Code Event";
        }
    }
}