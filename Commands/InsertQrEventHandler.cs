using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using QRCodeRevitAddin.Utils;

namespace QRCodeRevitAddin.Commands
{
    public class InsertQrEventHandler : IExternalEventHandler
    {
        private byte[] _qrBytes;
        private ViewSheet _sheet;
        private XYZ _insertionPoint;

        public void SetInsertData(byte[] qrBytes, ViewSheet sheet, XYZ insertionPoint)
        {
            _qrBytes = qrBytes;
            _sheet = sheet;
            _insertionPoint = insertionPoint;
            Logger.LogInfo("Insert data set for QR code insertion");
        }

        public void Execute(UIApplication app)
        {
            try
            {
                if (_qrBytes == null || _sheet == null)
                {
                    Logger.LogWarning("Insert event called with null data");
                    return;
                }

                Logger.LogInfo("Executing QR code insertion");

                Document doc = app.ActiveUIDocument.Document;
                Domain.QrCodeDomainService service = new Domain.QrCodeDomainService();

                Element result = service.InsertQrIntoSheet(doc, _sheet, _qrBytes, _insertionPoint);

                if (result != null)
                {
                    app.ActiveUIDocument.Selection.SetElementIds(new ElementId[] { result.Id });
                    Logger.LogInfo("QR code inserted and selected");
                }
            }
            catch (Exception ex)
            {
                Logger.LogError("QR insertion event failed", ex);
                TaskDialog.Show("Insert Error", $"Failed to insert QR code:\n\n{ex.Message}");
            }
        }

        public string GetName()
        {
            return "Insert QR Code Event";
        }
    }
}