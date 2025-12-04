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

        public void SetInsertData(byte[] qrBytes, ViewSheet sheet, XYZ insertionPoint)
        {
            _qrBytes = qrBytes;
            _sheet = sheet;
            _insertionPoint = insertionPoint;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                if (_qrBytes == null || _sheet == null)
                {
                    return;
                }

                Document doc = app.ActiveUIDocument.Document;
                Domain.QrCodeDomainService service = new Domain.QrCodeDomainService();

                Element result = service.InsertQrIntoSheet(doc, _sheet, _qrBytes, _insertionPoint);

                if (result != null)
                {
                    app.ActiveUIDocument.Selection.SetElementIds(new ElementId[] { result.Id });
                }
            }
            catch (Exception)
            {
            }
        }

        public string GetName()
        {
            return "Insert QR Code Event";
        }
    }
}