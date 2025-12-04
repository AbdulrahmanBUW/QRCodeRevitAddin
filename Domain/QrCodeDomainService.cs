using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using QRCoder;

namespace QRCodeRevitAddin.Domain
{
    public class QrCodeDomainService
    {
        private const int QR_PIXEL_SIZE = 300;

        public byte[] GenerateQrCodeBytes(string content)
        {
            if (string.IsNullOrWhiteSpace(content))
            {
                throw new ArgumentException("QR code content cannot be empty", nameof(content));
            }

            try
            {
                using (QRCodeGenerator qrGenerator = new QRCodeGenerator())
                {
                    QRCodeData qrCodeData = qrGenerator.CreateQrCode(content, QRCodeGenerator.ECCLevel.Q);
                    using (QRCode qrCode = new QRCode(qrCodeData))
                    {
                        using (Bitmap qrBitmap = qrCode.GetGraphic(
                            pixelsPerModule: 20,
                            darkColor: System.Drawing.Color.Black,
                            lightColor: System.Drawing.Color.White,
                            drawQuietZones: true))
                        {
                            using (Bitmap resizedBitmap = new Bitmap(qrBitmap, QR_PIXEL_SIZE, QR_PIXEL_SIZE))
                            {
                                using (MemoryStream ms = new MemoryStream())
                                {
                                    resizedBitmap.Save(ms, ImageFormat.Png);
                                    return ms.ToArray();
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to generate QR code: {ex.Message}", ex);
            }
        }

        public string CreateTempQrFile(byte[] qrBytes)
        {
            string tempPath = Path.Combine(Path.GetTempPath(), $"QRCode_{Guid.NewGuid()}.png");

            // Ensure directory exists
            string directory = Path.GetDirectoryName(tempPath);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllBytes(tempPath, qrBytes);
            return tempPath;
        }

        public Element InsertQrIntoSheet(Document doc, ViewSheet sheet, byte[] qrBytes, XYZ insertionPoint)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (qrBytes == null || qrBytes.Length == 0)
                throw new ArgumentException("QR code bytes cannot be null or empty", nameof(qrBytes));

            string tempFilePath = null;

            try
            {
                tempFilePath = CreateTempQrFile(qrBytes);

                using (Transaction trans = new Transaction(doc, "Insert QR Code"))
                {
                    trans.Start();

                    try
                    {
                        ImageTypeOptions imageTypeOptions = new ImageTypeOptions(tempFilePath, false, ImageTypeSource.Import);
                        ImageType imageType = ImageType.Create(doc, imageTypeOptions);

                        if (imageType == null)
                            throw new InvalidOperationException("Failed to create ImageType");

                        imageType.Name = $"QRCode_{DateTime.Now:yyyyMMdd_HHmmss}";

                        ImagePlacementOptions placementOptions = new ImagePlacementOptions();

                        ImageInstance imageInstance = ImageInstance.Create(doc, sheet, imageType.Id, placementOptions);

                        if (imageInstance != null)
                        {
                            SetImageSize(imageInstance, 2.0);

                            trans.Commit();
                            return imageInstance;
                        }

                        throw new InvalidOperationException("Could not create image instance");
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        throw new InvalidOperationException($"Failed to insert QR: {ex.Message}", ex);
                    }
                }
            }
            finally
            {
                if (!string.IsNullOrEmpty(tempFilePath) && File.Exists(tempFilePath))
                {
                    try { File.Delete(tempFilePath); } catch { }
                }
            }
        }

        private void SetImageSize(ImageInstance imageInstance, double sizeInInches)
        {
            try
            {
                double sizeInFeet = sizeInInches / 12.0;

                Parameter widthParam = imageInstance.LookupParameter("Width");
                if (widthParam == null)
                    widthParam = imageInstance.LookupParameter("Image Width");

                if (widthParam != null && !widthParam.IsReadOnly)
                {
                    widthParam.Set(sizeInFeet);
                }

                Parameter heightParam = imageInstance.LookupParameter("Height");
                if (heightParam == null)
                    heightParam = imageInstance.LookupParameter("Image Height");

                if (heightParam != null && !heightParam.IsReadOnly)
                {
                    heightParam.Set(sizeInFeet);
                }
            }
            catch
            {
                // The User can resize manually
            }
        }

        public Models.DocumentInfo ExtractSheetData(ViewSheet sheet)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));

            string dwgNo = GetParameterValueAsString(sheet, "TLBL_DWG_NO");
            string sheetName = sheet.Name ?? string.Empty;
            string revision = ExtractRevisionFromSheet(sheet);
            string date = GetSheetIssueDate(sheet);
            string checkedBy = GetParameterValueAsString(sheet, "TLBL_CHECKEDBY");

            return new Models.DocumentInfo(dwgNo, sheetName, revision, date, checkedBy);
        }

        private string GetParameterValueAsString(ViewSheet sheet, string parameterName)
        {
            try
            {
                Parameter param = sheet.LookupParameter(parameterName);
                if (param != null && param.HasValue)
                {
                    if (param.StorageType == StorageType.String)
                        return param.AsString() ?? "";
                    else if (param.StorageType == StorageType.Integer)
                        return param.AsInteger().ToString();
                    else if (param.StorageType == StorageType.Double)
                        return param.AsDouble().ToString();
                }
                return "";
            }
            catch { return ""; }
        }

        private string GetSheetIssueDate(ViewSheet sheet)
        {
            try
            {
                Parameter issueDateParam = sheet.get_Parameter(BuiltInParameter.SHEET_ISSUE_DATE);
                if (issueDateParam != null && issueDateParam.HasValue)
                {
                    string issueDate = issueDateParam.AsString();
                    if (!string.IsNullOrWhiteSpace(issueDate))
                        return issueDate;
                }
                return DateTime.Now.ToString("dd/MM/yyyy");
            }
            catch { return DateTime.Now.ToString("dd/MM/yyyy"); }
        }

        private string ExtractRevisionFromSheet(ViewSheet sheet)
        {
            try
            {
                Parameter currentRevParam = sheet.get_Parameter(BuiltInParameter.SHEET_CURRENT_REVISION);
                if (currentRevParam != null && currentRevParam.HasValue)
                {
                    string revValue = currentRevParam.AsString();
                    if (!string.IsNullOrWhiteSpace(revValue))
                        return revValue;
                }

                return "";
            }
            catch { return ""; }
        }

        public void OpenQrInViewer(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException("QR code file not found", filePath);

            try
            {
                System.Diagnostics.Process.Start(filePath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to open QR code in viewer: {ex.Message}", ex);
            }
        }
    }
}