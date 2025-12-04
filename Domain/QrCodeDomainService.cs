using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using QRCoder;
using QRCodeRevitAddin.Utils;

namespace QRCodeRevitAddin.Domain
{
    public class QrCodeDomainService
    {
        private const int QR_PIXEL_SIZE = 300;

        public byte[] GenerateQrCodeBytes(string content)
        {
            if (string.IsNullOrWhiteSpace(content))
            {
                Logger.LogError("QR code content cannot be empty");
                throw new ArgumentException("QR code content cannot be empty", nameof(content));
            }

            try
            {
                Logger.LogInfo("Generating QR code");

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
                                    byte[] result = ms.ToArray();
                                    Logger.LogInfo($"QR code generated successfully, size: {result.Length} bytes");
                                    return result;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogError("Failed to generate QR code", ex);
                throw new InvalidOperationException($"Failed to generate QR code: {ex.Message}", ex);
            }
        }

        public string CreateTempQrFile(byte[] qrBytes)
        {
            try
            {
                string tempPath = Path.Combine(Path.GetTempPath(), $"QRCode_{Guid.NewGuid()}.png");

                string directory = Path.GetDirectoryName(tempPath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                File.WriteAllBytes(tempPath, qrBytes);
                Logger.LogInfo($"Temporary QR file created: {tempPath}");
                return tempPath;
            }
            catch (Exception ex)
            {
                Logger.LogError("Failed to create temporary QR file", ex);
                throw;
            }
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
                Logger.LogInfo($"Inserting QR code into sheet: {sheet.Name}");

                tempFilePath = CreateTempQrFile(qrBytes);

                using (Transaction trans = new Transaction(doc, "Insert QR Code"))
                {
                    trans.Start();

                    try
                    {
                        ImageTypeOptions imageTypeOptions = new ImageTypeOptions(tempFilePath, false, ImageTypeSource.Import);
                        ImageType imageType = ImageType.Create(doc, imageTypeOptions);

                        if (imageType == null)
                        {
                            Logger.LogError("Failed to create ImageType");
                            throw new InvalidOperationException("Failed to create ImageType");
                        }

                        imageType.Name = $"QRCode_{DateTime.Now:yyyyMMdd_HHmmss}";

                        ImagePlacementOptions placementOptions = new ImagePlacementOptions();

                        ImageInstance imageInstance = ImageInstance.Create(doc, sheet, imageType.Id, placementOptions);

                        if (imageInstance != null)
                        {
                            SetImageSize(imageInstance, 2.0);
                            trans.Commit();
                            Logger.LogInfo($"QR code inserted successfully, ID: {imageInstance.Id}");
                            return imageInstance;
                        }

                        Logger.LogError("Could not create image instance");
                        throw new InvalidOperationException("Could not create image instance");
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        Logger.LogError("Failed to insert QR code, transaction rolled back", ex);
                        throw new InvalidOperationException($"Failed to insert QR: {ex.Message}", ex);
                    }
                }
            }
            finally
            {
                if (!string.IsNullOrEmpty(tempFilePath) && File.Exists(tempFilePath))
                {
                    try
                    {
                        File.Delete(tempFilePath);
                        Logger.LogInfo("Temporary file cleaned up");
                    }
                    catch (Exception ex)
                    {
                        Logger.LogWarning($"Failed to delete temporary file: {ex.Message}");
                    }
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

                Logger.LogInfo($"Image size set to {sizeInInches} inches");
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"Could not set image size: {ex.Message}");
            }
        }

        public Models.DocumentInfo ExtractSheetData(ViewSheet sheet)
        {
            if (sheet == null)
            {
                Logger.LogError("Sheet is null");
                throw new ArgumentNullException(nameof(sheet));
            }

            try
            {
                Logger.LogInfo($"Extracting data from sheet: {sheet.Name}");

                string dwgNo = GetParameterValueAsString(sheet, "TLBL_DWG_NO");
                string sheetName = sheet.Name ?? string.Empty;
                string revision = ExtractRevisionFromSheet(sheet);
                string date = GetSheetIssueDate(sheet);
                string checkedBy = GetParameterValueAsString(sheet, "TLBL_CHECKEDBY");

                Logger.LogInfo("Sheet data extracted successfully");
                return new Models.DocumentInfo(dwgNo, sheetName, revision, date, checkedBy);
            }
            catch (Exception ex)
            {
                Logger.LogError("Failed to extract sheet data", ex);
                throw;
            }
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
            catch (Exception ex)
            {
                Logger.LogWarning($"Failed to get parameter {parameterName}: {ex.Message}");
                return "";
            }
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
                return DateValidator.GetTodayFormatted();
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"Failed to get sheet issue date: {ex.Message}");
                return DateValidator.GetTodayFormatted();
            }
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
            catch (Exception ex)
            {
                Logger.LogWarning($"Failed to extract revision: {ex.Message}");
                return "";
            }
        }

        public void OpenQrInViewer(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                Logger.LogError($"QR code file not found: {filePath}");
                throw new FileNotFoundException("QR code file not found", filePath);
            }

            try
            {
                Logger.LogInfo($"Opening QR code in viewer: {filePath}");
                System.Diagnostics.Process.Start(filePath);
            }
            catch (Exception ex)
            {
                Logger.LogError("Failed to open QR code in viewer", ex);
                throw new InvalidOperationException($"Failed to open QR code in viewer: {ex.Message}", ex);
            }
        }
    }
}