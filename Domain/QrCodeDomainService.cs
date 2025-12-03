using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using QRCoder;

namespace QRCodeRevitAddin.Domain
{
    /// <summary>
    /// Domain service for QR code generation and insertion into Revit documents.
    /// </summary>
    public class QrCodeDomainService
    {
        private const int QR_PIXEL_SIZE = 300;

        /// <summary>
        /// Generates QR code image as PNG byte array.
        /// </summary>
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

        /// <summary>
        /// Creates a temporary file with QR code image and returns the path.
        /// </summary>
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

        /// <summary>
        /// Simple method that actually works - inserts QR code onto sheet
        /// </summary>
        public Element InsertQrIntoSheet(Document doc, ViewSheet sheet, byte[] qrBytes, XYZ insertionPoint)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (qrBytes == null || qrBytes.Length == 0)
                throw new ArgumentException("QR code bytes cannot be null or empty", nameof(qrBytes));

            string tempFilePath = null;

            try
            {
                // Save QR code to temp file
                tempFilePath = CreateTempQrFile(qrBytes);

                using (Transaction trans = new Transaction(doc, "Insert QR Code"))
                {
                    trans.Start();

                    try
                    {
                        // Import the image
                        ImageTypeOptions imageTypeOptions = new ImageTypeOptions(tempFilePath, false, ImageTypeSource.Import);
                        ImageType imageType = ImageType.Create(doc, imageTypeOptions);

                        if (imageType == null)
                            throw new InvalidOperationException("Failed to create ImageType");

                        imageType.Name = $"QRCode_{DateTime.Now:yyyyMMdd_HHmmss}";

                        // SIMPLEST APPROACH: Let Revit handle the placement
                        // Create ImagePlacementOptions with default constructor
                        ImagePlacementOptions placementOptions = new ImagePlacementOptions();

                        // Create the image instance on the sheet
                        ImageInstance imageInstance = ImageInstance.Create(doc, sheet, imageType.Id, placementOptions);

                        if (imageInstance != null)
                        {
                            // Success! QR code is on the sheet
                            SetImageSize(imageInstance, 2.0);
                            trans.Commit();

                            TaskDialog.Show("Success", "QR code inserted onto sheet!");
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

        /// <summary>
        /// Sets the image size in inches.
        /// </summary>
        private void SetImageSize(ImageInstance imageInstance, double sizeInInches)
        {
            try
            {
                double sizeInFeet = sizeInInches / 12.0;

                // Try to set width
                Parameter widthParam = imageInstance.LookupParameter("Width");
                if (widthParam == null)
                    widthParam = imageInstance.LookupParameter("Image Width");

                if (widthParam != null && !widthParam.IsReadOnly)
                {
                    widthParam.Set(sizeInFeet);
                }

                // Try to set height
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
                // If we can't set the size, that's OK - user can resize manually
            }
        }

        /// <summary>
        /// Quick insert at a reasonable location on the sheet.
        /// </summary>
        public Element QuickInsertQrIntoSheet(Document doc, ViewSheet sheet, byte[] qrBytes)
        {
            // Get sheet boundaries
            BoundingBoxUV outline = sheet.Outline;

            // Place in the upper left quadrant (avoiding title block)
            double x = outline.Min.U + (outline.Max.U - outline.Min.U) * 0.25;
            double y = outline.Min.V + (outline.Max.V - outline.Min.V) * 0.75;

            XYZ insertionPoint = new XYZ(x, y, 0);

            return InsertQrIntoSheet(doc, sheet, qrBytes, insertionPoint);
        }

        /// <summary>
        /// Extracts sheet information from a Revit ViewSheet.
        /// </summary>
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

        /// <summary>
        /// Gets parameter value as string by parameter name.
        /// </summary>
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

        /// <summary>
        /// Gets Sheet Issue Date from built-in parameter.
        /// </summary>
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

        /// <summary>
        /// Attempts to extract revision information from a sheet.
        /// </summary>
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

        /// <summary>
        /// Opens QR code file in external viewer/browser.
        /// </summary>
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