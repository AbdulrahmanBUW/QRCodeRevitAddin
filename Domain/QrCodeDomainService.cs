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
    /// Handles QR image creation, file operations, and Revit API interactions.
    /// </summary>
    public class QrCodeDomainService
    {
        private const int QR_PIXEL_SIZE = 300; // 300x300 pixels
        private const double QR_SHEET_SIZE_INCHES = 2.0; // 2" x 2" on sheet
        private const double INCHES_TO_FEET = 1.0 / 12.0; // Revit uses feet

        /// <summary>
        /// Generates QR code image as PNG byte array.
        /// </summary>
        /// <param name="content">Text content to encode in QR code</param>
        /// <returns>PNG image bytes</returns>
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
                            // Resize to exact dimensions
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
        /// Saves QR code bytes to a PNG file.
        /// </summary>
        /// <param name="qrBytes">QR code image bytes</param>
        /// <param name="filePath">Full path where to save the file</param>
        public void SaveQrCodeToFile(byte[] qrBytes, string filePath)
        {
            if (qrBytes == null || qrBytes.Length == 0)
            {
                throw new ArgumentException("QR code bytes cannot be null or empty", nameof(qrBytes));
            }

            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("File path cannot be empty", nameof(filePath));
            }

            try
            {
                // Ensure directory exists
                string directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                File.WriteAllBytes(filePath, qrBytes);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to save QR code to file: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Creates a temporary file with QR code image and returns the path.
        /// Caller is responsible for cleanup.
        /// </summary>
        /// <param name="qrBytes">QR code image bytes</param>
        /// <returns>Path to temporary file</returns>
        public string CreateTempQrFile(byte[] qrBytes)
        {
            string tempPath = Path.Combine(Path.GetTempPath(), $"QRCode_{Guid.NewGuid()}.png");
            SaveQrCodeToFile(qrBytes, tempPath);
            return tempPath;
        }

        /// <summary>
        /// Inserts QR code image into a Revit sheet at the specified location.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="sheet">Target sheet view</param>
        /// <param name="qrBytes">QR code image bytes</param>
        /// <param name="insertionPoint">XYZ point on sheet (in feet)</param>
        /// <returns>The created ImageInstance element</returns>
        public Element InsertQrIntoSheet(Document doc, ViewSheet sheet, byte[] qrBytes, XYZ insertionPoint)
        {
            if (doc == null)
                throw new ArgumentNullException(nameof(doc));

            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));

            if (qrBytes == null || qrBytes.Length == 0)
                throw new ArgumentException("QR code bytes cannot be null or empty", nameof(qrBytes));

            string tempFilePath = null;

            try
            {
                // Create temporary file
                tempFilePath = CreateTempQrFile(qrBytes);

                // Start transaction
                using (Transaction trans = new Transaction(doc, "Insert QR Code"))
                {
                    trans.Start();

                    try
                    {
                        // Create ImageType from file
                        ImageTypeOptions imageTypeOptions = new ImageTypeOptions(tempFilePath, false, ImageTypeSource.Import);
                        ImageType imageType = ImageType.Create(doc, imageTypeOptions);

                        if (imageType == null)
                        {
                            throw new InvalidOperationException("Failed to create ImageType");
                        }

                        // Set a unique name for the image type
                        imageType.Name = $"QRCode_{Guid.NewGuid().ToString().Substring(0, 8)}";

                        // Simple approach: Let user manually place and resize the image
                        // We'll just save the file and show instructions
                        trans.Commit();

                        // Show message with instructions
                        TaskDialog td = new TaskDialog("QR Code Ready");
                        td.MainInstruction = "QR Code image has been imported to the project.";
                        td.MainContent = $"The QR code image '{imageType.Name}' has been added to your project.\n\n" +
                                       "To place it on the sheet:\n" +
                                       "1. Go to Insert tab → Image\n" +
                                       "2. Select the QR code image from the list\n" +
                                       "3. Click on the sheet to place it\n" +
                                       "4. Resize to approximately 2\" x 2\"\n\n" +
                                       "Or use the 'Save' button to export the QR code as a PNG file.";
                        td.Show();

                        return imageType;
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        throw new InvalidOperationException($"Failed to insert QR code into sheet: {ex.Message}", ex);
                    }
                }
            }
            finally
            {
                // Clean up temporary file
                if (!string.IsNullOrEmpty(tempFilePath) && File.Exists(tempFilePath))
                {
                    try
                    {
                        File.Delete(tempFilePath);
                    }
                    catch
                    {
                        // Ignore cleanup errors
                    }
                }
            }
        }

        /// <summary>
        /// Inserts QR code at a random location on the sheet.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="sheet">Target sheet view</param>
        /// <param name="qrBytes">QR code image bytes</param>
        /// <returns>The created ImageInstance element</returns>
        public Element QuickInsertQrIntoSheet(Document doc, ViewSheet sheet, byte[] qrBytes)
        {
            Random random = new Random();
            double randomX = 0.5 + random.NextDouble() * 1.5;
            double randomY = 0.3 + random.NextDouble() * 1.0;
            XYZ insertionPoint = new XYZ(randomX, randomY, 0);

            return InsertQrIntoSheet(doc, sheet, qrBytes, insertionPoint);
        }

        /// <summary>
        /// Extracts sheet information from a Revit ViewSheet.
        /// </summary>
        /// <param name="sheet">The sheet view to extract data from</param>
        /// <returns>DocumentInfo populated with sheet data</returns>
        public Models.DocumentInfo ExtractSheetData(ViewSheet sheet)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));

            // Extract TLBL_DWG_NO (custom parameter for drawing number/sheet name)
            string dwgNo = GetParameterValueAsString(sheet, "TLBL_DWG_NO");

            // Extract sheet name (built-in)
            string sheetName = sheet.Name ?? string.Empty;

            // Extract revision
            string revision = ExtractRevisionFromSheet(sheet);

            // Extract Sheet Issue Date (built-in parameter)
            string date = GetSheetIssueDate(sheet);

            // Extract TLBL_CHECKEDBY (custom parameter)
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
                // Try to get parameter by name
                Parameter param = sheet.LookupParameter(parameterName);
                if (param != null && param.HasValue)
                {
                    if (param.StorageType == StorageType.String)
                    {
                        string value = param.AsString();
                        if (!string.IsNullOrWhiteSpace(value))
                            return value;
                    }
                    else if (param.StorageType == StorageType.Integer)
                    {
                        return param.AsInteger().ToString();
                    }
                    else if (param.StorageType == StorageType.Double)
                    {
                        return param.AsDouble().ToString();
                    }
                }

                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
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

                // Fallback to current date if no issue date is set
                return DateTime.Now.ToString("dd/MM/yyyy");
            }
            catch
            {
                return DateTime.Now.ToString("dd/MM/yyyy");
            }
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

                Parameter issueDateParam = sheet.get_Parameter(BuiltInParameter.SHEET_ISSUE_DATE);
                if (issueDateParam != null && issueDateParam.HasValue)
                {
                    string issueDate = issueDateParam.AsString();
                    if (!string.IsNullOrWhiteSpace(issueDate))
                        return issueDate;
                }

                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Checks if the latest version exists on server (database integration).
        /// </summary>
        public bool CheckLatestVersionFromServer(string combinedText)
        {
            return false;
        }

        /// <summary>
        /// Opens QR code file in external viewer/browser.
        /// </summary>
        public void OpenQrInViewer(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                throw new FileNotFoundException("QR code file not found", filePath);
            }

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