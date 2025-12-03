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

                        // Calculate size in feet (2 inches = 2/12 feet)
                        double sizeInFeet = QR_SHEET_SIZE_INCHES * INCHES_TO_FEET;

                        // For Revit 2023, ImagePlacementOptions constructor takes the four corner points
                        // Constructor: ImagePlacementOptions(XYZ lowerLeft, XYZ lowerRight, XYZ upperLeft, XYZ upperRight)

                        // Define the placement rectangle corners
                        XYZ lowerLeft = insertionPoint;
                        XYZ lowerRight = new XYZ(insertionPoint.X + sizeInFeet, insertionPoint.Y, insertionPoint.Z);
                        XYZ upperLeft = new XYZ(insertionPoint.X, insertionPoint.Y + sizeInFeet, insertionPoint.Z);
                        XYZ upperRight = new XYZ(insertionPoint.X + sizeInFeet, insertionPoint.Y + sizeInFeet, insertionPoint.Z);

                        // Create ImagePlacementOptions with the four corner points
                        ImagePlacementOptions placementOptions = new ImagePlacementOptions(lowerLeft, lowerRight, upperLeft, upperRight);

                        // Create the image instance on the sheet
                        ImageInstance imageInstance = ImageInstance.Create(doc, sheet, imageType.Id, placementOptions);

                        if (imageInstance == null)
                        {
                            throw new InvalidOperationException("Failed to create ImageInstance");
                        }

                        trans.Commit();

                        return imageInstance;
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
        /// Used for quick insert functionality.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="sheet">Target sheet view</param>
        /// <param name="qrBytes">QR code image bytes</param>
        /// <returns>The created ImageInstance element</returns>
        public Element QuickInsertQrIntoSheet(Document doc, ViewSheet sheet, byte[] qrBytes)
        {
            // Generate random position on sheet
            // Sheets typically have origin at (0,0) and extend to positive X,Y
            // Place QR in a random location, avoiding edges
            Random random = new Random();

            // Typical sheet size is around 3 feet x 2 feet (36" x 24")
            // Place QR somewhere in the middle third of the sheet
            double randomX = 0.5 + random.NextDouble() * 1.5; // 0.5 to 2.0 feet
            double randomY = 0.3 + random.NextDouble() * 1.0; // 0.3 to 1.3 feet

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

            string sheetNumber = sheet.SheetNumber ?? string.Empty;
            string sheetName = sheet.Name ?? string.Empty;
            string revision = ExtractRevisionFromSheet(sheet);
            string date = DateTime.Now.ToString("dd/MM/yyyy");

            return new Models.DocumentInfo(sheetNumber, sheetName, revision, date);
        }

        /// <summary>
        /// Attempts to extract revision information from a sheet.
        /// Tries multiple methods: Current Revision parameter, revision clouds, and schedules.
        /// </summary>
        /// <param name="sheet">The sheet to extract revision from</param>
        /// <returns>Revision string, or empty if not found</returns>
        private string ExtractRevisionFromSheet(ViewSheet sheet)
        {
            try
            {
                // Method 1: Try built-in "Current Revision" parameter
                Parameter currentRevParam = sheet.get_Parameter(BuiltInParameter.SHEET_CURRENT_REVISION);
                if (currentRevParam != null && currentRevParam.HasValue)
                {
                    string revValue = currentRevParam.AsString();
                    if (!string.IsNullOrWhiteSpace(revValue))
                        return revValue;
                }

                // Method 2: Try "Sheet Issue Date" parameter
                Parameter issueDateParam = sheet.get_Parameter(BuiltInParameter.SHEET_ISSUE_DATE);
                if (issueDateParam != null && issueDateParam.HasValue)
                {
                    string issueDate = issueDateParam.AsString();
                    if (!string.IsNullOrWhiteSpace(issueDate))
                        return issueDate;
                }

                // Method 3: Try to find revision from revision clouds on sheet
                // This is more complex and would require iterating through revision clouds
                // For now, we'll return empty if the above methods don't work

                // TODO: Implement revision cloud reading if needed
                // This would involve:
                // 1. Get all Revision elements in document
                // 2. Check which revisions are issued on this sheet
                // 3. Return the latest revision

                return string.Empty;
            }
            catch
            {
                // If any error occurs during revision extraction, return empty
                return string.Empty;
            }
        }

        /// <summary>
        /// Checks if the latest version exists on server (database integration).
        /// TODO: Implement database connectivity for version checking.
        /// </summary>
        /// <param name="combinedText">The QR code content to check</param>
        /// <returns>True if latest version exists on server, false otherwise</returns>
        public bool CheckLatestVersionFromServer(string combinedText)
        {
            // TODO: Database integration
            // This method should:
            // 1. Connect to local or remote database
            // 2. Query for the latest version of this document/sheet
            // 3. Compare with current version
            // 4. Return true if current version matches latest, false if outdated
            // 
            // Example implementation structure:
            // using (var connection = new SqlConnection(connectionString))
            // {
            //     connection.Open();
            //     var command = new SqlCommand("SELECT TOP 1 Version FROM Documents WHERE QRContent = @content ORDER BY DateModified DESC", connection);
            //     command.Parameters.AddWithValue("@content", combinedText);
            //     var result = command.ExecuteScalar();
            //     return result != null && result.ToString() == currentVersion;
            // }

            return false; // Default return until DB integration is implemented
        }

        /// <summary>
        /// Opens QR code file in external viewer/browser.
        /// </summary>
        /// <param name="filePath">Path to QR code PNG file</param>
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