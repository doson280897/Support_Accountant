using OfficeOpenXml;
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace Support_Accountant
{
    public partial class MainWindows : Form
    {
        private Timer curTimeTimer;
        private string[] xmlFiles;
        private string[] pdfFiles;
        private int currentProgress;
        private string folderDestination;

        public MainWindows()
        {
            InitializeComponent();
            InitializeTimer();
            ResetAllData();
        }

        #region Timer Management
        private void InitializeTimer()
        {
            curTimeTimer = new Timer
            {
                Interval = 1000 // 1 second
            };
            curTimeTimer.Tick += CurTimeTimer_Tick;
            curTimeTimer.Start();
        }

        private void CurTimeTimer_Tick(object sender, EventArgs e)
        {
            Invoke((Action)(() =>
            {
                label_CurTime.Text = DateTime.Now.ToString("ddd HH:mm:ss - dd/MM/yyyy", CultureInfo.InvariantCulture);

                // Only update progress bar if it's visible and currentProgress has changed
                if (progressBar.Visible && progressBar.Value != currentProgress)
                {
                    progressBar.Value = Math.Min(currentProgress, progressBar.Maximum);
                    progressBar.Refresh();
                }
            }));
        }
        #endregion

        #region Form Events
        private void MainWindows_Load(object sender, EventArgs e)
        {
            // Form load logic if needed
        }

        private void buttonTools_Click(object sender, EventArgs e)
        {
            contextMenuTools.Show(buttonTools, new System.Drawing.Point(0, buttonTools.Height));
        }

        private void buttonHelp_Click(object sender, EventArgs e)
        {
            contextMenuHelp.Show(buttonHelp, new System.Drawing.Point(0, buttonHelp.Height));
        }

        private void contextMenuTools_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            // Reset all data when switching tools
            ResetAllData();

            switch (e.ClickedItem.Text)
            {
                case "Trích xuất XML":
                    ShowXMLGroup();
                    break;
                case "Đổi tên PDF":
                    ShowRenameGroup();
                    break;
                default:
                    HideAllGroups();
                    break;
            }
        }

        private void contextMenuHelp_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (e.ClickedItem.Text == "Thông tin")
            {
                MessageBox.Show("Support Accountant v1.0\nDeveloped by dos4hc\nEmail:doson280897@gmail.com",
                    "Thông tin", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        #endregion

        #region Data Management
        private void ResetAllData()
        {
            xmlFiles = null;
            pdfFiles = null;
            currentProgress = 0;

            // Reset XML controls
            textBox_FolderXML.Text = string.Empty;
            label_TotalXML.Text = "0 file(s)";
            comboBox_ListXmls.Items.Clear();

            // Reset Rename controls
            textBox_FolderRename.Text = string.Empty;
            label_RenameFiles.Text = "0 XML / 0 PDF";
            comboBox_Rename.Items.Clear();

            // Reset progress bar
            progressBar.Value = 0;
            progressBar.Visible = false;
        }

        private void ShowXMLGroup()
        {
            groupBox_XML.Visible = true;
            groupBox_RenamePDF.Visible = false;
        }

        private void ShowRenameGroup()
        {
            groupBox_XML.Visible = false;
            groupBox_RenamePDF.Visible = true;
        }

        private void HideAllGroups()
        {
            groupBox_XML.Visible = false;
            groupBox_RenamePDF.Visible = false;
        }
        #endregion

        #region File Operations
        private bool LoadFilesFromFolder(string folderPath, out string[] xmlFilesFound, out string[] pdfFilesFound)
        {
            xmlFilesFound = Directory.GetFiles(folderPath, "*.xml", SearchOption.TopDirectoryOnly);
            pdfFilesFound = Directory.GetFiles(folderPath, "*.pdf", SearchOption.TopDirectoryOnly);

            return xmlFilesFound.Length > 0 || pdfFilesFound.Length > 0;
        }

        private void PopulateComboBox(ComboBox comboBox, string[] files)
        {
            comboBox.Items.Clear();
            foreach (var filePath in files)
            {
                comboBox.Items.Add(Path.GetFileName(filePath));
            }

            if (comboBox.Items.Count > 0)
            {
                comboBox.SelectedIndex = 0;
            }
        }

        private void ShowXMLContent(string filePath, string fileName)
        {
            try
            {
                string xmlContent = File.ReadAllText(filePath);
                var doc = new XmlDocument();
                doc.LoadXml(xmlContent);

                using (var stringWriter = new StringWriter())
                using (var xmlTextWriter = new XmlTextWriter(stringWriter))
                {
                    xmlTextWriter.Formatting = Formatting.Indented;
                    doc.WriteContentTo(xmlTextWriter);
                    xmlTextWriter.Flush();
                    xmlContent = stringWriter.GetStringBuilder().ToString();
                }

                using (var xmlViewer = new Form
                {
                    Text = $"XML Viewer - {fileName}",
                    Width = 800,
                    Height = 600,
                    StartPosition = FormStartPosition.CenterParent
                })
                {
                    var textBoxXml = new TextBox
                    {
                        Multiline = true,
                        ReadOnly = true,
                        ScrollBars = ScrollBars.Both,
                        Dock = DockStyle.Fill,
                        Font = new System.Drawing.Font("Consolas", 10),
                        Text = xmlContent
                    };

                    xmlViewer.Controls.Add(textBoxXml);
                    xmlViewer.ShowDialog(this);
                }
            }
            catch (Exception ex)
            {
                ShowError($"Error reading or formatting XML file: {ex.Message}");
            }
        }

        private void OpenFileWithDefaultProgram(string filePath)
        {
            try
            {
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                ShowError($"Error opening file: {ex.Message}");
            }
        }

        private void ShowError(string message)
        {
            MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void ShowInfo(string message, string title = "Information")
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        #endregion

        #region XML Group Events
        private void button_BrowseXML_Click(object sender, EventArgs e)
        {
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select XML Folder";
                folderDialog.ShowNewFolderButton = false;

                if (folderDialog.ShowDialog() != DialogResult.OK)
                    return;

                textBox_FolderXML.Text = folderDialog.SelectedPath;

                if (!LoadFilesFromFolder(folderDialog.SelectedPath, out string[] xmlFilesFound, out _))
                {
                    ShowError("No XML files found in the selected folder.");
                    label_TotalXML.Text = "0 file(s)";
                    comboBox_ListXmls.Items.Clear();
                    return;
                }

                xmlFiles = xmlFilesFound;
                int totalXml = xmlFiles.Length;

                ShowInfo($"Total XML files found: {totalXml}", "XML Files");
                label_TotalXML.Text = $"{totalXml} file(s)";
                PopulateComboBox(comboBox_ListXmls, xmlFiles);
            }
        }

        private void button_OpenXML_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox_FolderXML.Text) || comboBox_ListXmls.SelectedItem == null)
            {
                ShowError("Please select a folder and an XML file.");
                return;
            }

            string folderPath = textBox_FolderXML.Text;
            string fileName = comboBox_ListXmls.SelectedItem.ToString();
            string filePath = Path.Combine(folderPath, fileName);

            if (!File.Exists(filePath))
            {
                ShowError("Selected XML file does not exist.");
                return;
            }

            ShowXMLContent(filePath, fileName);
        }

        private void button_ExtractXMLs_Click(object sender, EventArgs e)
        {
            if (xmlFiles == null || xmlFiles.Length == 0)
            {
                ShowError("No XML files loaded. Please select a folder with XML files first.");
                return;
            }

            using (var saveDialog = new SaveFileDialog())
            {
                saveDialog.Title = "Export All Invoices to Excel";
                saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                saveDialog.FileName = "all_invoices_export.xlsx";

                if (saveDialog.ShowDialog() != DialogResult.OK)
                    return;

                ExportToExcel(saveDialog.FileName);
            }
        }
        #endregion

        #region Rename Group Events
        private void button_BrowseRename_Click(object sender, EventArgs e)
        {
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select XML/PDF Folder";
                folderDialog.ShowNewFolderButton = false;

                if (folderDialog.ShowDialog() != DialogResult.OK)
                    return;

                textBox_FolderRename.Text = folderDialog.SelectedPath;

                if (!LoadFilesFromFolder(folderDialog.SelectedPath, out string[] xmlFilesFound, out string[] pdfFilesFound))
                {
                    ShowError("No XML or PDF files found in the selected folder.");
                    label_RenameFiles.Text = "0 XML / 0 PDF";
                    comboBox_Rename.Items.Clear();
                    return;
                }

                xmlFiles = xmlFilesFound;
                pdfFiles = pdfFilesFound;

                int totalXml = xmlFiles.Length;
                int totalPdf = pdfFiles.Length;

                ShowInfo($"Found {totalXml} XML file(s) and {totalPdf} PDF file(s).", "Files Found");
                label_RenameFiles.Text = $"{totalXml} XML / {totalPdf} PDF";

                // Populate combo box with both XML and PDF files
                var allFiles = xmlFiles.Concat(pdfFiles).ToArray();
                PopulateComboBox(comboBox_Rename, allFiles);
            }
        }

        private void button_OpenRename_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox_FolderRename.Text) || comboBox_Rename.SelectedItem == null)
            {
                ShowError("Please select a folder and a file.");
                return;
            }

            string folderPath = textBox_FolderRename.Text;
            string fileName = comboBox_Rename.SelectedItem.ToString();
            string filePath = Path.Combine(folderPath, fileName);

            if (!File.Exists(filePath))
            {
                ShowError("Selected file does not exist.");
                return;
            }

            string extension = Path.GetExtension(fileName).ToLower();

            switch (extension)
            {
                case ".xml":
                    ShowXMLContent(filePath, fileName);
                    break;
                case ".pdf":
                    OpenFileWithDefaultProgram(filePath);
                    break;
                default:
                    ShowError("Unsupported file type.");
                    break;
            }
        }

        private void button_Rename_Click(object sender, EventArgs e)
        {
            if (xmlFiles == null || pdfFiles == null || (xmlFiles.Length == 0 && pdfFiles.Length == 0))
            {
                ShowError("No files loaded. Please select a folder with XML/PDF files first.");
                return;
            }

            // Ask user to select destination folder
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select folder to store renamed files";
                folderDialog.ShowNewFolderButton = true;

                if (folderDialog.ShowDialog() != DialogResult.OK)
                    return;

                folderDestination = folderDialog.SelectedPath;

                // Create required folders
                string renamedFolder = Path.Combine(folderDestination, "Renamed");
                string failedFolder = Path.Combine(folderDestination, "Renamed_failed");

                try
                {
                    Directory.CreateDirectory(renamedFolder);
                    Directory.CreateDirectory(failedFolder);

                    ShowInfo($"Folders created successfully:\n- {renamedFolder}\n- {failedFolder}", "Folders Created");
                }
                catch (Exception ex)
                {
                    ShowError($"Error creating folders: {ex.Message}");
                    return;
                }

                // Process XML files for renaming
                ProcessXmlFilesForRenaming(renamedFolder, failedFolder);
            }
        }

        private void button_FolderRename_Click(object sender, EventArgs e)
        {

        }

        private void ProcessXmlFilesForRenaming(string renamedFolder, string failedFolder)
        {
            int totalFiles = (xmlFiles?.Length ?? 0) + (pdfFiles?.Length ?? 0);

            if (totalFiles == 0)
            {
                ShowInfo("No files to process.", "Information");
                return;
            }

            progressBar.Minimum = 0;
            progressBar.Maximum = totalFiles;
            progressBar.Value = 0;
            progressBar.Visible = true;
            currentProgress = 0;

            int successCount = 0;
            int failedCount = 0;

            try
            {
                // Process XML files
                if (xmlFiles != null)
                {
                    foreach (string xmlFilePath in xmlFiles)
                    {
                        try
                        {
                            // Update progress bar immediately when starting to process each file
                            progressBar.Value = currentProgress;
                            progressBar.Refresh();
                            Application.DoEvents(); // Allow UI to update

                            if (ProcessSingleXmlFile(xmlFilePath, renamedFolder, failedFolder))
                            {
                                successCount++;
                            }
                            else
                            {
                                failedCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            CopyToFailedFolder(xmlFilePath, failedFolder, $"Processing error: {ex.Message}");
                            failedCount++;
                        }

                        currentProgress++;
                        // Update progress bar after processing each file
                        progressBar.Value = currentProgress;
                        progressBar.Refresh();
                        Application.DoEvents(); // Allow UI to update
                    }
                }

                // Process PDF files
                if (pdfFiles != null)
                {
                    foreach (string pdfFilePath in pdfFiles)
                    {
                        try
                        {
                            // Update progress bar immediately when starting to process each file
                            progressBar.Value = currentProgress;
                            progressBar.Refresh();
                            Application.DoEvents(); // Allow UI to update

                            if (ProcessSinglePdfFile(pdfFilePath, renamedFolder, failedFolder))
                            {
                                successCount++;
                            }
                            else
                            {
                                failedCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            CopyToFailedFolder(pdfFilePath, failedFolder, $"Processing error: {ex.Message}");
                            failedCount++;
                        }

                        currentProgress++;
                        // Update progress bar after processing each file
                        progressBar.Value = currentProgress;
                        progressBar.Refresh();
                        Application.DoEvents(); // Allow UI to update
                    }
                }

                // Ensure progress bar shows 100% completion
                progressBar.Value = progressBar.Maximum;
                progressBar.Refresh();

                ShowInfo($"Processing completed:\n- Successfully renamed: {successCount} files\n- Failed: {failedCount} files",
                        "Rename Process Complete");
            }
            catch (Exception ex)
            {
                ShowError($"Error during processing: {ex.Message}");
            }
            finally
            {
                progressBar.Value = 0;
                progressBar.Visible = false;
                currentProgress = 0;
            }
        }

        private bool ProcessSingleXmlFile(string xmlFilePath, string renamedFolder, string failedFolder)
        {
            try
            {
                // Load and parse XML
                var doc = new XmlDocument();
                doc.Load(xmlFilePath);

                // Extract SHDon and NLap
                string sHDon = doc.SelectSingleNode("//SHDon")?.InnerText?.Trim();
                string nLap = doc.SelectSingleNode("//NLap")?.InnerText?.Trim();

                // Validate extracted data
                if (string.IsNullOrEmpty(sHDon) || string.IsNullOrEmpty(nLap))
                {
                    CopyToFailedFolder(xmlFilePath, failedFolder, "Missing SHDon or NLap data");
                    return false;
                }

                // Parse date and format
                if (!DateTime.TryParse(nLap, out DateTime parsedDate))
                {
                    CopyToFailedFolder(xmlFilePath, failedFolder, "Invalid date format in NLap");
                    return false;
                }

                // Create new filename: yymmdd_SHDon.xml
                string datePrefix = parsedDate.ToString("yyMMdd");
                string newFileName = $"{datePrefix}_{sHDon}.xml";

                // Get unique filename if file already exists
                string finalFileName = GetUniqueFileName(renamedFolder, newFileName);
                string destinationPath = Path.Combine(renamedFolder, finalFileName);

                // Copy file to renamed folder
                File.Copy(xmlFilePath, destinationPath, false);

                return true;
            }
            catch (Exception)
            {
                // Will be handled by caller
                throw;
            }
        }

        private bool ProcessSinglePdfFile(string pdfFilePath, string renamedFolder, string failedFolder)
        {
            try
            {
                // Extract text from PDF
                string pdfText = ExtractTextFromPdf(pdfFilePath);

                if (string.IsNullOrEmpty(pdfText))
                {
                    CopyToFailedFolder(pdfFilePath, failedFolder, "Could not extract text from PDF");
                    return false;
                }

                // Debug: Save extracted text to help with pattern debugging
                //string debugTextPath = Path.Combine(failedFolder, Path.GetFileNameWithoutExtension(pdfFilePath) + "_extracted_text.txt");
                //File.WriteAllText(debugTextPath, pdfText);

                // Extract date and invoice number
                DateTime? extractedDate = ExtractDateFromPdf(pdfText);
                string invoiceNumber = ExtractInvoiceNumberFromPdf(pdfText);

                // Debug: Log what was found
                //string debugInfo = $"Extracted Date: {(extractedDate?.ToString("yyyy-MM-dd") ?? "NULL")}\n";
                //debugInfo += $"Extracted Invoice Number: {(string.IsNullOrEmpty(invoiceNumber) ? "NULL" : invoiceNumber)}\n";
                //debugInfo += $"PDF Text Preview (first 500 chars):\n{pdfText.Substring(0, Math.Min(500, pdfText.Length))}";
                //string debugInfoPath = Path.Combine(failedFolder, Path.GetFileNameWithoutExtension(pdfFilePath) + "_debug_info.txt");
                //File.WriteAllText(debugInfoPath, debugInfo);

                // Validate extracted data
                if (!extractedDate.HasValue || string.IsNullOrEmpty(invoiceNumber))
                {
                    string reason = "";
                    if (!extractedDate.HasValue) reason += "Date not found. ";
                    if (string.IsNullOrEmpty(invoiceNumber)) reason += "Invoice number not found.";

                    CopyToFailedFolder(pdfFilePath, failedFolder, reason.Trim());
                    return false;
                }

                // Create new filename: yymmdd_InvoiceNumber.pdf
                string datePrefix = extractedDate.Value.ToString("yyMMdd");
                string newFileName = $"{datePrefix}_{invoiceNumber}.pdf";

                // Get unique filename if file already exists
                string finalFileName = GetUniqueFileName(renamedFolder, newFileName);
                string destinationPath = Path.Combine(renamedFolder, finalFileName);

                // Copy file to renamed folder
                File.Copy(pdfFilePath, destinationPath, false);

                return true;
            }
            catch (Exception)
            {
                // Will be handled by caller
                throw;
            }
        }

        private string ExtractTextFromPdf(string pdfFilePath)
        {
            try
            {
                using (var document = PdfDocument.Open(pdfFilePath))
                {
                    var text = new System.Text.StringBuilder();

                    foreach (var page in document.GetPages())
                    {
                        text.AppendLine(page.Text);
                    }

                    return text.ToString();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error extracting text from PDF: {ex.Message}");
                return string.Empty;
            }
        }

        private DateTime? ExtractDateFromPdf(string pdfText)
        {
            var datePatterns = new[]
            {
                @"Ngày\s*lập:\s*(\d{1,2})/(\d{1,2})/(\d{4})",
                @"(\d{1,2})(\d{2})(\d{4})Ngày\s*năm",
                @"(\d{1,2})\s*tháng\s*(\d{1,2})\s*năm\s*(\d{4})",
                @"Ngày\s*(\d{1,2})\s*tháng\s*(\d{1,2})\s*năm\s*(\d{4})",
                @"Số:\s*[A-Z0-9]+\s*Ngày\s*(\d{1,2})\s*tháng\s*(\d{1,2})\s*năm\s*(\d{4})",
                @"Ngày\s*(?:\([^)]*\))?\s*(\d{1,2})\s*tháng\s*(?:\([^)]*\))?\s*(\d{1,2})\s*năm\s*(?:\([^)]*\))?\s*(\d{4})",
                @"Ngày\s*tháng\s*năm/?\s*Date:\s*(\d{1,2})/(\d{1,2})/(\d{4})",
                @"Ngày\s*:?\s*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})",
                @"Số:[A-Z0-9]+\d+(\d{2})tháng(\d{2})(\d{4})Ngàynăm"
            };

            foreach (string pattern in datePatterns)
            {
                var matches = Regex.Matches(pdfText, pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);

                foreach (Match match in matches)
                {
                    if (match.Success && match.Groups.Count >= 4)
                    {
                        try
                        {
                            int day = int.Parse(match.Groups[1].Value);
                            int month = int.Parse(match.Groups[2].Value);
                            int year = int.Parse(match.Groups[3].Value);

                            // Validate date ranges and prioritize recent years
                            if (year >= 2020 && year <= 2099 && month >= 1 && month <= 12 && day >= 1 && day <= 31)
                            {
                                // Additional validation: if we find multiple dates, prefer the one that's more recent
                                // and appears in the invoice header section (not in contract details)
                                var context = GetContextAroundMatch(pdfText, match);

                                // Skip dates that appear near contract references
                                if (context.Contains("Hợp đồng") || context.Contains("HDTV") || context.Contains("ngày 01/07/2024"))
                                {
                                    continue;
                                }

                                return new DateTime(year, month, day);
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Date parsing failed: {ex.Message}");
                            continue;
                        }
                    }
                }
            }

            return null;
        }

        private string GetContextAroundMatch(string text, Match match)
        {
            int start = Math.Max(0, match.Index - 100);
            int length = Math.Min(200, text.Length - start);
            return text.Substring(start, length);
        }

        private string ExtractInvoiceNumberFromPdf(string pdfText)
        {
            // Clean the text first - remove extra whitespaces and normalize
            string cleanText = Regex.Replace(pdfText, @"\s+", " ");

            var numberPatterns = new[]
            {
                @"[A-Z0-9]{0,10}(\d{8})(?=Ký hiệu|Số\(No\))",
                @"Số\(No\.\):\s*(\d+)",
                @"Số \(No\.\):\s*(\d+)",
                @"Số hóa đơn:\s*(\d+)",
                @"Số:\s*[A-Z0-9]*?(\d{8})",
            };

            foreach (string pattern in numberPatterns)
            {
                // Try with original text first
                var matches = Regex.Matches(pdfText, pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.Singleline);

                if (matches.Count == 0)
                {
                    // Try with cleaned text
                    matches = Regex.Matches(cleanText, pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.Singleline);
                }

                foreach (Match match in matches)
                {
                    if (match.Success)
                    {
                        // Find the first non-empty captured group
                        for (int i = 1; i < match.Groups.Count; i++)
                        {
                            if (match.Groups[i].Success && !string.IsNullOrEmpty(match.Groups[i].Value))
                            {
                                string number = match.Groups[i].Value.Trim();

                                // Basic validation - just check length and that it's all digits
                                if (number.Length >= 1 && number.Length <= 20 && Regex.IsMatch(number, @"^\d+$"))
                                {
                                    return number;
                                }
                            }
                        }
                    }
                }
            }

            // If no pattern matches, try a more aggressive approach
            // Look for any number that appears after "Số" or "No"
            var aggressivePatterns = new[]
            {
        @"Số.*?(\d{1,8})",
        @"No.*?(\d{1,8})",
        @"Serial.*?(\d{1,8})"
    };

            foreach (string pattern in aggressivePatterns)
            {
                var matches = Regex.Matches(cleanText, pattern, RegexOptions.IgnoreCase);

                foreach (Match match in matches)
                {
                    if (match.Success && match.Groups.Count > 1)
                    {
                        string number = match.Groups[1].Value.Trim();
                        if (number.Length >= 1 && number.Length <= 8)
                        {
                            return number;
                        }
                    }
                }
            }

            return string.Empty;
        }

        private void CopyToFailedFolder(string sourceFilePath, string failedFolder, string reason)
        {
            try
            {
                string fileName = Path.GetFileName(sourceFilePath);
                string failedFilePath = Path.Combine(failedFolder, fileName);

                // Get unique filename if file already exists in failed folder
                string uniqueFailedPath = GetUniqueFileName(failedFolder, fileName);
                failedFilePath = Path.Combine(failedFolder, uniqueFailedPath);

                File.Copy(sourceFilePath, failedFilePath, false);

                // Optionally create a log file with the reason
                //string logFileName = Path.GetFileNameWithoutExtension(uniqueFailedPath) + "_error.txt";
                //string logFilePath = Path.Combine(failedFolder, logFileName);
                //File.WriteAllText(logFilePath, $"File: {fileName}\nReason: {reason}\nDate: {DateTime.Now}");
            }
            catch (Exception ex)
            {
                // Silent fail - just log to debug if needed
                System.Diagnostics.Debug.WriteLine($"Failed to copy to failed folder: {ex.Message}");
            }
        }

        private string GetUniqueFileName(string directory, string fileName)
        {
            string nameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
            string extension = Path.GetExtension(fileName);
            string finalFileName = fileName;
            int counter = 1;

            while (File.Exists(Path.Combine(directory, finalFileName)))
            {
                finalFileName = $"{nameWithoutExtension}({counter}){extension}";
                counter++;
            }

            return finalFileName;
        }
        #endregion

        #region Excel Export
        private void ExportToExcel(string excelPath)
        {
            progressBar.Minimum = 0;
            progressBar.Maximum = xmlFiles.Length;
            progressBar.Value = 0;
            progressBar.Visible = true;
            currentProgress = 0;

            try
            {
                ExcelPackage.License.SetNonCommercialPersonal("dos4hc");
                using (var package = new ExcelPackage())
                {
                    var summarySheet = package.Workbook.Worksheets.Add("Summary");
                    CreateSummaryHeaders(summarySheet);

                    int row = 2;
                    foreach (var xmlFile in xmlFiles)
                    {
                        try
                        {
                            ProcessXmlFile(package, summarySheet, xmlFile, ref row);
                        }
                        catch (Exception)
                        {
                            // Continue processing other files if one fails
                            continue;
                        }

                        currentProgress++;
                    }

                    summarySheet.Cells[summarySheet.Dimension.Address].AutoFitColumns();
                    package.SaveAs(new FileInfo(excelPath));
                }

                ShowInfo("Excel export completed successfully.", "Export");
            }
            catch (Exception ex)
            {
                ShowError($"Error exporting to Excel: {ex.Message}");
            }
            finally
            {
                progressBar.Value = 0;
                progressBar.Visible = false;
                currentProgress = 0;
            }
        }

        private void CreateSummaryHeaders(OfficeOpenXml.ExcelWorksheet summarySheet)
        {
            var headers = new System.Collections.Generic.List<string>
            {
                "Tên Tệp", "Tên Sheet", "Số Hóa Đơn", "Ngày Lập"
            };

            if (checkBox_Seller.Checked)
            {
                headers.AddRange(new[] { "Tên Người Bán", "MST Người Bán", "Địa Chỉ Người Bán" });
            }
            if (checkBox_Buyer.Checked)
            {
                headers.AddRange(new[] { "Tên Người Mua", "MST Người Mua", "Địa Chỉ Người Mua" });
            }

            headers.AddRange(new[] { "Tổng Số Lượng", "Tổng Tiền", "Tiền Thuế", "Thành Tiền", "Đơn Vị Tiền Tệ" });

            for (int i = 0; i < headers.Count; i++)
            {
                summarySheet.Cells[1, i + 1].Value = headers[i];
                summarySheet.Cells[1, i + 1].Style.Font.Bold = true;
            }
        }

        private void ProcessXmlFile(ExcelPackage package, OfficeOpenXml.ExcelWorksheet summarySheet, string xmlFile, ref int row)
        {
            var doc = new XmlDocument();
            doc.Load(xmlFile);

            string fileName = Path.GetFileNameWithoutExtension(xmlFile);
            string sheetName = CreateSafeSheetName(fileName);

            var fileSheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName);
            if (fileSheet == null)
            {
                fileSheet = package.Workbook.Worksheets.Add(sheetName);
                CreateDetailSheet(fileSheet, doc);
            }

            var values = ExtractSummaryData(doc, fileName, sheetName);
            PopulateSummaryRow(summarySheet, fileSheet, values, row);
            row++;
        }

        private string CreateSafeSheetName(string name)
        {
            string safeName = name.Length > 31 ? name.Substring(0, 31) : name;
            return string.Join("_", safeName.Split(Path.GetInvalidFileNameChars()));
        }

        private void CreateDetailSheet(OfficeOpenXml.ExcelWorksheet sheet, XmlDocument doc)
        {
            var detailHeaders = new System.Collections.Generic.List<string>
            {
                "STT", "THHDVu (Tên hàng hóa/dịch vụ)", "DVTinh (Đơn vị tính)",
                "SLuong (Số lượng)", "DGia (Đơn giá)", "ThTien (Thành tiền)",
                "TSuat (Thuế suất)", "DVTTe (Đơn vị tiền tệ)"
            };

            if (checkBox_Seller2.Checked)
            {
                detailHeaders.AddRange(new[] { "Tên Người Bán", "MST Người Bán", "Địa Chỉ Người Bán" });
            }
            if (checkBox_Buyer2.Checked)
            {
                detailHeaders.AddRange(new[] { "Tên Người Mua", "MST Người Mua", "Địa Chỉ Người Mua" });
            }

            // Set headers
            for (int i = 0; i < detailHeaders.Count; i++)
            {
                sheet.Cells[1, i + 1].Value = detailHeaders[i];
                sheet.Cells[1, i + 1].Style.Font.Bold = true;
            }

            // Process detail rows
            var nodes = doc.SelectNodes("//HHDVu");
            int detailRow = 2;
            string currency = doc.SelectSingleNode("//DVTTe")?.InnerText ?? "";

            foreach (XmlNode node in nodes)
            {
                PopulateDetailRow(sheet, node, doc, currency, detailRow);
                detailRow++;
            }

            // Add summary table
            AddSummaryTable(sheet, doc, detailRow + 2);
            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
        }

        private void PopulateDetailRow(OfficeOpenXml.ExcelWorksheet sheet, XmlNode node, XmlDocument doc, string currency, int row)
        {
            int col = 1;
            sheet.Cells[row, col++].Value = node.SelectSingleNode("STT")?.InnerText ?? "";
            sheet.Cells[row, col++].Value = node.SelectSingleNode("THHDVu")?.InnerText ?? "";
            sheet.Cells[row, col++].Value = node.SelectSingleNode("DVTinh")?.InnerText ?? "";
            sheet.Cells[row, col++].Value = FormatDecimalString(node.SelectSingleNode("SLuong")?.InnerText ?? "");

            string dGia = node.SelectSingleNode("DGia")?.InnerText ?? "";
            string thTien = node.SelectSingleNode("ThTien")?.InnerText ?? "";
            string tSuat = node.SelectSingleNode("TSuat")?.InnerText ?? "";

            sheet.Cells[row, col++].Value = string.IsNullOrEmpty(dGia) ? "" : $"{FormatDecimalString(dGia)} {currency}";
            sheet.Cells[row, col++].Value = string.IsNullOrEmpty(thTien) ? "" : $"{FormatDecimalString(thTien)} {currency}";
            sheet.Cells[row, col++].Value = string.IsNullOrEmpty(tSuat) ? "" : FormatDecimalString(tSuat);
            sheet.Cells[row, col++].Value = currency;

            if (checkBox_Seller2.Checked)
            {
                sheet.Cells[row, col++].Value = doc.SelectSingleNode("//NBan/Ten")?.InnerText ?? "";
                sheet.Cells[row, col++].Value = doc.SelectSingleNode("//NBan/MST")?.InnerText ?? "";
                sheet.Cells[row, col++].Value = doc.SelectSingleNode("//NBan/DChi")?.InnerText ?? "";
            }
            if (checkBox_Buyer2.Checked)
            {
                sheet.Cells[row, col++].Value = doc.SelectSingleNode("//NMua/Ten")?.InnerText ?? "";
                sheet.Cells[row, col++].Value = doc.SelectSingleNode("//NMua/MST")?.InnerText ?? "";
                sheet.Cells[row, col++].Value = doc.SelectSingleNode("//NMua/DChi")?.InnerText ?? "";
            }
        }

        private void AddSummaryTable(OfficeOpenXml.ExcelWorksheet sheet, XmlDocument doc, int startRow)
        {
            var tToanNode = doc.SelectSingleNode("//TToan");
            if (tToanNode == null) return;

            // Tax summary table
            sheet.Cells[startRow, 1].Value = "Thành tiền";
            sheet.Cells[startRow, 2].Value = "Thuế suất";
            sheet.Cells[startRow, 3].Value = "Tiền thuế";

            for (int i = 1; i <= 3; i++)
                sheet.Cells[startRow, i].Style.Font.Bold = true;

            var ltsuatNodes = tToanNode.SelectNodes("THTTLTSuat/LTSuat");
            int tRow = startRow + 1;
            foreach (XmlNode ltsuat in ltsuatNodes)
            {
                sheet.Cells[tRow, 1].Value = FormatDecimalString(ltsuat.SelectSingleNode("ThTien")?.InnerText ?? "");
                sheet.Cells[tRow, 2].Value = ltsuat.SelectSingleNode("TSuat")?.InnerText ?? "";
                sheet.Cells[tRow, 3].Value = FormatDecimalString(ltsuat.SelectSingleNode("TThue")?.InnerText ?? "");
                tRow++;
            }

            // Total summary
            int summaryRow = tRow + 1;
            sheet.Cells[summaryRow, 1].Value = "Tổng cộng (chưa thuế):";
            sheet.Cells[summaryRow, 2].Value = FormatDecimalString(tToanNode.SelectSingleNode("TgTCThue")?.InnerText ?? "");
            sheet.Cells[summaryRow + 1, 1].Value = "Tổng tiền thuế:";
            sheet.Cells[summaryRow + 1, 2].Value = FormatDecimalString(tToanNode.SelectSingleNode("TgTThue")?.InnerText ?? "");
            sheet.Cells[summaryRow + 2, 1].Value = "Tổng cộng (đã thuế):";
            sheet.Cells[summaryRow + 2, 2].Value = FormatDecimalString(tToanNode.SelectSingleNode("TgTTTBSo")?.InnerText ?? "");
            sheet.Cells[summaryRow + 3, 1].Value = "Bằng chữ:";
            sheet.Cells[summaryRow + 3, 2].Value = tToanNode.SelectSingleNode("TgTTTBChu")?.InnerText ?? "";
        }

        private System.Collections.Generic.List<object> ExtractSummaryData(XmlDocument doc, string fileName, string sheetName)
        {
            var values = new System.Collections.Generic.List<object>
            {
                fileName,
                sheetName,
                doc.SelectSingleNode("//SHDon")?.InnerText ?? "",
                doc.SelectSingleNode("//NLap")?.InnerText ?? ""
            };

            if (checkBox_Seller.Checked)
            {
                values.Add(doc.SelectSingleNode("//NBan/Ten")?.InnerText ?? "");
                values.Add(doc.SelectSingleNode("//NBan/MST")?.InnerText ?? "");
                values.Add(doc.SelectSingleNode("//NBan/DChi")?.InnerText ?? "");
            }
            if (checkBox_Buyer.Checked)
            {
                values.Add(doc.SelectSingleNode("//NMua/Ten")?.InnerText ?? "");
                values.Add(doc.SelectSingleNode("//NMua/MST")?.InnerText ?? "");
                values.Add(doc.SelectSingleNode("//NMua/DChi")?.InnerText ?? "");
            }

            values.Add(doc.SelectNodes("//HHDVu/STT")?.Count ?? 0);
            values.Add(FormatDecimalString(doc.SelectSingleNode("//TgTCThue")?.InnerText ?? ""));
            values.Add(FormatDecimalString(doc.SelectSingleNode("//TgTThue")?.InnerText ?? ""));
            values.Add(FormatDecimalString(doc.SelectSingleNode("//TgTTTBSo")?.InnerText ?? ""));
            values.Add(doc.SelectSingleNode("//DVTTe")?.InnerText ?? "");

            return values;
        }

        private void PopulateSummaryRow(OfficeOpenXml.ExcelWorksheet summarySheet, OfficeOpenXml.ExcelWorksheet fileSheet, System.Collections.Generic.List<object> values, int row)
        {
            for (int col = 0; col < values.Count; col++)
            {
                if (col == 1) // Sheet name column with hyperlink
                {
                    summarySheet.Cells[row, col + 1].Hyperlink = new OfficeOpenXml.ExcelHyperLink($"'{fileSheet.Name}'!A1", fileSheet.Name);
                    summarySheet.Cells[row, col + 1].Value = fileSheet.Name;
                    summarySheet.Cells[row, col + 1].Style.Font.UnderLine = true;
                    summarySheet.Cells[row, col + 1].Style.Font.Color.SetColor(System.Drawing.Color.Blue);
                }
                else
                {
                    summarySheet.Cells[row, col + 1].Value = values[col];
                }
            }
        }

        private string FormatDecimalString(string value)
        {
            if (decimal.TryParse(value, out decimal result))
            {
                if (result == Math.Truncate(result))
                    return result.ToString("#,##0", CultureInfo.InvariantCulture);
                return result.ToString("#,##0.###", CultureInfo.InvariantCulture);
            }
            return value;
        }
        #endregion

        #region IDisposable
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                curTimeTimer?.Stop();
                curTimeTimer?.Dispose();
            }
            base.Dispose(disposing);
        }
        #endregion
    }
}
