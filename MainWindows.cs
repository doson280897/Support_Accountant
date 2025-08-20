using System;
using System.Windows.Forms;
using System.Globalization;
using OfficeOpenXml;
using System.Linq;

namespace Support_Accountant
{
    public partial class MainWindows : Form
    {
        private Timer curTimeTimer;
        // Change the type of xmlFiles from string to string[] at the field declaration
        private string[] xmlFiles;

        public MainWindows()
        {
            InitializeComponent();
            curTimeTimer = new Timer();
            curTimeTimer.Interval = 1000; // 1 second
            curTimeTimer.Tick += CurTimeTimer_Tick;
            curTimeTimer.Start();
        }

        private void CurTimeTimer_Tick(object sender, EventArgs e)
        {
            label_CurTime.Text = DateTime.Now.ToString("ddd HH:mm:ss - dd/MM/yyyy", CultureInfo.InvariantCulture);
        }

        private void MainWindows_Load(object sender, EventArgs e)
        {

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
            if (e.ClickedItem.Text == "Trích xuất XML")
            {
                groupBox_XML.Visible = true;
                groupBox_RenamePDF.Visible = false;
            }
            else if (e.ClickedItem.Text == "Đổi tên PDF")
            {
                groupBox_XML.Visible = false;
                groupBox_RenamePDF.Visible = true;
            }
            else
            {
                groupBox_XML.Visible = false;
                groupBox_RenamePDF.Visible = false;
            }
        }

        private void contextMenuHelp_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (e.ClickedItem.Text == "Thông tin")
            { 
                MessageBox.Show("Support Accountant v1.0\nDeveloped by dos4hc\nEmail:doson280897@gmail.com", "Thông tin", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button_BrowseXML_Click(object sender, EventArgs e)
        {
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select XML Folder";
                folderDialog.ShowNewFolderButton = false;
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    textBox_FolderXML.Text = folderDialog.SelectedPath;

                    // In button_BrowseXML_Click, keep the assignment as is:
                    xmlFiles = System.IO.Directory.GetFiles(folderDialog.SelectedPath, "*.xml", System.IO.SearchOption.TopDirectoryOnly);
                    int totalXml = xmlFiles.Length;

                    if (totalXml == 0)
                    {
                        MessageBox.Show("No XML files found in the selected folder.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        label_TotalXML.Text = "0 file(s)";
                        comboBox_ListXmls.Items.Clear();
                        return;
                    }

                    // Show message box with total XML files found
                    MessageBox.Show($"Total XML files found: {totalXml}", "XML Files", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // Update label_TotalXML
                    label_TotalXML.Text = $"{totalXml} file(s)";

                    // Add file names to comboBox_ListXmls
                    comboBox_ListXmls.Items.Clear();
                    foreach (var filePath in xmlFiles)
                    {
                        comboBox_ListXmls.Items.Add(System.IO.Path.GetFileName(filePath));
                    }
                    // After adding items to comboBox_ListXmls, select the first item if available
                    comboBox_ListXmls.Items.Clear();
                    foreach (var filePath in xmlFiles)
                    {
                        comboBox_ListXmls.Items.Add(System.IO.Path.GetFileName(filePath));
                    }
                    if (comboBox_ListXmls.Items.Count > 0)
                    {
                        comboBox_ListXmls.SelectedIndex = 0;
                    }
                }
            }
        }

        private void button_OpenXML_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox_FolderXML.Text) || comboBox_ListXmls.SelectedItem == null)
            {
                MessageBox.Show("Please select a folder and an XML file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string folderPath = textBox_FolderXML.Text;
            string fileName = comboBox_ListXmls.SelectedItem.ToString();
            string filePath = System.IO.Path.Combine(folderPath, fileName);

            if (!System.IO.File.Exists(filePath))
            {
                MessageBox.Show("Selected XML file does not exist.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string xmlContent;
            try
            {
                xmlContent = System.IO.File.ReadAllText(filePath);
                // Format XML
                var doc = new System.Xml.XmlDocument();
                doc.LoadXml(xmlContent);
                using (var stringWriter = new System.IO.StringWriter())
                using (var xmlTextWriter = new System.Xml.XmlTextWriter(stringWriter))
                {
                    xmlTextWriter.Formatting = System.Xml.Formatting.Indented;
                    doc.WriteContentTo(xmlTextWriter);
                    xmlTextWriter.Flush();
                    xmlContent = stringWriter.GetStringBuilder().ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading or formatting XML file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Form xmlViewer = new Form
            {
                Text = $"XML Viewer - {fileName}",
                Width = 800,
                Height = 600
            };

            TextBox textBoxXml = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                Dock = DockStyle.Fill,
                Font = new System.Drawing.Font("Consolas", 10),
                Text = xmlContent
            };

            xmlViewer.Controls.Add(textBoxXml);
            xmlViewer.StartPosition = FormStartPosition.CenterParent;
            xmlViewer.ShowDialog(this);
        }

        private async void button_ExtractXMLs_Click(object sender, EventArgs e)
        {
            // Step 1: Check if xmlFiles is loaded
            if (xmlFiles == null || xmlFiles.Length == 0)
            {
                MessageBox.Show("No XML files loaded. Please select a folder with XML files first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Step 2: Show SaveFileDialog for Excel file
            using (var saveDialog = new SaveFileDialog())
            {
                saveDialog.Title = "Export All Invoices to Excel";
                saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                saveDialog.FileName = "all_invoices_export.xlsx";
                if (saveDialog.ShowDialog() != DialogResult.OK)
                    return;

                string excelPath = saveDialog.FileName;

                // Initialize progress bar
                progressBar.Minimum = 0;
                progressBar.Maximum = xmlFiles.Length;
                progressBar.Value = 0;
                progressBar.Visible = true;

                try
                {
                    ExcelPackage.License.SetNonCommercialPersonal("dos4hc");
                    using (var package = new ExcelPackage())
                    {
                        var summarySheet = package.Workbook.Worksheets.Add("Summary");

                        // Step 4: Write header
                        var headers = new System.Collections.Generic.List<string>
                        {
                            "Tên Tệp", "Tên Sheet", "Số Hóa Đơn", "Ngày Lập"
                        };
                        if (checkBox_Seller.Checked)
                        {
                            headers.Add("Tên Người Bán");
                            headers.Add("MST Người Bán");
                            headers.Add("Địa Chỉ Người Bán");
                        }
                        if (checkBox_Buyer.Checked)
                        {
                            headers.Add("Tên Người Mua");
                            headers.Add("MST Người Mua");
                            headers.Add("Địa Chỉ Người Mua");
                        }
                        headers.Add("Tổng Số Lượng");
                        headers.Add("Tổng Tiền");
                        headers.Add("Tiền Thuế");
                        headers.Add("Thành Tiền");
                        headers.Add("Đơn Vị Tiền Tệ");

                        for (int i = 0; i < headers.Count; i++)
                        {
                            summarySheet.Cells[1, i + 1].Value = headers[i];
                            summarySheet.Cells[1, i + 1].Style.Font.Bold = true;
                        }

                        // Step 5: Write data rows and create sheets for each file
                        int row = 2;
                        int currentXml = 0;
                        foreach (var xmlFile in xmlFiles)
                        {
                            try
                            {
                                var doc = new System.Xml.XmlDocument();
                                doc.Load(xmlFile);

                                var values = new System.Collections.Generic.List<object>();
                                string fileName = System.IO.Path.GetFileNameWithoutExtension(xmlFile);
                                string sheetName = fileName;

                                // Create a new worksheet for this file
                                var fileSheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName);
                                if (fileSheet == null)
                                {
                                    string safeSheetName = sheetName.Length > 31 ? sheetName.Substring(0, 31) : sheetName;
                                    safeSheetName = string.Join("_", safeSheetName.Split(System.IO.Path.GetInvalidFileNameChars()));
                                    fileSheet = package.Workbook.Worksheets.Add(safeSheetName);

                                    var detailHeaders = new System.Collections.Generic.List<string> {
                                        "STT", "THHDVu (Tên hàng hóa/dịch vụ)", "DVTinh (Đơn vị tính)", "SLuong (Số lượng)", "DGia (Đơn giá)", "ThTien (Thành tiền)", "TSuat (Thuế suất)", "DVTTe (Đơn vị tiền tệ)"
                                    };

                                    if (checkBox_Seller2.Checked)
                                    {
                                        detailHeaders.Add("Tên Người Bán");
                                        detailHeaders.Add("MST Người Bán");
                                        detailHeaders.Add("Địa Chỉ Người Bán");
                                    }
                                    if (checkBox_Buyer2.Checked)
                                    {
                                        detailHeaders.Add("Tên Người Mua");
                                        detailHeaders.Add("MST Người Mua");
                                        detailHeaders.Add("Địa Chỉ Người Mua");
                                    }

                                    for (int i = 0; i < detailHeaders.Count; i++)
                                    {
                                        fileSheet.Cells[1, i + 1].Value = detailHeaders[i];
                                        fileSheet.Cells[1, i + 1].Style.Font.Bold = true;
                                    }

                                    var nodes = doc.SelectNodes("//HHDVu");
                                    int detailRow = 2;
                                    foreach (System.Xml.XmlNode node in nodes)
                                    {
                                        int col = 1;
                                        fileSheet.Cells[detailRow, col++].Value = node.SelectSingleNode("STT")?.InnerText ?? "";
                                        fileSheet.Cells[detailRow, col++].Value = node.SelectSingleNode("THHDVu")?.InnerText ?? "";
                                        fileSheet.Cells[detailRow, col++].Value = node.SelectSingleNode("DVTinh")?.InnerText ?? "";
                                        fileSheet.Cells[detailRow, col++].Value = FormatDecimalString(node.SelectSingleNode("SLuong")?.InnerText ?? "");
                                        string dGia = node.SelectSingleNode("DGia")?.InnerText ?? "";
                                        string thTien = node.SelectSingleNode("ThTien")?.InnerText ?? "";
                                        string tSuat = node.SelectSingleNode("TSuat")?.InnerText ?? "";
                                        string dvtTe = doc.SelectSingleNode("//DVTTe")?.InnerText ?? "";

                                        fileSheet.Cells[detailRow, col++].Value = string.IsNullOrEmpty(dGia) ? "" : $"{FormatDecimalString(dGia)} {dvtTe}";
                                        fileSheet.Cells[detailRow, col++].Value = string.IsNullOrEmpty(thTien) ? "" : $"{FormatDecimalString(thTien)} {dvtTe}";
                                        fileSheet.Cells[detailRow, col++].Value = string.IsNullOrEmpty(tSuat) ? "" : $"{FormatDecimalString(tSuat)}";
                                        fileSheet.Cells[detailRow, col++].Value = dvtTe;

                                        if (checkBox_Seller2.Checked)
                                        {
                                            fileSheet.Cells[detailRow, col++].Value = doc.SelectSingleNode("//NBan/Ten")?.InnerText ?? "";
                                            fileSheet.Cells[detailRow, col++].Value = doc.SelectSingleNode("//NBan/MST")?.InnerText ?? "";
                                            fileSheet.Cells[detailRow, col++].Value = doc.SelectSingleNode("//NBan/DChi")?.InnerText ?? "";
                                        }
                                        if (checkBox_Buyer2.Checked)
                                        {
                                            fileSheet.Cells[detailRow, col++].Value = doc.SelectSingleNode("//NMua/Ten")?.InnerText ?? "";
                                            fileSheet.Cells[detailRow, col++].Value = doc.SelectSingleNode("//NMua/MST")?.InnerText ?? "";
                                            fileSheet.Cells[detailRow, col++].Value = doc.SelectSingleNode("//NMua/DChi")?.InnerText ?? "";
                                        }
                                        detailRow++;
                                    }

                                    int tableStartRow = detailRow + 2;
                                    var tToanNode = doc.SelectSingleNode("//TToan");
                                    if (tToanNode != null)
                                    {
                                        fileSheet.Cells[tableStartRow, 1].Value = "Thành tiền";
                                        fileSheet.Cells[tableStartRow, 2].Value = "Thuế suất";
                                        fileSheet.Cells[tableStartRow, 3].Value = "Tiền thuế";
                                        for (int i = 1; i <= 3; i++)
                                            fileSheet.Cells[tableStartRow, i].Style.Font.Bold = true;

                                        var ltsuatNodes = tToanNode.SelectNodes("THTTLTSuat/LTSuat");
                                        int tRow = tableStartRow + 1;
                                        foreach (System.Xml.XmlNode ltsuat in ltsuatNodes)
                                        {
                                            fileSheet.Cells[tRow, 1].Value = FormatDecimalString(ltsuat.SelectSingleNode("ThTien")?.InnerText ?? "");
                                            fileSheet.Cells[tRow, 2].Value = ltsuat.SelectSingleNode("TSuat")?.InnerText ?? "";
                                            fileSheet.Cells[tRow, 3].Value = FormatDecimalString(ltsuat.SelectSingleNode("TThue")?.InnerText ?? "");
                                            tRow++;
                                        }

                                        int summaryRow = tRow + 1;
                                        fileSheet.Cells[summaryRow, 1].Value = "Tổng cộng (chưa thuế):";
                                        fileSheet.Cells[summaryRow, 2].Value = FormatDecimalString(tToanNode.SelectSingleNode("TgTCThue")?.InnerText ?? "");
                                        fileSheet.Cells[summaryRow + 1, 1].Value = "Tổng tiền thuế:";
                                        fileSheet.Cells[summaryRow + 1, 2].Value = FormatDecimalString(tToanNode.SelectSingleNode("TgTThue")?.InnerText ?? "");
                                        fileSheet.Cells[summaryRow + 2, 1].Value = "Tổng cộng (đã thuế):";
                                        fileSheet.Cells[summaryRow + 2, 2].Value = FormatDecimalString(tToanNode.SelectSingleNode("TgTTTBSo")?.InnerText ?? "");
                                        fileSheet.Cells[summaryRow + 3, 1].Value = "Bằng chữ:";
                                        fileSheet.Cells[summaryRow + 3, 2].Value = tToanNode.SelectSingleNode("TgTTTBChu")?.InnerText ?? "";
                                    }
                                    fileSheet.Cells[fileSheet.Dimension.Address].AutoFitColumns();
                                }

                                string invoiceNumber = doc.SelectSingleNode("//SHDon")?.InnerText ?? "";
                                string invoiceDate = doc.SelectSingleNode("//NLap")?.InnerText ?? "";

                                values.Add(fileName);
                                values.Add(sheetName);
                                values.Add(invoiceNumber);
                                values.Add(invoiceDate);

                                if (checkBox_Seller.Checked)
                                {
                                    string sellerName = doc.SelectSingleNode("//NBan/Ten")?.InnerText ?? "";
                                    string sellerTax = doc.SelectSingleNode("//NBan/MST")?.InnerText ?? "";
                                    string sellerAddr = doc.SelectSingleNode("//NBan/DChi")?.InnerText ?? "";
                                    values.Add(sellerName);
                                    values.Add(sellerTax);
                                    values.Add(sellerAddr);
                                }
                                if (checkBox_Buyer.Checked)
                                {
                                    string buyerName = doc.SelectSingleNode("//NMua/Ten")?.InnerText ?? "";
                                    string buyerTax = doc.SelectSingleNode("//NMua/MST")?.InnerText ?? "";
                                    string buyerAddr = doc.SelectSingleNode("//NMua/DChi")?.InnerText ?? "";
                                    values.Add(buyerName);
                                    values.Add(buyerTax);
                                    values.Add(buyerAddr);
                                }

                                int totalQty = doc.SelectNodes("//HHDVu/STT")?.Count ?? 0;
                                string currency = doc.SelectSingleNode("//DVTTe")?.InnerText ?? "";
                                string totalAmt = FormatDecimalString(doc.SelectSingleNode("//TgTCThue")?.InnerText ?? "");
                                string taxAmt = FormatDecimalString(doc.SelectSingleNode("//TgTThue")?.InnerText ?? "");
                                string finalAmt = FormatDecimalString(doc.SelectSingleNode("//TgTTTBSo")?.InnerText ?? "");

                                values.Add(totalQty);
                                values.Add(totalAmt);
                                values.Add(taxAmt);
                                values.Add(finalAmt);
                                values.Add(currency);

                                for (int col = 0; col < values.Count; col++)
                                {
                                    if (col == 1)
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

                                row++;
                            }
                            catch (Exception)
                            {
                                continue;
                            }

                            // Update progress bar
                            currentXml++;
                            progressBar.Value = currentXml;
                            progressBar.Refresh();
                            await System.Threading.Tasks.Task.Delay(10); // allow UI update
                        }

                        summarySheet.Cells[summarySheet.Dimension.Address].AutoFitColumns();

                        // Step 6: Save Excel file
                        package.SaveAs(new System.IO.FileInfo(excelPath));
                    }

                    MessageBox.Show("Excel export completed successfully.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error exporting to Excel: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    progressBar.Value = 0;
                    progressBar.Visible = false;
                }
            }
        }

        private string FormatDecimalString(string value)
        {
            if (decimal.TryParse(value, out decimal result))
            {
                // For other currencies, show up to 3 decimals, remove trailing zeros
                if (result == Math.Truncate(result))
                    return result.ToString("#,##0", CultureInfo.InvariantCulture);

                // Format with thousands separator and up to 3 decimals, trim trailing zeros
                string formatted = result.ToString("#,##0.###", CultureInfo.InvariantCulture);
                return formatted;
            }
            return value;
        }
    }
}
