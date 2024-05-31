using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using ExcelDataReader;
using ZXing;
using ZXing.Common;

using ZXing.QrCode;
using ZXing.Common;
using ZXing.Rendering;
using ZXing.Windows.Compatibility;
using System.Diagnostics;

namespace BarCodeGenerator
{
    public partial class Form1 : Form
    {
        private TextBox excelFilePathTextBox;
        private TextBox saveDirectoryTextBox;
        private TextBox widthTextBox;
        private TextBox heightTextBox;
        private Button uploadButton;
        private Button chooseDirectoryButton;
        private Button generateButton;
        private Label excelFileLabel;
        private Label saveDirectoryLabel;
        private Label widthLabel;
        private Label heightLabel;
        private Label processingLabel;

        public Form1()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            excelFilePathTextBox = new TextBox();
            saveDirectoryTextBox = new TextBox();
            widthTextBox = new TextBox();
            heightTextBox = new TextBox();
            uploadButton = new Button();
            chooseDirectoryButton = new Button();
            generateButton = new Button();
            excelFileLabel = new Label();
            saveDirectoryLabel = new Label();
            widthLabel = new Label();
            heightLabel = new Label();
            processingLabel = new Label();
            SuspendLayout();
            // 
            // excelFilePathTextBox
            // 
            excelFilePathTextBox.Location = new Point(150, 30);
            excelFilePathTextBox.Name = "excelFilePathTextBox";
            excelFilePathTextBox.Size = new Size(300, 31);
            excelFilePathTextBox.TabIndex = 5;
            // 
            // saveDirectoryTextBox
            // 
            saveDirectoryTextBox.Location = new Point(150, 80);
            saveDirectoryTextBox.Name = "saveDirectoryTextBox";
            saveDirectoryTextBox.Size = new Size(300, 31);
            saveDirectoryTextBox.TabIndex = 6;
            // 
            // widthTextBox
            // 
            widthTextBox.Location = new Point(150, 130);
            widthTextBox.Name = "widthTextBox";
            widthTextBox.PlaceholderText = "e.g., 500";
            widthTextBox.Size = new Size(80, 31);
            widthTextBox.TabIndex = 7;
            // 
            // heightTextBox
            // 
            heightTextBox.Location = new Point(350, 130);
            heightTextBox.Name = "heightTextBox";
            heightTextBox.PlaceholderText = "e.g., 500";
            heightTextBox.Size = new Size(80, 31);
            heightTextBox.TabIndex = 8;
            // 
            // uploadButton
            // 
            uploadButton.Location = new Point(470, 30);
            uploadButton.Name = "uploadButton";
            uploadButton.Size = new Size(100, 40);
            uploadButton.TabIndex = 9;
            uploadButton.Text = "Upload File";
            uploadButton.Click += uploadButton_Click;
            // 
            // chooseDirectoryButton
            // 
            chooseDirectoryButton.Location = new Point(470, 80);
            chooseDirectoryButton.Name = "chooseDirectoryButton";
            chooseDirectoryButton.Size = new Size(100, 40);
            chooseDirectoryButton.TabIndex = 10;
            chooseDirectoryButton.Text = "Choose Folder";
            chooseDirectoryButton.Click += chooseDirectoryButton_Click;
            // 
            // generateButton
            // 
            generateButton.Location = new Point(209, 232);
            generateButton.Name = "generateButton";
            generateButton.Size = new Size(200, 30);
            generateButton.TabIndex = 11;
            generateButton.Text = "Generate Barcodes";
            generateButton.Click += generateButton_Click;
            // 
            // excelFileLabel
            // 
            excelFileLabel.Location = new Point(50, 30);
            excelFileLabel.Name = "excelFileLabel";
            excelFileLabel.Size = new Size(100, 40);
            excelFileLabel.TabIndex = 0;
            excelFileLabel.Text = "Excel File:";
            // 
            // saveDirectoryLabel
            // 
            saveDirectoryLabel.Location = new Point(50, 80);
            saveDirectoryLabel.Name = "saveDirectoryLabel";
            saveDirectoryLabel.Size = new Size(100, 40);
            saveDirectoryLabel.TabIndex = 1;
            saveDirectoryLabel.Text = "Folder:";
            // 
            // widthLabel
            // 
            widthLabel.Location = new Point(50, 130);
            widthLabel.Name = "widthLabel";
            widthLabel.Size = new Size(100, 31);
            widthLabel.TabIndex = 2;
            widthLabel.Text = "Width:";
            // 
            // heightLabel
            // 
            heightLabel.Location = new Point(250, 130);
            heightLabel.Name = "heightLabel";
            heightLabel.Size = new Size(100, 31);
            heightLabel.TabIndex = 3;
            heightLabel.Text = "Height:";
            // 
            // processingLabel
            // 
            processingLabel.ForeColor = Color.Blue;
            processingLabel.Location = new Point(50, 193);
            processingLabel.Name = "processingLabel";
            processingLabel.Size = new Size(340, 27);
            processingLabel.TabIndex = 4;
            // 
            // Form1
            // 
            ClientSize = new Size(650, 300);
            Controls.Add(excelFileLabel);
            Controls.Add(saveDirectoryLabel);
            Controls.Add(widthLabel);
            Controls.Add(heightLabel);
            Controls.Add(processingLabel);
            Controls.Add(excelFilePathTextBox);
            Controls.Add(saveDirectoryTextBox);
            Controls.Add(widthTextBox);
            Controls.Add(heightTextBox);
            Controls.Add(uploadButton);
            Controls.Add(chooseDirectoryButton);
            Controls.Add(generateButton);
            Name = "Form1";
            Text = "Barcode Generator";
            ResumeLayout(false);
            PerformLayout();
        }

        private void uploadButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    excelFilePathTextBox.Text = openFileDialog.FileName;
                }
            }
        }

        private void chooseDirectoryButton_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    saveDirectoryTextBox.Text = folderBrowserDialog.SelectedPath;
                }
            }
        }

        private void generateButton_Click(object sender, EventArgs e)
        {
            string excelFilePath = excelFilePathTextBox.Text;
            string saveDirectory = saveDirectoryTextBox.Text;
            if (!int.TryParse(widthTextBox.Text, out int barcodeWidth) || !int.TryParse(heightTextBox.Text, out int barcodeHeight))
            {
                MessageBox.Show("Please enter valid numeric values for width and height.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!File.Exists(excelFilePath) || !Directory.Exists(saveDirectory))
            {
                MessageBox.Show("Please select a valid Excel file and save directory.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            processingLabel.Text = "Loading...";
            this.Refresh(); // Ensure the label update is shown immediately

            try
            {
                GenerateBarcodes(excelFilePath, saveDirectory, barcodeWidth, barcodeHeight);
                MessageBox.Show("Barcodes generated successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                processingLabel.Text = "";
                // Open the main folder
                Process.Start("explorer.exe", saveDirectory);
            }
        }

        private void GenerateBarcodes(string excelFilePath, string saveDirectory, int barcodeWidth, int barcodeHeight)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                if (!reader.Read()) // Read the header row
                {
                    MessageBox.Show("Excel file is empty or invalid.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Read headers
                string[] headers = new string[reader.FieldCount];
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    headers[i] = reader.GetValue(i).ToString();
                }

                // Create directories for each header
                foreach (var header in headers)
                {
                    string headerFolderPath = Path.Combine(saveDirectory, header);
                    Directory.CreateDirectory(headerFolderPath);
                }

                // Read data rows
                while (reader.Read())
                {
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        string header = headers[i];
                        string barcodeData = reader.GetValue(i)?.ToString() ?? string.Empty;

                        if (!string.IsNullOrEmpty(barcodeData))
                        {
                            GenerateBarcodeImage(header, barcodeData, saveDirectory, barcodeWidth, barcodeHeight);
                        }
                    }
                }
            }
        }


        private void GenerateBarcodeImage(string columnName, string barcodeData, string saveDirectory, int barcodeWidth, int barcodeHeight)
        {
            var options = new EncodingOptions { Height = barcodeHeight, Width = barcodeWidth };
            var writer = new BarcodeWriter<Bitmap>
            {
                Format = BarcodeFormat.CODE_128,
                Options = options,
                Renderer = new BitmapRenderer()
            };

            Bitmap barcodeBitmap = writer.Write(barcodeData);

            string outputFolder = Path.Combine(saveDirectory, columnName);
            Directory.CreateDirectory(outputFolder);

            string outputFile = Path.Combine(outputFolder, barcodeData + ".png");
            barcodeBitmap.Save(outputFile, System.Drawing.Imaging.ImageFormat.Png);
        }
    }
}
