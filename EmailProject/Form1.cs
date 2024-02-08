using System;
using System.IO;
using System.Windows.Forms;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using OfficeOpenXml;
using System.Data;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace EmailProjectConsole
{
    public partial class Form1 : Form
    {
        private Button button1;
        private Button button2;
        private Button button3;
        private CheckedListBox checkedListBox;
        private string selectedPSTFilePath;
        private Label lblFileInfo;
        private Label lblFileInfo1;
        private Label step1;
        private Label step2;
        private Label step3;
        private Panel panelDragAndDrop;

        

        public Form1()
        {

            InitializeComponent();
            InitializePanel();
            InitializeDragAndDrop();
            this.Text = "PST Exctractor";
            this.button1 = new Button();
            this.button2 = new Button();
            this.button3 = new Button();
            this.checkedListBox = new CheckedListBox();

            this.button1.Text = "Upload PST File";
            this.button1.Click += new EventHandler(button1_Click);
            this.button1.Size = new System.Drawing.Size(200, 23);
            this.button1.Location = new System.Drawing.Point(350, 60);
            this.checkedListBox.Items.AddRange(new object[] { "Name", "Address", "Subject", "Body", "Date & Time" });
            this.checkedListBox.Location = new System.Drawing.Point(350, 200);
            this.button2.Text = "Export To Excel";
            this.button2.Click += new EventHandler(button2_Click);
            this.button2.Size = new System.Drawing.Size(200, 23);
            this.button2.Location = new System.Drawing.Point(350, 404);
            this.button3.Text = "Export To Excel";
            this.button3.Click += new EventHandler(button3_Click);
            this.button3.Size = new System.Drawing.Size(200, 23);
            this.button3.Location = new System.Drawing.Point(350, 404);
            this.button3.Visible = false;
            this.Controls.Add(this.button1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.checkedListBox);


            this.lblFileInfo = new Label();
            this.lblFileInfo.Size = new System.Drawing.Size(300, 20);
            this.lblFileInfo.Location = new System.Drawing.Point(350, 150); // Adjust the location as needed
            this.lblFileInfo.Visible = false; // Initially set the visibility to false
            this.Controls.Add(this.lblFileInfo);


            this.lblFileInfo1 = new Label();
            this.lblFileInfo1.Size = new System.Drawing.Size(300, 20);
            this.lblFileInfo1.Location = new System.Drawing.Point(350, 100); // Adjust the location as needed
            this.lblFileInfo1.Visible = true; // Initially set the visibility to false
            this.lblFileInfo1.Text = "Drag and drop";
            this.Controls.Add(this.lblFileInfo1);


            this.step1 = new Label();
            this.step1.Size = new System.Drawing.Size(300, 20);
            this.step1.Location = new System.Drawing.Point(350, 40); // Adjust the location as needed
            this.step1.Text = $"Step 1 - Upload PST File";
            this.Controls.Add(this.step1);


            this.step2 = new Label();
            this.step2.Size = new System.Drawing.Size(300, 20);
            this.step2.Location = new System.Drawing.Point(350, 180); // Adjust the location as needed
            this.step2.Text = $"Step 2 - Customise Extraction Information";
            this.Controls.Add(this.step2);



            this.step3 = new Label();
            this.step3.Size = new System.Drawing.Size(300, 20);
            this.step3.Location = new System.Drawing.Point(350, 380); // Adjust the location as needed
            this.step3.Text = $"Step 3 - Process and Save as Excel File";
            this.Controls.Add(this.step3);


            

            ///////////////////////////////// styling  ///////////////////////////////////////

            ////////////////////////////////color scheme////////////////////////////////////////

            // Add these lines to your constructor
            this.BackColor = Color.FromArgb(240, 240, 255); // Set a light lavender background color

            // Example color scheme with calming tones
            Color buttonColor = Color.FromArgb(128, 191, 219); // Sky blue
            Color labelColor = Color.FromArgb(85, 96, 116); // Slate gray
            Color whiteColor = Color.White;

            // Apply colors to controls
            button1.BackColor = buttonColor;
            button1.ForeColor = whiteColor; // Set text color
            button2.BackColor = buttonColor;
            button2.ForeColor = whiteColor;
            checkedListBox.BackColor = Color.White;
            // Set text color
            step1.BackColor = labelColor;
            step1.ForeColor = whiteColor; // Set text color
            step2.BackColor = labelColor;
            step2.ForeColor = whiteColor; // Set text color
            step3.BackColor = labelColor;
            step3.ForeColor = whiteColor; // Set text color


            //////////////////////font///////////////////////////////////

            // Add these lines to your constructor
            Font labelFont = new Font("Arial", 12, FontStyle.Bold); // Example font for labels
            Font buttonFont = new Font("Arial", 10, FontStyle.Bold); // Example font for buttons

            // Apply fonts to controls
            button1.Font = buttonFont;
            button2.Font = buttonFont;
            checkedListBox.Font = new Font("Arial", 10); // Example font for CheckedListBox
            lblFileInfo.Font = labelFont;
            step1.Font = labelFont;
            step2.Font = labelFont;
            step3.Font = labelFont;

            ///////////////////////////////////alignment////////////////////////////////////////

            button1.TextAlign = ContentAlignment.MiddleCenter; // Center the text in the button
            button2.TextAlign = ContentAlignment.MiddleCenter;

            lblFileInfo.TextAlign = ContentAlignment.MiddleCenter; // Align text to the center in the label
            step1.TextAlign = ContentAlignment.MiddleLeft;
            step2.TextAlign = ContentAlignment.MiddleLeft;
            step3.TextAlign = ContentAlignment.MiddleLeft;






        }
        private void InitializePanel()
        {
            // Create the panel with specified properties
            panelDragAndDrop = new Panel
            {
                Location = new Point(350, 80),
                Name = "panel1",
                Size = new Size(200, 75),
                BackColor = Color.AliceBlue,
                BorderStyle = BorderStyle.FixedSingle
            };

            // Create a label within the panel
            lblFileInfo = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Visible = false
            };

            lblFileInfo1 = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Visible = true,
                Text = "Drag And Drop PST File"


            };
            // Add the label to the panel
            panelDragAndDrop.Controls.Add(lblFileInfo);
            panelDragAndDrop.Controls.Add(lblFileInfo1);
            // Add the panel to the form's Controls collection
            Controls.Add(panelDragAndDrop);
        }
        private void InitializeDragAndDrop()
        {
            // Allow the panel to accept files when dragged and dropped
            panelDragAndDrop.AllowDrop = true;

            // Event handlers for drag and drop functionality
            panelDragAndDrop.DragEnter += PanelDragAndDrop_DragEnter;
            panelDragAndDrop.DragDrop += PanelDragAndDrop_DragDrop;
        }

        private void PanelDragAndDrop_DragEnter(object sender, DragEventArgs e)
        {
            // Check if the dragged data is a file
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void PanelDragAndDrop_DragDrop(object sender, DragEventArgs e)
        {
            // Get the array of file paths from the dropped data
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

            // Process each file
            foreach (string file in files)
            {
                ProcessUploadedFile(file);
            }
        }

        private void ProcessUploadedFile(string filePath)
        {
            // Add your file upload logic here
            // For now, let's just display the file path in a MessageBox
            MessageBox.Show("File Uploaded: " + Path.GetFileName(filePath));

            // Update the label with file information
            lblFileInfo.Text = $"File Upload Sucess: {Path.GetFileName(filePath)}";
            lblFileInfo.Visible = true;
            lblFileInfo1.Visible = false;
            panelDragAndDrop.AllowDrop = false;
            button3.Visible = true;
        }
        private void RemovePST (string filepath)
        {
            MessageBox.Show("PST File Removed: "+Path.GetFileName(filepath));
            panelDragAndDrop.AllowDrop = true;
            lblFileInfo.Visible = false;
            lblFileInfo1.Visible = true;
        }
    private void Form1_Paint(object sender, PaintEventArgs e)
        {
            int gradientStartX = 0;    // X-coordinate where the gradient starts
            int gradientEndX = this.Width;   // X-coordinate where the gradient ends

            using (LinearGradientBrush brush = new LinearGradientBrush(
                new Point(gradientStartX, 0), new Point(gradientEndX, 0),
                Color.DarkBlue, Color.Blue))
            {
                // Fill the left side with DarkBlue
                e.Graphics.FillRectangle(brush, new Rectangle(0, 0, gradientStartX, this.Height));

                // Fill the middle with Grey
                brush.InterpolationColors = new ColorBlend
                {
                    Colors = new Color[] { Color.DarkBlue, Color.Gray, Color.Blue },
                    Positions = new float[] { 0, 0.5f, 1 },
                };
                e.Graphics.FillRectangle(brush, new Rectangle(gradientStartX, 0, gradientEndX - gradientStartX, this.Height));
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "PST files (*.pst)|*.pst|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                selectedPSTFilePath = openFileDialog.FileName; // Store the file path
                lblFileInfo.Text = $"File uploaded: {Path.GetFileName(selectedPSTFilePath)}";
                lblFileInfo.Visible = true; // Make the label visible
                MessageBox.Show("PST file uploaded: " + Path.GetFileName(selectedPSTFilePath));
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(selectedPSTFilePath))
            {
                Dictionary<string, EmailExportModel> emailList = ImportPST(selectedPSTFilePath);
                SaveToExcel(emailList, checkedListBox);
            }
            else
            {
                MessageBox.Show("Please upload a PST file first.");
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            RemovePST(selectedPSTFilePath);
        }
        private static void SaveToExcel(Dictionary<string, EmailExportModel> EmailList, CheckedListBox checkedListBox)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Email List");

                // Add headers based on selected items in checkedListBox
                for (int i = 0; i < checkedListBox.CheckedItems.Count; i++)
                {
                    string propertyName = checkedListBox.CheckedItems[i].ToString();
                    worksheet.Cells[1, i + 1].Value = propertyName;
                }

                int rowNumber = 2;

                foreach (var item in EmailList)
                {
                    int colNumber = 1;

                    foreach (var checkedItem in checkedListBox.CheckedItems)
                    {
                        string propertyName = checkedItem.ToString();

                        // Handle "Date & Time" column
                        if (propertyName == "Date & Time")
                        {
                            DateTime dateTimeToUse;

                            // Use ReceivedDate if available, otherwise use SentDate
                            if (item.Value.ReceivedDate != DateTime.MinValue)
                            {
                                dateTimeToUse = item.Value.ReceivedDate;
                            }
                            else
                            {
                                dateTimeToUse = item.Value.SentDate;
                            }

                            worksheet.Cells[rowNumber, colNumber].Value = dateTimeToUse.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        else
                        {
                            // Use reflection to get property value based on the checkedListBox item
                            object propertyValue = typeof(EmailExportModel).GetProperty(propertyName)?.GetValue(item.Value);

                            worksheet.Cells[rowNumber, colNumber].Value = propertyValue?.ToString() ?? "";
                        }

                        colNumber++;
                    }

                    rowNumber++;
                }

                worksheet.Cells["A:AZ"].AutoFitColumns();

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string outputFilePath = saveFileDialog.FileName;
                    package.SaveAs(new FileInfo(outputFilePath));
                    MessageBox.Show("Excel File Saved To:" + outputFilePath);
                }
            }
        }
        //
        public static Dictionary<string, EmailExportModel> EmailList = new Dictionary<string, EmailExportModel>();
        public static Dictionary<string, List<string>> DictionaryForCategories = new Dictionary<string, List<string>>();

        static void TraverseSubfolders(FolderInfo folderInfo, PersonalStorage personalStorage)
        {
            try
            {
                FolderInfoCollection subfolders = folderInfo.GetSubFolders();

                if (subfolders.Count > 0)
                {
                    foreach (FolderInfo subfolder in subfolders)
                    {
                        MessageInfoCollection messageInfoCollection = subfolder.GetContents();

                        if (subfolder.DisplayName != "Deleted Items" &&
                            subfolder.DisplayName != "Drafts" &&
                            subfolder.DisplayName != "EventCheckPoints")
                        {
                            int iIndex = 0;

                            foreach (MessageInfo messageInfo in messageInfoCollection)
                            {
                                MapiMessage mapi = personalStorage.ExtractMessage(messageInfo);

                                EmailExportModel EmailItemModel = new EmailExportModel();
                                EmailItemModel.Index = iIndex;
                                EmailItemModel.Name = mapi.SenderName;
                                EmailItemModel.Address = mapi.SenderEmailAddress;
                                EmailItemModel.Subject = mapi.Subject;
                                EmailItemModel.Body = mapi.Body;
                                EmailItemModel.Categories = subfolder.DisplayName;

                                DateTime defaultDate = new DateTime(2000, 1, 1);

                                // Use ClientSubmitTime for sent date
                                EmailItemModel.SentDate = (mapi.ClientSubmitTime != DateTime.MinValue) ? mapi.ClientSubmitTime : defaultDate;

                                // Use DeliveryDate for received date
                                EmailItemModel.ReceivedDate = (mapi.DeliveryTime != DateTime.MinValue) ? mapi.DeliveryTime : defaultDate;

                                if (EmailItemModel.Address != null)
                                {
                                    if (EmailList.ContainsKey(EmailItemModel.Address))
                                    {
                                        string categories;

                                        if (DictionaryForCategories[EmailItemModel.Address].Contains(EmailItemModel.Categories))
                                        {
                                            categories = EmailList[EmailItemModel.Address].Categories;
                                        }
                                        else
                                        {
                                            categories = EmailList[EmailItemModel.Address].Categories + ", " + EmailItemModel.Categories;
                                            DictionaryForCategories[EmailItemModel.Address].Add(EmailItemModel.Categories);
                                        }

                                        EmailList[EmailItemModel.Address] = EmailItemModel;
                                        EmailList[EmailItemModel.Address].Categories = categories;
                                    }
                                    else
                                    {
                                        EmailList.Add(EmailItemModel.Address, EmailItemModel);
                                        List<string> listOfCategories = new List<string>();
                                        listOfCategories.Add(EmailItemModel.Categories);
                                        DictionaryForCategories.Add(EmailItemModel.Address, listOfCategories);
                                    }
                                }
                                iIndex++;
                            }
                        }

                        TraverseSubfolders(subfolder, personalStorage);
                    }
                }
                else
                {
                    MessageInfoCollection messageInfoCollection = folderInfo.GetContents();

                    if (folderInfo.DisplayName != "Deleted Items" &&
                        folderInfo.DisplayName != "Drafts" &&
                        folderInfo.DisplayName != "EventCheckPoints")
                    {
                        int iIndex = 0;

                        foreach (MessageInfo messageInfo in messageInfoCollection)
                        {
                            // Get the contact information
                            MapiMessage mapi = personalStorage.ExtractMessage(messageInfo);

                            foreach (var recipient in mapi.Recipients)
                            {
                                EmailExportModel EmailItemModel = new EmailExportModel();
                                EmailItemModel.Index = iIndex;
                                EmailItemModel.Name = recipient.DisplayName;
                                EmailItemModel.Address = recipient.EmailAddress;
                                EmailItemModel.Subject = mapi.Subject;
                                EmailItemModel.Body = mapi.Body;
                                EmailItemModel.Categories = folderInfo.DisplayName;

                                DateTime defaultDate = new DateTime(2000, 1, 1); // or any other default value you prefer

                                // Use ClientSubmitTime for sent date
                                EmailItemModel.SentDate = (mapi.ClientSubmitTime != DateTime.MinValue) ? mapi.ClientSubmitTime : defaultDate;

                                // Use DeliveryDate for received date
                                EmailItemModel.ReceivedDate = (mapi.DeliveryTime != DateTime.MinValue) ? mapi.DeliveryTime : defaultDate;

                                if (EmailItemModel.Address != null)
                                {
                                    if (EmailList.ContainsKey(EmailItemModel.Address))
                                    {
                                        string categories;

                                        if (DictionaryForCategories[EmailItemModel.Address].Contains(EmailItemModel.Categories))
                                        {
                                            categories = EmailList[EmailItemModel.Address].Categories;
                                        }
                                        else
                                        {
                                            categories = EmailList[EmailItemModel.Address].Categories + ", " + EmailItemModel.Categories;
                                            DictionaryForCategories[EmailItemModel.Address].Add(EmailItemModel.Categories);
                                        }

                                        EmailList[EmailItemModel.Address] = EmailItemModel;
                                        EmailList[EmailItemModel.Address].Categories = categories;
                                    }
                                    else
                                    {
                                        EmailList.Add(EmailItemModel.Address, EmailItemModel);
                                        List<string> listOfCategories = new List<string>();
                                        listOfCategories.Add(EmailItemModel.Categories);
                                        DictionaryForCategories.Add(EmailItemModel.Address, listOfCategories);
                                    }
                                }
                                iIndex++;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
        private static DateTime GetReceivedDateFromHeaders(MapiMessage mapi)
        {
            string dateHeaderValue = mapi.Headers.Get("Date");

            if (DateTime.TryParse(dateHeaderValue, out DateTime receivedDate))
            {
                return receivedDate;
            }

            // If parsing fails, return a default date or handle the situation accordingly
            return DateTime.MinValue;
        }

        private Dictionary<string, EmailExportModel> ImportPST(string pstFilePath)
        {
            try
            {


                using (PersonalStorage personalStorage = PersonalStorage.FromFile(pstFilePath))
                {
                    FolderInfoCollection rootFolders = personalStorage.RootFolder.GetSubFolders();

                    foreach (FolderInfo folderInfo in rootFolders)
                    {
                        TraverseSubfolders(folderInfo, personalStorage);
                    }
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error importing PST: {ex.Message}");
            }
            return EmailList;
        }

        private static string[] GetExcelSheetNames(string strFilename)
        {
            OleDbConnection objConn = null;
            DataTable dt = null;

            try
            {
                String connString = string.Empty;
                if (Path.GetExtension(strFilename).EndsWith("xls"))
                    connString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + strFilename + ";" + "Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\"";
                else
                    connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strFilename + ";Extended Properties=\"Excel 12.0 Xml;HDR=Yes;IMEX=1\"";

                objConn = new OleDbConnection(connString);
                objConn.Open();

                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    return null;
                }

                string[] excelSheets = new string[dt.Rows.Count];
                int i = 0;

                foreach (DataRow row in dt.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString();
                    i++;
                }

                return excelSheets;
            }
            catch (Exception ex)
            {
                return null;
            }
            finally
            {
                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        
    }

    public class EmailExportModel
    {
        public int Index { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string Categories { get; set; }
        public DateTime SentDate { get; set; }
        public DateTime ReceivedDate { get; set; }

    }
}