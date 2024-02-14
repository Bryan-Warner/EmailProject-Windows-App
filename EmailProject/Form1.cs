using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using OfficeOpenXml;

namespace EmailProjectConsole
{
    public partial class Form1 : Form
    {
        private Button button1;
        private Button button3;
        private CheckBox selectall;
        private CheckedListBox checkedListBox;// Updated to use CustomCheckedListBox
        private string selectedPSTFilePath;
        private Label lblFileInfo;
        private Label lblFileInfo1;
        private Label lblFileInfo2;
        private Label step1;

        private Label step3;
        private Panel panelDragAndDrop;


        public Form1()
        {
            InitializeComponent();
            InitializePanel();
            InitializeDragAndDrop();

            this.Text = "PST Extractor";
            this.button1 = new Button();
            this.button1.Text = "Browse Files";
            this.button1.Click += new EventHandler(button1_Click);
            this.button1.Size = new System.Drawing.Size(200, 23);
            this.button1.Location = new System.Drawing.Point(100, 420);
            this.button1.Visible = true; // Set visibility to true
            this.button1.BackColor = Color.White;
            this.Controls.Add(this.button1);

            this.button3 = new Button();

            this.selectall = new CheckBox();
            this.selectall.Text= "Select All";
            this.selectall.Size = new System.Drawing.Size(200, 23);
            this.selectall.Location = new System.Drawing.Point(500, 160);
            this.selectall.Font = new Font("Arial", 14);
            this.Controls.Add(this.selectall);
            this.selectall.CheckedChanged += new EventHandler(selectallchange);



            // Use the updated CustomCheckedListBox
            this.checkedListBox = new CheckedListBox();
            this.checkedListBox.Items.AddRange(new object[] { "Name", "Address", "Subject", "Body", "Categories", "Date & Time", "Last Email From Each Address" });
            this.checkedListBox.Location = new System.Drawing.Point(500, 190);
            this.checkedListBox.Size = new System.Drawing.Size(300, 250);

            this.checkedListBox.BorderStyle = BorderStyle.FixedSingle;




            this.button3.Text = "Export To Excel";
            this.button3.Click += new EventHandler(button2_Click);
            this.button3.Size = new System.Drawing.Size(200, 23);
            this.button3.Location = new System.Drawing.Point(600, 417);
            this.button3.Visible = false;
            this.button3.BackColor = Color.LightGreen;

            this.Controls.Add(this.button3);
            this.Controls.Add(this.checkedListBox);

            this.lblFileInfo = new Label();
            this.lblFileInfo.Size = new System.Drawing.Size(300, 20);
            this.lblFileInfo.Location = new System.Drawing.Point(350, 150); // Adjust the location as needed
            this.lblFileInfo.Visible = false; // Initially set the visibility to false
            this.lblFileInfo.Font = new Font("Arial", 14);
            this.Controls.Add(this.lblFileInfo);

            this.step1 = new Label();
            this.step1.Size = new System.Drawing.Size(300, 30);
            this.step1.Location = new System.Drawing.Point(100, 100); // Adjust the location as needed
            this.step1.Text = $"Step 1: Upload PST File";
            this.step1.Font = new Font("Arial", 14);
            this.Controls.Add(this.step1);


            this.step3 = new Label();
            this.step3.Size = new System.Drawing.Size(300, 30);
            this.step3.Location = new System.Drawing.Point(500, 100); // Adjust the location as needed
            this.step3.Text = $"Step 2: Filter By";
            this.step3.Font = new Font("Arial", 14);
            this.Controls.Add(this.step3);

            // Styling
            this.BackColor = Color.LightGray; // Set a light lavender background color

            // Example color scheme with calming tones


            // Apply colors to controls
            checkedListBox.BackColor = Color.White;

            // Font
            Font labelFont = new Font("Arial", 12); // Example font for labels
            Font buttonFont = new Font("Arial", 12); // Example font for buttons

            // Apply fonts to controls
            checkedListBox.Font = new Font("Arial", 15); // Example font for CustomCheckedListBox
            lblFileInfo.Font = labelFont;

        }

        private void InitializePanel()
        {
            // Create the panel with specified properties
            panelDragAndDrop = new Panel
            {
                Location = new Point(100, 140),
                Name = "panel1",
                Size = new Size(300, 280),
                BackColor = Color.AliceBlue,
                BorderStyle = BorderStyle.FixedSingle
            };

            // Create a label within the panel
            lblFileInfo1 = new Label
            {
                Location = new Point(50, 100),
                TextAlign = ContentAlignment.MiddleCenter,
                Visible = true,
                Text = "Drag And Drop PST File",
                Size = new Size(200, 60),
                Font = new Font("Arial", 14)
            };
            lblFileInfo2 = new Label
            {
                Location = new Point(50, 180),
                TextAlign = ContentAlignment.MiddleCenter,
                Visible = true,
                Text = "Or",
                Size = new Size(200, 60),
                Font = new Font("Arial", 14)
                
            };

            // Use the class-level button1 instead of creating a new local one

            // Adjust the values as needed
            // Attach the event handler to the button click event

            // Add the label and button to the panel
            panelDragAndDrop.Controls.Add(lblFileInfo1);
            panelDragAndDrop.Controls.Add(lblFileInfo2);
            panelDragAndDrop.Controls.Add(button1);

            // Add the panel to the form's Controls collection
            Controls.Add(panelDragAndDrop);
        }
        //private CheckBox selectall;
       // private CheckedListBox checkedListBox;/

        private void selectallchange(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox.Items.Count; i++)
            {
                if(selectall.Checked ==true)
                {
                    checkedListBox.SetItemChecked(i, true);
                }
                else
                {
                    checkedListBox.SetItemChecked(i, false);
                }
            }
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
                // Get the array of file paths from the dropped data
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                // Check if all files have the ".pst" extension
                bool allPstFiles = files.All(file => Path.GetExtension(file)?.Equals(".pst", StringComparison.OrdinalIgnoreCase) == true);

                if (allPstFiles)
                {
                    e.Effect = DragDropEffects.Copy;
                }
                else
                {
                    // Display a message notifying the user that only PST files are allowed
                    MessageBox.Show("Invalid File Type Please Upload A PST File.", "Invalid File Type", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
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
            selectedPSTFilePath = filePath;
            panelDragAndDrop.AllowDrop = false;
            button3.Visible = true;
            lblFileInfo2.Text = "Remove File";
            lblFileInfo2.ForeColor = Color.Blue;
            lblFileInfo2.Click += button3_Click;
            lblFileInfo1.Text = $"Upload Successful: {Path.GetFileName(filePath)}";
            
        }

        private void RemovePST(string filepath)
        {
            MessageBox.Show("PST File Removed: " + Path.GetFileName(filepath));
            panelDragAndDrop.AllowDrop = true;
            lblFileInfo1.Text = "Drag And Drop PST File";
            lblFileInfo2.Text = $"Or";
            lblFileInfo2.ForeColor = Color.Black;
            button3.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "PST files (*.pst)|*.pst|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                selectedPSTFilePath = openFileDialog.FileName; // Store the file path
                MessageBox.Show("Upload Successful: " + Path.GetFileName(selectedPSTFilePath));
                button3.Visible = true;
                lblFileInfo1.Text = $"Upload Successful: {Path.GetFileName(selectedPSTFilePath)}";
                lblFileInfo2.Text = $"Remove File";
                panelDragAndDrop.AllowDrop = false;
                lblFileInfo2.ForeColor = Color.Blue;
                lblFileInfo2.Click += button3_Click;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(selectedPSTFilePath))
            {
                Dictionary<string, EmailExportModel> emailList = ImportPST(selectedPSTFilePath);

                // Check if the "Last Email From Each Address" item is checked
                

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

        private static void SaveToExcel(Dictionary<string, EmailExportModel> emailList, CheckedListBox checkedListBox)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Email List");

                // Add headers based on selected items in checkedListBox
                for (int i = 0; i < checkedListBox.CheckedItems.Count; i++)
                {
                    string propertyName = checkedListBox.CheckedItems[i].ToString();
                    if (propertyName != "Last Email From Each Address") { 
                        worksheet.Cells[1, i + 1].Value = propertyName;
                }
                }
                int rowNumber = 2;

                foreach (var item in emailList)
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

        public static Dictionary<string, EmailExportModel> EmailList = new Dictionary<string, EmailExportModel>();
        public static Dictionary<string, List<string>> DictionaryForCategories = new Dictionary<string, List<string>>();
        static void TraverseSubfolders(FolderInfo folderInfo, PersonalStorage personalStorage, CheckedListBox checkedListBox)
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

                                EmailExportModel emailItemModel = new EmailExportModel();
                                emailItemModel.Index = iIndex;
                                emailItemModel.Name = mapi.SenderName;
                                emailItemModel.Address = mapi.SenderEmailAddress;
                                emailItemModel.Subject = mapi.Subject;
                                emailItemModel.Body = mapi.Body;
                                emailItemModel.Categories = subfolder.DisplayName;

                                DateTime defaultDate = new DateTime(2000, 1, 1);

                                // Use ClientSubmitTime for sent date
                                emailItemModel.SentDate = (mapi.ClientSubmitTime != DateTime.MinValue) ? mapi.ClientSubmitTime : defaultDate;

                                // Use DeliveryDate for received date
                                emailItemModel.ReceivedDate = (mapi.DeliveryTime != DateTime.MinValue) ? mapi.DeliveryTime : defaultDate;

                                if (emailItemModel.Address != null)
                                {
                                    if (EmailList.ContainsKey(emailItemModel.Address))
                                    {
                                        string categories;

                                        if (DictionaryForCategories[emailItemModel.Address].Contains(emailItemModel.Categories))
                                        {
                                            categories = EmailList[emailItemModel.Address].Categories;
                                        }
                                        else
                                        {
                                            categories = EmailList[emailItemModel.Address].Categories + ", " + emailItemModel.Categories;
                                            DictionaryForCategories[emailItemModel.Address].Add(emailItemModel.Categories);
                                        }

                                        EmailList[emailItemModel.Address] = emailItemModel;
                                        EmailList[emailItemModel.Address].Categories = categories;
                                    }
                                    else
                                    {
                                        EmailList.Add(emailItemModel.Address, emailItemModel);
                                        List<string> listOfCategories = new List<string>();
                                        listOfCategories.Add(emailItemModel.Categories);
                                        DictionaryForCategories.Add(emailItemModel.Address, listOfCategories);
                                    }
                                }
                                iIndex++;
                            }
                        }

                        TraverseSubfolders(subfolder, personalStorage,checkedListBox);
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
                                EmailExportModel emailItemModel = new EmailExportModel();
                                emailItemModel.Index = iIndex;
                                emailItemModel.Name = recipient.DisplayName;
                                emailItemModel.Address = recipient.EmailAddress;
                                emailItemModel.Subject = mapi.Subject;
                                emailItemModel.Body = mapi.Body;
                                emailItemModel.Categories = folderInfo.DisplayName;

                                DateTime defaultDate = new DateTime(2000, 1, 1); // or any other default value you prefer

                                // Use ClientSubmitTime for sent date
                                emailItemModel.SentDate = (mapi.ClientSubmitTime != DateTime.MinValue) ? mapi.ClientSubmitTime : defaultDate;

                                // Use DeliveryDate for received date
                                emailItemModel.ReceivedDate = (mapi.DeliveryTime != DateTime.MinValue) ? mapi.DeliveryTime : defaultDate;


                                if (emailItemModel.Address != null)

                                {

                                    if (EmailList.ContainsKey(emailItemModel.Address))
                                    {
                                        if (checkedListBox.CheckedItems.Contains("Last Email From Each Address"))
                                        {
                                            string categories;

                                            if (DictionaryForCategories[emailItemModel.Address].Contains(emailItemModel.Categories))
                                            {
                                                categories = EmailList[emailItemModel.Address].Categories;
                                            }
                                            else
                                            {
                                                categories = EmailList[emailItemModel.Address].Categories + ", " + emailItemModel.Categories;
                                                DictionaryForCategories[emailItemModel.Address].Add(emailItemModel.Categories);
                                            }

                                            EmailList[emailItemModel.Address] = emailItemModel;
                                            EmailList[emailItemModel.Address].Categories = categories;
                                        }

                                       
                                    }
                                    else
                                    {
                                        EmailList.Add(emailItemModel.Address, emailItemModel);
                                        List<string> listOfCategories = new List<string>();
                                        listOfCategories.Add(emailItemModel.Categories);
                                        DictionaryForCategories.Add(emailItemModel.Address, listOfCategories);
                                    }
                                }
                                }
                                

                                iIndex++;

                            }
                        }
                    
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
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
                        TraverseSubfolders(folderInfo, personalStorage, checkedListBox);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error importing PST: {ex.Message}");
            }

            return EmailList;
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