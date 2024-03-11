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
        private CheckedListBox checkedListBox;
        private string selectedPSTFilePath;
        private Label lblFileInfo;
        private Label lblFileInfo1;
        private Label lblFileInfo2;
        private Label lblFileInfo5;
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
            this.button1.Size = new System.Drawing.Size(125, 33);
            this.button1.Location = new System.Drawing.Point(180, 455);
            this.button1.Visible = true; 
            this.button1.BackColor = Color.White;
            this.Controls.Add(this.button1);

            this.button3 = new Button();

            this.selectall = new CheckBox();
            this.selectall.Text= "Select All";
            this.selectall.Size = new System.Drawing.Size(200, 23);
            this.selectall.Location = new System.Drawing.Point(500, 145);
            this.selectall.Font = new Font("Arial", 14);
            this.Controls.Add(this.selectall);
            this.selectall.CheckedChanged += new EventHandler(selectallchange);



            
            this.checkedListBox = new CheckedListBox();
            this.checkedListBox.Items.AddRange(new object[] { "Name", "Address", "Subject", "Body", "Categories", "Date & Time" });
            this.checkedListBox.Location = new System.Drawing.Point(500, 225);
            this.checkedListBox.Size = new System.Drawing.Size(300, 250);
            this.checkedListBox.BackColor = Color.LightGray;

            this.lblFileInfo5 = new Label();
            this.lblFileInfo5.Text = "OR";
            this.lblFileInfo5.Size = new System.Drawing.Size(125, 20);
            this.lblFileInfo5.Location = new System.Drawing.Point(230, 430);
            this.lblFileInfo5.Visible = true;
            this.lblFileInfo5.Font = new Font("Arial", 14);
            this.Controls.Add(this.lblFileInfo5);


            this.button3.Text = "Export To Excel";
            this.button3.Click += new EventHandler(button2_Click);
            this.button3.Size = new System.Drawing.Size(125, 33);
            this.button3.Location = new System.Drawing.Point(500, 455);
            this.button3.Visible = false;
            this.button3.BackColor = Color.LightBlue;

            this.Controls.Add(this.button3);
            this.Controls.Add(this.checkedListBox);

            this.lblFileInfo = new Label();
            this.lblFileInfo.Size = new System.Drawing.Size(100, 20);
            this.lblFileInfo.Location = new System.Drawing.Point(350, 150); 
            this.lblFileInfo.Visible = false; 
            this.lblFileInfo.Font = new Font("Arial", 14);
            this.Controls.Add(this.lblFileInfo);

            this.step1 = new Label();
            this.step1.Size = new System.Drawing.Size(300, 30);
            this.step1.Location = new System.Drawing.Point(100, 100); 
            this.step1.Text = $"Step 1: Upload PST File";
            this.step1.Font = new Font("Arial", 14);
            this.Controls.Add(this.step1);


            this.step3 = new Label();
            this.step3.Size = new System.Drawing.Size(300, 30);
            this.step3.Location = new System.Drawing.Point(500, 100);
            this.step3.Text = $"Step 2: Filter By";
            this.step3.Font = new Font("Arial", 14);
            this.Controls.Add(this.step3);

            
            this.BackColor = Color.LightGray; 

           


         
            checkedListBox.BackColor = Color.LightGray;

            checkedListBox.BorderStyle = BorderStyle.None;
            Font labelFont = new Font("Arial", 12); 
            Font buttonFont = new Font("Arial", 12); 

            
            checkedListBox.Font = new Font("Arial", 15); 
            lblFileInfo.Font = labelFont;

        }

        private void InitializePanel()
        {

            panelDragAndDrop = new Panel
            {
                Location = new Point(100, 140),
                Name = "panel1",
                Size = new Size(300, 280),
                BackColor = Color.AliceBlue,
                BorderStyle = BorderStyle.FixedSingle
            };

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
                Visible = false,
                Text = "Or",
                Size = new Size(200, 60),
                Font = new Font("Arial", 14)

            };

            panelDragAndDrop.Controls.Add(lblFileInfo1);
            panelDragAndDrop.Controls.Add(lblFileInfo2);
            panelDragAndDrop.Controls.Add(button1);

            
            Controls.Add(panelDragAndDrop);
        }


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
            
            panelDragAndDrop.AllowDrop = true;

            
            panelDragAndDrop.DragEnter += PanelDragAndDrop_DragEnter;
            panelDragAndDrop.DragDrop += PanelDragAndDrop_DragDrop;
        }

        private void PanelDragAndDrop_DragEnter(object sender, DragEventArgs e)
        {
            
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                
                bool allPstFiles = files.All(file => Path.GetExtension(file)?.Equals(".pst", StringComparison.OrdinalIgnoreCase) == true);

                if (allPstFiles)
                {
                    e.Effect = DragDropEffects.Copy;
                }
                else
                {
                    
                    MessageBox.Show("Invalid File Type Please Upload A PST File.", "Invalid File Type", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void PanelDragAndDrop_DragDrop(object sender, DragEventArgs e)
        {
            
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

           
            foreach (string file in files)
            {
                ProcessUploadedFile(file);
            }
        }

        private void ProcessUploadedFile(string filePath)
        {
          
            MessageBox.Show("File Uploaded: " + Path.GetFileName(filePath));

            
            selectedPSTFilePath = filePath;
            panelDragAndDrop.AllowDrop = false;
            button3.Visible = true;
            lblFileInfo2.Text = "Remove File";
            lblFileInfo2.Visible = true;
            lblFileInfo2.ForeColor = Color.Blue;
            lblFileInfo2.Click += button3_Click;
            lblFileInfo1.Text = $"Upload Successful: {Path.GetFileName(filePath)}";
            EmailList.Clear();

        }

        private void RemovePST(string filepath)
        {
            MessageBox.Show("PST File Removed: " + Path.GetFileName(filepath));
            panelDragAndDrop.AllowDrop = true;
            lblFileInfo1.Text = "Drag And Drop PST File";
            lblFileInfo2.Visible=true;
            lblFileInfo2.ForeColor = Color.Black;
            button3.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "PST files (*.pst)|*.pst|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                selectedPSTFilePath = openFileDialog.FileName; 
                MessageBox.Show("Upload Successful: " + Path.GetFileName(selectedPSTFilePath));
                button3.Visible = true;
                lblFileInfo1.Text = $"Upload Successful: {Path.GetFileName(selectedPSTFilePath)}";
                lblFileInfo2.Text = $"Remove File";
                lblFileInfo2.Visible = true;
                panelDragAndDrop.AllowDrop = false;
                lblFileInfo2.ForeColor = Color.Blue;
                lblFileInfo2.Click += button3_Click;
                lblFileInfo2.Cursor = Cursors.Hand;
                EmailList.Clear();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(selectedPSTFilePath))
            {
                
                if (checkedListBox.CheckedItems.Count == 0)
                {
                    MessageBox.Show("Please select a minimum of 1 filter.");
                    return; 
                }

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

        private static void SaveToExcel(Dictionary<string, EmailExportModel> emailList, CheckedListBox checkedListBox)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Email List");

                
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

                        
                        if (propertyName == "Date & Time")
                        {
                            DateTime dateTimeToUse;

                            
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

                                EmailExportModel EmailItemModel = new EmailExportModel();
                                EmailItemModel.Index = iIndex;
                                EmailItemModel.Name = mapi.SenderName;
                                EmailItemModel.Address = mapi.SenderEmailAddress;
                                EmailItemModel.Subject = mapi.Subject;
                                EmailItemModel.Body = mapi.Body;
                                EmailItemModel.Categories = subfolder.DisplayName;

                                DateTime defaultDate = new DateTime(2000, 1, 1);

                               
                                EmailItemModel.SentDate = (mapi.ClientSubmitTime != DateTime.MinValue) ? mapi.ClientSubmitTime : defaultDate;

                               
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

                        TraverseSubfolders(subfolder, personalStorage, checkedListBox);
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

                                DateTime defaultDate = new DateTime(2000, 1, 1); 

                               
                                EmailItemModel.SentDate = (mapi.ClientSubmitTime != DateTime.MinValue) ? mapi.ClientSubmitTime : defaultDate;

                                
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