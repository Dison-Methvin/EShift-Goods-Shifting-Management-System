using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TheArtOfDevHtmlRenderer.Adapters.Entities;

namespace E_Shift_V02
{
    public partial class Admin_Dashboard : Form
    {
        private Timer timerDateTime;
        private Timer timerJobCounts;
        private string _adminUsername;

        public Admin_Dashboard(string username)
        {
            InitializeComponent();

            _adminUsername = username;
            lblAdminUsername_AdminDashboard.Text = _adminUsername;

            timerDateTime = new Timer();
            timerDateTime.Interval = 1000; // 1 second
            timerDateTime.Tick += TimerDateTime_Tick; 
            timerDateTime.Start();

            LoadJobsGrid();
            LoadDriversGrid();

            timerJobCounts = new Timer();
            timerJobCounts.Interval = 1000; // 1 second
            timerJobCounts.Tick += (s, e) =>
            {
                FetchAndDisplayAcceptedJobsCount();
                FetchAndDisplayPendingJobsCount();
            };
            timerJobCounts.Start();

            PopulateDriverIDComboBox();
            ComboBoxDriverID_AdminDriverTab.SelectedIndexChanged += ComboBoxDriverID_AdminDriverTab_SelectedIndexChanged;
            ComboBoxDriverID_AdminDriverTab.DropDown += ComboBoxDriverID_AdminDriverTab_DropDown;
            btnDriverDetailsRegister_AdminDriverTab.Enabled = true;

            dataGridViewMain_AdminMyDashboardTab.SelectionChanged += DataGridViewMain_AdminMyDashboardTab_SelectionChanged;

            btnRejectRequestedJob.Enabled = false;
            btnAcceptedRequestedJob_AdminMyDashboardTab.Enabled = false;

            LoadAssistantsGrid();
            PopulateAssistantIDComboBox();
            ComboBoxAssistantID_AdminAssistantTab.SelectedIndexChanged -= ComboBoxAssistantID_AdminAssistantTab_SelectedIndexChanged;
            ComboBoxAssistantID_AdminAssistantTab.SelectedIndexChanged += ComboBoxAssistantID_AdminAssistantTab_SelectedIndexChanged;
            if (ComboBoxAssistantID_AdminAssistantTab.Items.Count > 0)
                ComboBoxAssistantID_AdminAssistantTab.SelectedIndex = 0;

            PopulateLorryIDComboBox();
            ComboBoxLorryID_AdminLorryTab.SelectedIndexChanged += ComboBoxLorryID_AdminLorryTab_SelectedIndexChanged;

            LoadLorriesGrid();

            PopulateContainerIDComboBox();
            ComboBoxContainerID_AdminContainerTab.SelectedIndexChanged += ComboBoxContainerID_AdminContainerTab_SelectedIndexChanged;

            LoadContainersGrid();

            LoadProductTypesGrid();

            PopulateMailingCustomerIDComboBox();
            ComboBoxMailingCustomerID_AdminMailingTab.SelectedIndexChanged += ComboBoxMailingCustomerID_AdminMailingTab_SelectedIndexChanged;

            PopulateMailingDriverIDComboBox();
            ComboBoxMailingDriverID_AdminMailingTab.SelectedIndexChanged += ComboBoxMailingDriverID_AdminMailingTab_SelectedIndexChanged;

            PopulateMailingAssistantIDComboBox();
            ComboBoxMailingAssistantID.SelectedIndexChanged += ComboBoxMailingAssistantID_SelectedIndexChanged;

            LoadAdminProfilePicture();

            PopulateDriverIDComboBox_AdminJobsTab();
            ComboBoxDriverID_AdminJobsTab.SelectedIndexChanged += ComboBoxDriverID_AdminJobsTab_SelectedIndexChanged;

            PopulateAssistantIDComboBox_AdminJobsTab();
            ComboBoxAssistantID_AdminJobsTab.SelectedIndexChanged += ComboBoxAssistantID_AdminJobsTab_SelectedIndexChanged;

            PopulateLorryIDComboBox_AdminJobsTab();
            ComboBoxLorryID_AdminJobsTab.SelectedIndexChanged += ComboBoxLorryID_AdminJobsTab_SelectedIndexChanged;

            PopulateContainerIDComboBox_AdminJobsTab();
            ComboBoxContainerID_AdminJobsTab.SelectedIndexChanged += ComboBoxContainerID_AdminJobsTab_SelectedIndexChanged;

            PopulateJobIDComboBox_AdminJobsTab();
            ComboBoxJobID_AdminJobsTab.SelectedIndexChanged += ComboBoxJobID_AdminJobsTabb_SelectedIndexChanged;

            this.txtJobTotalDistance_AdminJobsTab.TextChanged += new System.EventHandler(this.txtJobTotalDistance_AdminJobsTab_TextChanged);
            Admin_Dashboard_MapInit();
        }

        private void ComboBoxDriverID_AdminDriverTab_DropDown(object sender, EventArgs e)
        {
            PopulateDriverIDComboBox();
        }
        private void LoadJobsGrid()
        {
            try
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(@"SELECT JobID, CustomerID, LorryID, DriverID, AssistantID, ContainerID, FirstProduct, FirstProductQuantity, SecondProduct, SecondProductQuantity, FirstName, LastName, PhoneNumber, Email, PickupAddress, PickupCity, DeliveryAddress, DeliveryCity, LocationType, PickupDate, TotalDistance, EstimatedTime, JobCostPerDistance, TotalCost, Status, CreatedDate, LoadStatus FROM Jobs ORDER BY CreatedDate DESC", con);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridViewMain_AdminMyDashboardTab.DataSource = dt;
                dataGridViewMain_AdminMyDashboardTab.ClearSelection();
            }
            catch (Exception)
            {
            }
            finally
            {
                con.Close();
            }
        }
        SqlConnection con = new SqlConnection("Data Source=DISON_PC\\SQLEXPRESS;Initial Catalog=ESHIFT;Integrated Security=True;Encrypt=False");

        private void TimerDateTime_Tick(object sender, EventArgs e)
        {
            TimeZoneInfo sriLankaTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Sri Lanka Standard Time");
            DateTime sriLankaTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, sriLankaTimeZone);
            lblDateTime_AdminDashboard.Text = sriLankaTime.ToString("HH:mm:ss");
        }

        private void btnLogout_AdminMyDashboardTab_Click(object sender, EventArgs e)
        {
            Role_Selection signUpForm = new Role_Selection();
            signUpForm.Show();
            this.Hide();
        }

        private void btnsearchIcon_AdminMyDashboardTab_Click(object sender, EventArgs e)
        {
            string jobIdText = txtSearchByJobID_AdminMyDashboardTab.Text.Trim();
            string customerIdText = txtSearchByCustomerID_AdminMyDashboardTab.Text.Trim();

            if (!string.IsNullOrWhiteSpace(jobIdText))
            {
                if (!int.TryParse(jobIdText, out int jobId))
                {
                    MessageBox.Show("Please enter a valid numeric Job ID.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                try
                {
                    con.Open();
                    SqlCommand searchCmd = new SqlCommand(@"SELECT JobID, CustomerID, FirstProduct, FirstProductQuantity, SecondProduct, SecondProductQuantity, FirstName, LastName, PhoneNumber, Email, PickupAddress, PickupCity, DeliveryAddress, DeliveryCity, LocationType, PickupDate, TotalDistance, EstimatedTime, JobCostPerDistance, TotalCost, Status, CreatedDate, LoadStatus FROM Jobs WHERE JobID = @JobID", con);
                    searchCmd.Parameters.AddWithValue("@JobID", jobId);
                    SqlDataAdapter da = new SqlDataAdapter(searchCmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridViewMain_AdminMyDashboardTab.DataSource = dt;
                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("No job found with the specified Job ID.", "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error searching for job: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    con.Close();
                }
                return;
            }
            else if (!string.IsNullOrWhiteSpace(customerIdText))
            {
                if (!int.TryParse(customerIdText, out int customerId))
                {
                    MessageBox.Show("Please enter a valid numeric Customer ID.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                try
                {
                    con.Open();
                    SqlCommand searchCmd = new SqlCommand(@"SELECT JobID, CustomerID, FirstProduct, FirstProductQuantity, SecondProduct, SecondProductQuantity, FirstName, LastName, PhoneNumber, Email, PickupAddress, PickupCity, DeliveryAddress, DeliveryCity, LocationType, PickupDate, TotalDistance, EstimatedTime, JobCostPerDistance, TotalCost, Status, CreatedDate, LoadStatus FROM Jobs WHERE CustomerID = @CustomerID", con);
                    searchCmd.Parameters.AddWithValue("@CustomerID", customerId);
                    SqlDataAdapter da = new SqlDataAdapter(searchCmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridViewMain_AdminMyDashboardTab.DataSource = dt;
                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("No jobs found for the specified Customer ID.", "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error searching for jobs: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    con.Close();
                }
                return;
            }
            else
            {
                MessageBox.Show("Please enter a Job ID or Customer ID to search.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnRefreshIcon_AdminMyDashboardTab_Click(object sender, EventArgs e)
        {
            LoadJobsGrid();
        }

        private void RadioButtonAll_AdminMyDashboardTab_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonAll_AdminMyDashboardTab.Checked)
            {
                LoadJobsGrid();
            }
        }

        private void RadioButtonNotPickedUp_AdminMyDashboardTab_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonNotPickedUp_AdminMyDashboardTab.Checked)
            {
                try
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand(@"SELECT JobID, CustomerID, FirstProduct, FirstProductQuantity, SecondProduct, SecondProductQuantity, FirstName, LastName, PhoneNumber, Email, PickupAddress, PickupCity, DeliveryAddress, DeliveryCity, LocationType, PickupDate, TotalDistance, EstimatedTime, JobCostPerDistance, TotalCost, Status, CreatedDate, LoadStatus FROM Jobs WHERE LoadStatus = @LoadStatus ORDER BY CreatedDate DESC", con);
                    cmd.Parameters.AddWithValue("@LoadStatus", "Not PickedUp");
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridViewMain_AdminMyDashboardTab.DataSource = dt;
                    dataGridViewMain_AdminMyDashboardTab.ClearSelection();
                }
                catch (Exception)
                {
                }
                finally
                {
                    con.Close();
                }
            }
        }

        private void RadioButtonPickedUp_AdminMyDashboardTab_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonPickedUp_AdminMyDashboardTab.Checked)
            {
                try
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand(@"SELECT JobID, CustomerID, FirstProduct, FirstProductQuantity, SecondProduct, SecondProductQuantity, FirstName, LastName, PhoneNumber, Email, PickupAddress, PickupCity, DeliveryAddress, DeliveryCity, LocationType, PickupDate, TotalDistance, EstimatedTime, JobCostPerDistance, TotalCost, Status, CreatedDate, LoadStatus FROM Jobs WHERE LoadStatus = @LoadStatus ORDER BY CreatedDate DESC", con);
                    cmd.Parameters.AddWithValue("@LoadStatus", "PickedUp");
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridViewMain_AdminMyDashboardTab.DataSource = dt;
                    dataGridViewMain_AdminMyDashboardTab.ClearSelection();
                }
                catch (Exception)
                {
                }
                finally
                {
                    con.Close();
                }
            }
        }

        private void RadioButtonDelivered_AdminMyDashboardTab_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonDelivered_AdminMyDashboardTab.Checked)
            {
                try
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand(@"SELECT JobID, CustomerID, FirstProduct, FirstProductQuantity, SecondProduct, SecondProductQuantity, FirstName, LastName, PhoneNumber, Email, PickupAddress, PickupCity, DeliveryAddress, DeliveryCity, LocationType, PickupDate, TotalDistance, EstimatedTime, JobCostPerDistance, TotalCost, Status, CreatedDate, LoadStatus FROM Jobs WHERE LoadStatus = @LoadStatus ORDER BY CreatedDate DESC", con);
                    cmd.Parameters.AddWithValue("@LoadStatus", "Delivered");
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridViewMain_AdminMyDashboardTab.DataSource = dt;
                    dataGridViewMain_AdminMyDashboardTab.ClearSelection();
                }
                catch (Exception)
                {
                }
                finally
                {
                    con.Close();
                }
            }
        }

        private void RadioButtonAccepted_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonAccepted.Checked)
            {
                try
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand(@"SELECT JobID, CustomerID, FirstProduct, FirstProductQuantity, SecondProduct, SecondProductQuantity, FirstName, LastName, PhoneNumber, Email, PickupAddress, PickupCity, DeliveryAddress, DeliveryCity, LocationType, PickupDate, TotalDistance, EstimatedTime, JobCostPerDistance, TotalCost, Status, CreatedDate, LoadStatus FROM Jobs WHERE Status = @Status ORDER BY CreatedDate DESC", con);
                    cmd.Parameters.AddWithValue("@Status", "Accepted");
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridViewMain_AdminMyDashboardTab.DataSource = dt;
                    dataGridViewMain_AdminMyDashboardTab.ClearSelection();
                }
                catch (Exception)
                {
                }
                finally
                {
                    con.Close();
                }
            }
        }

        private void RadioButtonRejected_AdminMyDashboardTab_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonRejected_AdminMyDashboardTab.Checked)
            {
                try
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand(@"SELECT JobID, CustomerID, FirstProduct, FirstProductQuantity, SecondProduct, SecondProductQuantity, FirstName, LastName, PhoneNumber, Email, PickupAddress, PickupCity, DeliveryAddress, DeliveryCity, LocationType, PickupDate, TotalDistance, EstimatedTime, JobCostPerDistance, TotalCost, Status, CreatedDate, LoadStatus FROM Jobs WHERE Status = @Status ORDER BY CreatedDate DESC", con);
                    cmd.Parameters.AddWithValue("@Status", "Rejected");
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridViewMain_AdminMyDashboardTab.DataSource = dt;
                    dataGridViewMain_AdminMyDashboardTab.ClearSelection();
                }
                catch (Exception)
                {
                }
                finally
                {
                    con.Close();
                }
            }
        }

        private void RadioButtonPendingRequests_AdminMyDashboardTab_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonPendingRequests_AdminMyDashboardTab.Checked)
            {
                try
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand(@"SELECT JobID, CustomerID, FirstProduct, FirstProductQuantity, SecondProduct, SecondProductQuantity, FirstName, LastName, PhoneNumber, Email, PickupAddress, PickupCity, DeliveryAddress, DeliveryCity, LocationType, PickupDate, TotalDistance, EstimatedTime, JobCostPerDistance, TotalCost, Status, CreatedDate, LoadStatus FROM Jobs WHERE Status = @Status ORDER BY CreatedDate DESC", con);
                    cmd.Parameters.AddWithValue("@Status", "Pending");
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridViewMain_AdminMyDashboardTab.DataSource = dt;
                    dataGridViewMain_AdminMyDashboardTab.ClearSelection();
                }
                catch (Exception)
                {
                }
                finally
                {
                    con.Close();
                }
            }
        }

        private void DataGridViewMain_AdminMyDashboardTab_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridViewMain_AdminMyDashboardTab.SelectedRows.Count > 0)
            {
                var row = dataGridViewMain_AdminMyDashboardTab.SelectedRows[0];
                var statusCell = row.Cells["Status"];
                if (statusCell != null && statusCell.Value != null && statusCell.Value.ToString() == "Pending")
                {
                    btnRejectRequestedJob.Enabled = true;
                    btnAcceptedRequestedJob_AdminMyDashboardTab.Enabled = true;
                }
                else
                {
                    btnRejectRequestedJob.Enabled = false;
                    btnAcceptedRequestedJob_AdminMyDashboardTab.Enabled = false;
                }
            }
            else
            {
                btnRejectRequestedJob.Enabled = false;
                btnAcceptedRequestedJob_AdminMyDashboardTab.Enabled = false;
            }
        }

        private void btnRejectRequestedJob_Click(object sender, EventArgs e)
        {
            if (dataGridViewMain_AdminMyDashboardTab.SelectedRows.Count > 0)
            {
                var row = dataGridViewMain_AdminMyDashboardTab.SelectedRows[0];
                var jobId = row.Cells["JobID"].Value?.ToString() ?? string.Empty;
                var customerId = row.Cells["CustomerID"].Value?.ToString() ?? string.Empty;
                Admin_Reject_Request_PopUp rejectPopUp = new Admin_Reject_Request_PopUp(jobId, customerId);
                rejectPopUp.Show();
            }
        }

        private void FetchAndDisplayAcceptedJobsCount()
        {
            int acceptedCount = 0;
            try
            {
                con.Open();
                using (SqlCommand countCmd = new SqlCommand("SELECT COUNT(*) FROM Jobs WHERE Status = 'Accepted'", con))
                {
                    object countResult = countCmd.ExecuteScalar();
                    if (countResult != null && countResult != DBNull.Value)
                        acceptedCount = Convert.ToInt32(countResult);
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                con.Close();
            }
            lblJobsAcceptedNumber_AdminMyDashboardTab.Text = acceptedCount.ToString("D2");
        }

        private void FetchAndDisplayPendingJobsCount()
        {
            int pendingCount = 0;
            try
            {
                con.Open();
                using (SqlCommand countCmd = new SqlCommand("SELECT COUNT(*) FROM Jobs WHERE Status = 'Pending'", con))
                {
                    object countResult = countCmd.ExecuteScalar();
                    if (countResult != null && countResult != DBNull.Value)
                        pendingCount = Convert.ToInt32(countResult);
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                con.Close();
            }
            lblJobsPendingNumber_AdminMyDashboardTab.Text = pendingCount.ToString("D2");
        }

        private void btnPrintJobReport_AdminMyDashboardTab_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "PDF files (*.pdf)|*.pdf";
                sfd.Title = "Save Job Report as PDF";
                sfd.FileName = "JobReport.pdf";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 40f, 40f, 60f, 50f);
                        iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, new System.IO.FileStream(sfd.FileName, System.IO.FileMode.Create));
                        doc.Open();

                        var titleFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 18);
                        var subTitleFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 12);
                        var bodyFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA, 9);
                        var boldTableFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 8);
                        var endingFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_OBLIQUE, 8);
                        var noteFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.RED);

                        iTextSharp.text.pdf.PdfPTable headerTable = new iTextSharp.text.pdf.PdfPTable(2);
                        headerTable.WidthPercentage = 100;
                        headerTable.SetWidths(new float[] { 20, 80 });

                        if (pictureBoxESHIFTLogo_AdminMyDashboardTab != null && pictureBoxESHIFTLogo_AdminMyDashboardTab.Image != null)
                        {
                            using (var ms = new System.IO.MemoryStream())
                            {
                                pictureBoxESHIFTLogo_AdminMyDashboardTab.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                ms.Position = 0;
                                var logo = iTextSharp.text.Image.GetInstance(ms);
                                logo.ScaleAbsolute(100f, 50f);
                                var logoCell = new iTextSharp.text.pdf.PdfPCell(logo);
                                logoCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER;
                                logoCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
                                logoCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_TOP;
                                headerTable.AddCell(logoCell);
                            }
                        }
                        else
                        {
                            var emptyLogoCell = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(" "));
                            emptyLogoCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER;
                            headerTable.AddCell(emptyLogoCell);
                        }

                        var companyInfoCell = new iTextSharp.text.pdf.PdfPCell();
                        companyInfoCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER;
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("E-Shift Company", subTitleFont));
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("123 Business St,\nColombo, Sri Lanka\nPhone: +94 712345678", bodyFont));
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("Report Date: " + DateTime.Now.ToString("yyyy-MM-dd"), bodyFont));
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("Report #: JOBREP-" + DateTime.Now.ToString("yyyyMMddHHmmss"), bodyFont));
                        headerTable.AddCell(companyInfoCell);

                        doc.Add(headerTable);
                        doc.Add(new iTextSharp.text.Paragraph("\n", bodyFont)); 

                        var pdfTable = new iTextSharp.text.pdf.PdfPTable(dataGridViewMain_AdminMyDashboardTab.Columns.Count);
                        pdfTable.WidthPercentage = 100;
                        pdfTable.SpacingBefore = 10f;
                        pdfTable.SpacingAfter = 10f;

                        float[] columnWidths = new float[dataGridViewMain_AdminMyDashboardTab.Columns.Count];
                        for (int i = 0; i < dataGridViewMain_AdminMyDashboardTab.Columns.Count; i++)
                            columnWidths[i] = 4f;
                        pdfTable.SetWidths(columnWidths);

                        foreach (DataGridViewColumn column in dataGridViewMain_AdminMyDashboardTab.Columns)
                        {
                            var headerCell = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(column.HeaderText, boldTableFont));
                            headerCell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                            headerCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                            headerCell.Padding = 5;
                            pdfTable.AddCell(headerCell);
                        }

                        foreach (DataGridViewRow row in dataGridViewMain_AdminMyDashboardTab.Rows)
                        {
                            if (!row.IsNewRow)
                            {
                                foreach (DataGridViewCell cell in row.Cells)
                                {
                                    var dataCell = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(cell.Value?.ToString() ?? "", bodyFont));
                                    dataCell.Padding = 4;
                                    dataCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                                    pdfTable.AddCell(dataCell);
                                }
                            }
                        }

                        doc.Add(pdfTable);

                        var notePara = new iTextSharp.text.Paragraph("Note: For Admin and Company use only", noteFont);
                        notePara.Alignment = iTextSharp.text.Element.ALIGN_LEFT;
                        doc.Add(notePara);

                        doc.Close();
                        writer.Close();

                        MessageBox.Show("Job Report PDF exported successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + sfd.FileName + "\"");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failed to export PDF: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnDriverDetailsRegister_AdminDriverTab_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(ComboBoxDriverTItle_AdminDriverTab.Text) ||
                    string.IsNullOrWhiteSpace(txtDriverFname_AdminDriverTab.Text) ||
                    string.IsNullOrWhiteSpace(txtDriverLName_AdminDriverTab.Text) ||
                    (!RadioButtonDriverMale_AdminDriverTab.Checked && !RadioButtonDriverFemale.Checked) ||
                    string.IsNullOrWhiteSpace(txtDriverAddress_AdminDriverTab.Text) ||
                    string.IsNullOrWhiteSpace(txtDriverPNumber_AdminDriverTab.Text) ||
                    string.IsNullOrWhiteSpace(txtDriverEmail_AdminDriverTab.Text) ||
                    string.IsNullOrWhiteSpace(txtDriverNICNumber_AdminDriverTab.Text) ||
                    string.IsNullOrWhiteSpace(txtDriverLicenseNumber_AdminDriverTab.Text))
                {
                    MessageBox.Show("Please fill out all required fields.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string gender = RadioButtonDriverMale_AdminDriverTab.Checked ? "Male" : "Female";

                byte[] profilePhotoBytes = null;
                if (pictureBoxDriverProfilePhtoto_AdminDriverTab.Image == null)
                {
                    using (OpenFileDialog ofd = new OpenFileDialog())
                    {
                        ofd.Title = "Select Driver Profile Photo";
                        ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp";
                        if (ofd.ShowDialog() == DialogResult.OK)
                        {
                            var img = System.Drawing.Image.FromFile(ofd.FileName);
                            pictureBoxDriverProfilePhtoto_AdminDriverTab.Image = img;
                            LoadDriversGrid();
                        }

                        else
                        {
                            MessageBox.Show("Please select a profile photo.", "Missing Photo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                }
                if (pictureBoxDriverProfilePhtoto_AdminDriverTab.Image != null)
                {
                    using (var ms = new System.IO.MemoryStream())
                    {
                        pictureBoxDriverProfilePhtoto_AdminDriverTab.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                        profilePhotoBytes = ms.ToArray();
                    }
                }

                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string insertQuery = @"INSERT INTO Drivers (Title, FirstName, LastName, Gender, DateOfBirth, Address, PhoneNumber, Email, NIC, LicenseNumber, LicenseExpiryDate, RegisterDate, ProfilePhoto, IsActive)
                        VALUES (@Title, @FirstName, @LastName, @Gender, @DateOfBirth, @Address, @PhoneNumber, @Email, @NIC, @LicenseNumber, @LicenseExpiryDate, @RegisterDate, @ProfilePhoto, @IsActive)";
                    using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@Title", ComboBoxDriverTItle_AdminDriverTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@FirstName", txtDriverFname_AdminDriverTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@LastName", txtDriverLName_AdminDriverTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@Gender", gender);
                        cmd.Parameters.AddWithValue("@DateOfBirth", dateTimePickerDriverDOB_AdminDriverTab.Value.Date);
                        cmd.Parameters.AddWithValue("@Address", txtDriverAddress_AdminDriverTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@PhoneNumber", txtDriverPNumber_AdminDriverTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@Email", txtDriverEmail_AdminDriverTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@NIC", txtDriverNICNumber_AdminDriverTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@LicenseNumber", txtDriverLicenseNumber_AdminDriverTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@LicenseExpiryDate", dateTimePickerDriverLicenseExpiryDate_AdminDriverTab.Value.Date);
                        cmd.Parameters.AddWithValue("@RegisterDate", dateTimePickerDriverRegisteringDate_AdminDriverTab.Value);
                        cmd.Parameters.AddWithValue("@ProfilePhoto", (object)profilePhotoBytes ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@IsActive", 1);
                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                        {
                            MessageBox.Show("Driver registered successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            LoadDriversGrid();
                            PopulateMailingDriverIDComboBox();
                            ComboBoxMailingDriverID_AdminMailingTab.SelectedIndexChanged += ComboBoxMailingDriverID_AdminMailingTab_SelectedIndexChanged;
                            PopulateDriverIDComboBox_AdminJobsTab();
                            ComboBoxDriverID_AdminJobsTab.SelectedIndexChanged += ComboBoxDriverID_AdminJobsTab_SelectedIndexChanged;
                        }
                        else
                        {
                            MessageBox.Show("Failed to register driver.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (SqlException ex) when (ex.Message.Contains("UNIQUE"))
            {
                MessageBox.Show("Email, NIC, or License Number already exists in the system.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error registering driver: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            ComboBoxDriverID_AdminDriverTab.SelectedIndex = -1;
            ComboBoxDriverTItle_AdminDriverTab.SelectedIndex = -1;
            txtDriverFname_AdminDriverTab.Clear();
            txtDriverLName_AdminDriverTab.Clear();
            RadioButtonDriverMale_AdminDriverTab.Checked = false;
            RadioButtonDriverFemale.Checked = false;
            dateTimePickerDriverDOB_AdminDriverTab.Value = DateTime.Now;
            txtDriverAddress_AdminDriverTab.Clear();
            txtDriverPNumber_AdminDriverTab.Clear();
            txtDriverEmail_AdminDriverTab.Clear();
            txtDriverNICNumber_AdminDriverTab.Clear();
            txtDriverLicenseNumber_AdminDriverTab.Clear();
            dateTimePickerDriverLicenseExpiryDate_AdminDriverTab.Value = DateTime.Now;
            dateTimePickerDriverRegisteringDate_AdminDriverTab.Value = DateTime.Now;
            pictureBoxDriverProfilePhtoto_AdminDriverTab.Image = null;
            PopulateDriverIDComboBox();
            btnDriverDetailsRegister_AdminDriverTab.Enabled = true;
        }

        private void pictureBoxDriverProfilePhtoto_AdminDriverTab_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Select Driver Profile Photo";
                ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var img = System.Drawing.Image.FromFile(ofd.FileName);
                        pictureBoxDriverProfilePhtoto_AdminDriverTab.Image = img;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failed to load image: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void PopulateDriverIDComboBox()
        {
            try
            {
                ComboBoxDriverID_AdminDriverTab.Items.Clear();
                ComboBoxDriverID_AdminDriverTab.Items.Add("New register");
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string query = "SELECT DriverID FROM Drivers ORDER BY DriverID";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            ComboBoxDriverID_AdminDriverTab.Items.Add(reader["DriverID"].ToString());
                        }
                    }
                }
                ComboBoxDriverID_AdminDriverTab.SelectedIndex = 0; 
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load Driver IDs: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Enable/disable register button based on ComboBox selection
        private void ComboBoxDriverID_AdminDriverTab_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ComboBoxDriverID_AdminDriverTab.SelectedItem != null && ComboBoxDriverID_AdminDriverTab.SelectedItem.ToString() != "New register")
            {
                // Fetch driver data from DB and populate fields
                int driverId;
                if (int.TryParse(ComboBoxDriverID_AdminDriverTab.SelectedItem.ToString(), out driverId))
                {
                    using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                    {
                        conn.Open();
                        string query = "SELECT * FROM Drivers WHERE DriverID = @DriverID";
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@DriverID", driverId);
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    ComboBoxDriverTItle_AdminDriverTab.Text = reader["Title"].ToString();
                                    txtDriverFname_AdminDriverTab.Text = reader["FirstName"].ToString();
                                    txtDriverLName_AdminDriverTab.Text = reader["LastName"].ToString();
                                    string gender = reader["Gender"].ToString();
                                    RadioButtonDriverMale_AdminDriverTab.Checked = (gender == "Male");
                                    RadioButtonDriverFemale.Checked = (gender == "Female");
                                    if (reader["DateOfBirth"] != DBNull.Value)
                                        dateTimePickerDriverDOB_AdminDriverTab.Value = Convert.ToDateTime(reader["DateOfBirth"]);
                                    txtDriverAddress_AdminDriverTab.Text = reader["Address"].ToString();
                                    txtDriverPNumber_AdminDriverTab.Text = reader["PhoneNumber"].ToString();
                                    txtDriverEmail_AdminDriverTab.Text = reader["Email"].ToString();
                                    txtDriverNICNumber_AdminDriverTab.Text = reader["NIC"].ToString();
                                    txtDriverLicenseNumber_AdminDriverTab.Text = reader["LicenseNumber"].ToString();
                                    if (reader["LicenseExpiryDate"] != DBNull.Value)
                                        dateTimePickerDriverLicenseExpiryDate_AdminDriverTab.Value = Convert.ToDateTime(reader["LicenseExpiryDate"]);
                                    if (reader["RegisterDate"] != DBNull.Value)
                                        dateTimePickerDriverRegisteringDate_AdminDriverTab.Value = Convert.ToDateTime(reader["RegisterDate"]);
                                    // Profile photo
                                    if (reader["ProfilePhoto"] != DBNull.Value)
                                    {
                                        byte[] imgBytes = (byte[])reader["ProfilePhoto"];
                                        using (var ms = new System.IO.MemoryStream(imgBytes))
                                        {
                                            pictureBoxDriverProfilePhtoto_AdminDriverTab.Image = System.Drawing.Image.FromStream(ms);
                                        }
                                    }
                                    else
                                    {
                                        pictureBoxDriverProfilePhtoto_AdminDriverTab.Image = null;
                                    }
                                }
                            }
                        }
                    }
                }
                btnDriverDetailsRegister_AdminDriverTab.Enabled = false;
            }
            else
            {
                ComboBoxDriverTItle_AdminDriverTab.SelectedIndex = -1;
                txtDriverFname_AdminDriverTab.Clear();
                txtDriverLName_AdminDriverTab.Clear();
                RadioButtonDriverMale_AdminDriverTab.Checked = false;
                RadioButtonDriverFemale.Checked = false;
                dateTimePickerDriverDOB_AdminDriverTab.Value = DateTime.Now;
                txtDriverAddress_AdminDriverTab.Clear();
                txtDriverPNumber_AdminDriverTab.Clear();
                txtDriverEmail_AdminDriverTab.Clear();
                txtDriverNICNumber_AdminDriverTab.Clear();
                txtDriverLicenseNumber_AdminDriverTab.Clear();
                dateTimePickerDriverLicenseExpiryDate_AdminDriverTab.Value = DateTime.Now;
                dateTimePickerDriverRegisteringDate_AdminDriverTab.Value = DateTime.Now;
                pictureBoxDriverProfilePhtoto_AdminDriverTab.Image = null;
                btnDriverDetailsRegister_AdminDriverTab.Enabled = true;
            }
        }

        private void ComboBoxDriverID_AdminDriverTab_SelectedIndexChanged_1(object sender, EventArgs e)
        {
        }

        private void btnDriverDetailsUpdate_AdminDriverTab_Click(object sender, EventArgs e)
        {
            if (ComboBoxDriverID_AdminDriverTab.SelectedItem == null || ComboBoxDriverID_AdminDriverTab.SelectedItem.ToString() == "New register")
            {
                MessageBox.Show("Please select a Driver ID to update.", "No Driver Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            int driverId;
            if (!int.TryParse(ComboBoxDriverID_AdminDriverTab.SelectedItem.ToString(), out driverId))
            {
                MessageBox.Show("Invalid Driver ID selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                // Fetch current DB values for the selected driver
                Dictionary<string, object> dbValues = new Dictionary<string, object>();
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string selectQuery = "SELECT * FROM Drivers WHERE DriverID = @DriverID";
                    using (SqlCommand selectCmd = new SqlCommand(selectQuery, conn))
                    {
                        selectCmd.Parameters.AddWithValue("@DriverID", driverId);
                        using (SqlDataReader reader = selectCmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                dbValues["Title"] = reader["Title"].ToString();
                                dbValues["FirstName"] = reader["FirstName"].ToString();
                                dbValues["LastName"] = reader["LastName"].ToString();
                                dbValues["Gender"] = reader["Gender"].ToString();
                                dbValues["DateOfBirth"] = reader["DateOfBirth"] == DBNull.Value ? null : (object)Convert.ToDateTime(reader["DateOfBirth"]);
                                dbValues["Address"] = reader["Address"].ToString();
                                dbValues["PhoneNumber"] = reader["PhoneNumber"].ToString();
                                dbValues["Email"] = reader["Email"].ToString();
                                dbValues["NIC"] = reader["NIC"].ToString();
                                dbValues["LicenseNumber"] = reader["LicenseNumber"].ToString();
                                dbValues["LicenseExpiryDate"] = reader["LicenseExpiryDate"] == DBNull.Value ? null : (object)Convert.ToDateTime(reader["LicenseExpiryDate"]);
                                dbValues["RegisterDate"] = reader["RegisterDate"] == DBNull.Value ? null : (object)Convert.ToDateTime(reader["RegisterDate"]);
                                dbValues["ProfilePhoto"] = reader["ProfilePhoto"] == DBNull.Value ? null : (object)reader["ProfilePhoto"];
                            }
                        }
                    }
                }

                // Prepare new values from form
                string gender = RadioButtonDriverMale_AdminDriverTab.Checked ? "Male" : "Female";
                byte[] profilePhotoBytes = null;
                if (pictureBoxDriverProfilePhtoto_AdminDriverTab.Image != null)
                {
                    // Clone the image to avoid file lock issues
                    using (var clonedImg = new Bitmap(pictureBoxDriverProfilePhtoto_AdminDriverTab.Image))
                    using (var ms = new System.IO.MemoryStream())
                    {
                        clonedImg.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                        profilePhotoBytes = ms.ToArray();
                    }
                }

                // Compare and build update query only for changed fields
                var updates = new List<string>();
                var parameters = new List<SqlParameter>();
                if (ComboBoxDriverTItle_AdminDriverTab.Text.Trim() != (dbValues["Title"]?.ToString() ?? ""))
                {
                    updates.Add("Title = @Title");
                    parameters.Add(new SqlParameter("@Title", ComboBoxDriverTItle_AdminDriverTab.Text.Trim()));
                }
                if (txtDriverFname_AdminDriverTab.Text.Trim() != (dbValues["FirstName"]?.ToString() ?? ""))
                {
                    updates.Add("FirstName = @FirstName");
                    parameters.Add(new SqlParameter("@FirstName", txtDriverFname_AdminDriverTab.Text.Trim()));
                }
                if (txtDriverLName_AdminDriverTab.Text.Trim() != (dbValues["LastName"]?.ToString() ?? ""))
                {
                    updates.Add("LastName = @LastName");
                    parameters.Add(new SqlParameter("@LastName", txtDriverLName_AdminDriverTab.Text.Trim()));
                }
                if (gender != (dbValues["Gender"]?.ToString() ?? ""))
                {
                    updates.Add("Gender = @Gender");
                    parameters.Add(new SqlParameter("@Gender", gender));
                }
                if (dateTimePickerDriverDOB_AdminDriverTab.Value.Date != (dbValues["DateOfBirth"] is DateTime dt ? dt.Date : DateTime.MinValue))
                {
                    updates.Add("DateOfBirth = @DateOfBirth");
                    parameters.Add(new SqlParameter("@DateOfBirth", dateTimePickerDriverDOB_AdminDriverTab.Value.Date));
                }
                if (txtDriverAddress_AdminDriverTab.Text.Trim() != (dbValues["Address"]?.ToString() ?? ""))
                {
                    updates.Add("Address = @Address");
                    parameters.Add(new SqlParameter("@Address", txtDriverAddress_AdminDriverTab.Text.Trim()));
                }
                if (txtDriverPNumber_AdminDriverTab.Text.Trim() != (dbValues["PhoneNumber"]?.ToString() ?? ""))
                {
                    updates.Add("PhoneNumber = @PhoneNumber");
                    parameters.Add(new SqlParameter("@PhoneNumber", txtDriverPNumber_AdminDriverTab.Text.Trim()));
                }
                if (txtDriverEmail_AdminDriverTab.Text.Trim() != (dbValues["Email"]?.ToString() ?? ""))
                {
                    updates.Add("Email = @Email");
                    parameters.Add(new SqlParameter("@Email", txtDriverEmail_AdminDriverTab.Text.Trim()));
                }
                if (txtDriverNICNumber_AdminDriverTab.Text.Trim() != (dbValues["NIC"]?.ToString() ?? ""))
                {
                    updates.Add("NIC = @NIC");
                    parameters.Add(new SqlParameter("@NIC", txtDriverNICNumber_AdminDriverTab.Text.Trim()));
                }
                if (txtDriverLicenseNumber_AdminDriverTab.Text.Trim() != (dbValues["LicenseNumber"]?.ToString() ?? ""))
                {
                    updates.Add("LicenseNumber = @LicenseNumber");
                    parameters.Add(new SqlParameter("@LicenseNumber", txtDriverLicenseNumber_AdminDriverTab.Text.Trim()));
                }
                if (dateTimePickerDriverLicenseExpiryDate_AdminDriverTab.Value.Date != (dbValues["LicenseExpiryDate"] is DateTime dt2 ? dt2.Date : DateTime.MinValue))
                {
                    updates.Add("LicenseExpiryDate = @LicenseExpiryDate");
                    parameters.Add(new SqlParameter("@LicenseExpiryDate", dateTimePickerDriverLicenseExpiryDate_AdminDriverTab.Value.Date));
                }
                if (dateTimePickerDriverRegisteringDate_AdminDriverTab.Value != (dbValues["RegisterDate"] is DateTime dt3 ? dt3 : DateTime.MinValue))
                {
                    updates.Add("RegisterDate = @RegisterDate");
                    parameters.Add(new SqlParameter("@RegisterDate", dateTimePickerDriverRegisteringDate_AdminDriverTab.Value));
                }
                // Profile photo: compare byte arrays
                bool photoChanged = false;
                if (profilePhotoBytes != null)
                {
                    var dbPhoto = dbValues["ProfilePhoto"] as byte[];
                    if (dbPhoto == null || !profilePhotoBytes.SequenceEqual(dbPhoto))
                    {
                        photoChanged = true;
                    }
                }
                else if (dbValues["ProfilePhoto"] != null)
                {
                    photoChanged = true;
                }
                if (photoChanged)
                {
                    updates.Add("ProfilePhoto = @ProfilePhoto");
                    parameters.Add(new SqlParameter("@ProfilePhoto", (object)profilePhotoBytes ?? DBNull.Value));
                }

                if (updates.Count == 0)
                {
                    MessageBox.Show("No changes detected. Please modify at least one field to update.", "No Update", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string updateQuery = $"UPDATE Drivers SET {string.Join(", ", updates)} WHERE DriverID = @DriverID";
                    using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                    {
                        foreach (var param in parameters)
                            cmd.Parameters.Add(param);
                        cmd.Parameters.AddWithValue("@DriverID", driverId);
                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                        {
                            MessageBox.Show("Driver details updated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            LoadDriversGrid();
                        }
                        else
                        {
                            MessageBox.Show("No changes were made. Please check the details and try again.", "No Update", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
            }
            catch (SqlException ex) when (ex.Message.Contains("UNIQUE"))
            {
                MessageBox.Show("Email, NIC, or License Number already exists in the system.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error updating driver: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDriverDetailsDelete_AdminDriverTab_Click(object sender, EventArgs e)
        {
            if (ComboBoxDriverID_AdminDriverTab.SelectedItem == null || ComboBoxDriverID_AdminDriverTab.SelectedItem.ToString() == "New register")
            {
                MessageBox.Show("Please select a Driver ID to delete.", "No Driver Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            int driverId;
            if (!int.TryParse(ComboBoxDriverID_AdminDriverTab.SelectedItem.ToString(), out driverId))
            {
                MessageBox.Show("Invalid Driver ID selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var confirm = MessageBox.Show($"Are you sure you want to delete Driver ID {driverId}? This action cannot be undone.", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (confirm != DialogResult.Yes)
                return;
            try
            {
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string deleteQuery = "DELETE FROM Drivers WHERE DriverID = @DriverID";
                    using (SqlCommand cmd = new SqlCommand(deleteQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@DriverID", driverId);
                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                        {
                            MessageBox.Show("Driver deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            LoadDriversGrid();
                        }
                        else
                        {
                            MessageBox.Show("No driver found to delete. Please check the Driver ID.", "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                }
                ComboBoxDriverID_AdminDriverTab.SelectedIndex = -1;
                ComboBoxDriverTItle_AdminDriverTab.SelectedIndex = -1;
                txtDriverFname_AdminDriverTab.Clear();
                txtDriverLName_AdminDriverTab.Clear();
                RadioButtonDriverMale_AdminDriverTab.Checked = false;
                RadioButtonDriverFemale.Checked = false;
                dateTimePickerDriverDOB_AdminDriverTab.Value = DateTime.Now;
                txtDriverAddress_AdminDriverTab.Clear();
                txtDriverPNumber_AdminDriverTab.Clear();
                txtDriverEmail_AdminDriverTab.Clear();
                txtDriverNICNumber_AdminDriverTab.Clear();
                txtDriverLicenseNumber_AdminDriverTab.Clear();
                dateTimePickerDriverLicenseExpiryDate_AdminDriverTab.Value = DateTime.Now;
                dateTimePickerDriverRegisteringDate_AdminDriverTab.Value = DateTime.Now;
                pictureBoxDriverProfilePhtoto_AdminDriverTab.Image = null;
                PopulateDriverIDComboBox();
                btnDriverDetailsRegister_AdminDriverTab.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error deleting driver: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDriverDetailsClear_AdminDriverTab_Click(object sender, EventArgs e)
        {
            ComboBoxDriverID_AdminDriverTab.SelectedIndex = -1;
            ComboBoxDriverTItle_AdminDriverTab.SelectedIndex = -1;
            txtDriverFname_AdminDriverTab.Clear();
            txtDriverLName_AdminDriverTab.Clear();
            RadioButtonDriverMale_AdminDriverTab.Checked = false;
            RadioButtonDriverFemale.Checked = false;
            dateTimePickerDriverDOB_AdminDriverTab.Value = DateTime.Now;
            txtDriverAddress_AdminDriverTab.Clear();
            txtDriverPNumber_AdminDriverTab.Clear();
            txtDriverEmail_AdminDriverTab.Clear();
            txtDriverNICNumber_AdminDriverTab.Clear();
            txtDriverLicenseNumber_AdminDriverTab.Clear();
            dateTimePickerDriverLicenseExpiryDate_AdminDriverTab.Value = DateTime.Now;
            dateTimePickerDriverRegisteringDate_AdminDriverTab.Value = DateTime.Now;
            pictureBoxDriverProfilePhtoto_AdminDriverTab.Image = null;
            PopulateDriverIDComboBox();
            btnDriverDetailsRegister_AdminDriverTab.Enabled = true;
        }

        private void LoadDriversGrid()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string query = "SELECT DriverID, Title, FirstName, LastName, Gender, DateOfBirth, Address, PhoneNumber, Email, NIC, LicenseNumber, LicenseExpiryDate, RegisterDate, IsActive FROM Drivers ORDER BY DriverID";
                    using (SqlDataAdapter da = new SqlDataAdapter(query, conn))
                    {
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dataGridViewDriverList_AdminDriverTab.DataSource = dt;
                        dataGridViewDriverList_AdminDriverTab.ClearSelection();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load drivers: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDriverListReportPrint_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "PDF files (*.pdf)|*.pdf";
                sfd.Title = "Save Drivers Report as PDF";
                sfd.FileName = "DriversReport.pdf";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 40f, 40f, 60f, 50f);
                        iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, new System.IO.FileStream(sfd.FileName, System.IO.FileMode.Create));
                        doc.Open();

                        var titleFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 18);
                        var subTitleFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 12);
                        var bodyFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA, 9);
                        var boldTableFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 8);
                        var endingFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_OBLIQUE, 8);
                        var noteFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.RED);
                        var blueTitleFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 16, iTextSharp.text.Font.BOLD, new iTextSharp.text.BaseColor(0, 102, 204));

                        iTextSharp.text.pdf.PdfPTable headerTable = new iTextSharp.text.pdf.PdfPTable(2);
                        headerTable.WidthPercentage = 100;
                        headerTable.SetWidths(new float[] { 20, 80 });

                        // Use logo  pictureBoxESHIFTLogo_AdminMyDashboardTab if available
                        if (pictureBoxESHIFTLogo_AdminMyDashboardTab != null && pictureBoxESHIFTLogo_AdminMyDashboardTab.Image != null)
                        {
                            using (var ms = new System.IO.MemoryStream())
                            {
                                pictureBoxESHIFTLogo_AdminMyDashboardTab.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                ms.Position = 0;
                                var logo = iTextSharp.text.Image.GetInstance(ms);
                                logo.ScaleAbsolute(100f, 50f);
                                var logoCell = new iTextSharp.text.pdf.PdfPCell(logo);
                                logoCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER;
                                logoCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
                                logoCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_TOP;
                                headerTable.AddCell(logoCell);
                            }
                        }
                        else
                        {
                            var emptyLogoCell = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(" "));
                            emptyLogoCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER;
                            headerTable.AddCell(emptyLogoCell);
                        }

                        // Company Info
                        var companyInfoCell = new iTextSharp.text.pdf.PdfPCell();
                        companyInfoCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER;
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("E-Shift Company", subTitleFont));
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("123 Business St,\nColombo, Sri Lanka\nPhone: +94 712345678", bodyFont));
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("Report Date: " + DateTime.Now.ToString("yyyy-MM-dd"), bodyFont));
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("Report #: DRIVREP-" + DateTime.Now.ToString("yyyyMMddHHmmss"), bodyFont));
                        headerTable.AddCell(companyInfoCell);

                        doc.Add(headerTable);
                        doc.Add(new iTextSharp.text.Paragraph("\n", bodyFont)); // Spacing

                        // === Blue Drivers Report Title ===
                        var driversReportTitle = new iTextSharp.text.Paragraph("Drivers Report", blueTitleFont);
                        driversReportTitle.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                        doc.Add(driversReportTitle);
                        doc.Add(new iTextSharp.text.Paragraph("\n", bodyFont)); // Spacing

                        // === Table Section ===
                        var pdfTable = new iTextSharp.text.pdf.PdfPTable(dataGridViewDriverList_AdminDriverTab.Columns.Count);
                        pdfTable.WidthPercentage = 100;
                        pdfTable.SpacingBefore = 10f;
                        pdfTable.SpacingAfter = 10f;

                        float[] columnWidths = new float[dataGridViewDriverList_AdminDriverTab.Columns.Count];
                        for (int i = 0; i < dataGridViewDriverList_AdminDriverTab.Columns.Count; i++)
                            columnWidths[i] = 4f;
                        pdfTable.SetWidths(columnWidths);

                        // Table Headers
                        foreach (DataGridViewColumn column in dataGridViewDriverList_AdminDriverTab.Columns)
                        {
                            var headerCell = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(column.HeaderText, boldTableFont));
                            headerCell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                            headerCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                            headerCell.Padding = 5;
                            pdfTable.AddCell(headerCell);
                        }

                        // Table Rows
                        foreach (DataGridViewRow row in dataGridViewDriverList_AdminDriverTab.Rows)
                        {
                            if (!row.IsNewRow)
                            {
                                foreach (DataGridViewCell cell in row.Cells)
                                {
                                    var dataCell = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(cell.Value?.ToString() ?? "", bodyFont));
                                    dataCell.Padding = 4;
                                    dataCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                                    pdfTable.AddCell(dataCell);
                                }
                            }
                        }

                        doc.Add(pdfTable);

                        // === Note for Admin/Company Use Only ===
                        var notePara = new iTextSharp.text.Paragraph("Note: For Admin and Company use only", noteFont);
                        notePara.Alignment = iTextSharp.text.Element.ALIGN_LEFT;
                        doc.Add(notePara);

                        doc.Close();
                        writer.Close();

                        MessageBox.Show("Drivers Report PDF exported successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + sfd.FileName + "\"");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failed to export PDF: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void pictureBoxAssistantProfilePhoto_AdminAssistantTab_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Select Assistant Profile Photo";
                ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var img = System.Drawing.Image.FromFile(ofd.FileName);
                        pictureBoxAssistantProfilePhoto_AdminAssistantTab.Image = img;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failed to load image: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void PopulateAssistantIDComboBox()
        {
            try
            {
                ComboBoxAssistantID_AdminAssistantTab.Items.Clear();
                ComboBoxAssistantID_AdminAssistantTab.Items.Add("New register");
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string query = "SELECT AssistantID FROM Assistants ORDER BY AssistantID";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            ComboBoxAssistantID_AdminAssistantTab.Items.Add(reader["AssistantID"].ToString());
                        }
                    }
                }
                ComboBoxAssistantID_AdminAssistantTab.SelectedIndex = 0; // Default to 'New register'
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load Assistant IDs: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ComboBoxAssistantID_AdminAssistantTab_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ComboBoxAssistantID_AdminAssistantTab.SelectedItem != null && ComboBoxAssistantID_AdminAssistantTab.SelectedItem.ToString() != "New register")
            {
                int assistantId;
                if (int.TryParse(ComboBoxAssistantID_AdminAssistantTab.SelectedItem.ToString(), out assistantId))
                {
                    using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                    {
                        conn.Open();
                        string query = "SELECT * FROM Assistants WHERE AssistantID = @AssistantID";
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@AssistantID", assistantId);
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    ComboBoxAssistantTitle_AdminAssistantTab.Text = reader["Title"].ToString();
                                    txtAssistantFName_AdminAssistantTab.Text = reader["FirstName"].ToString();
                                    txtAssistantLName_AdminAssistantTab.Text = reader["LastName"].ToString();
                                    string gender = reader["Gender"].ToString();
                                    RadioButtonAssistantMale_AdminAssistantTab.Checked = (gender == "Male");
                                    RadioButtonAssistantFemale_AdminAssistantTab.Checked = (gender == "Female");
                                    if (reader["DateOfBirth"] != DBNull.Value)
                                        dateTimePickerAssistantDOB_AdminAssistantTab.Value = Convert.ToDateTime(reader["DateOfBirth"]);
                                    txtAssistantAdress_AdminAssistantTab.Text = reader["Address"].ToString();
                                    txtAssistantPNumber_AdminAssistantTab.Text = reader["PhoneNumber"].ToString();
                                    txtAssistantEmail_AdminAssistantTab.Text = reader["Email"].ToString();
                                    lblAssistantNICNumber_AdminAssistantTab.Text = reader["NIC"].ToString();
                                    btnAssistantDetailsRegister_AdminAssistantTab.Enabled = false;
                                    if (reader["RegisterDate"] != DBNull.Value)
                                        dateTimePickerAssistantRegisteringDate_AdminAssistantTab.Value = Convert.ToDateTime(reader["RegisterDate"]);
                                    // Profile photo
                                    if (reader["ProfilePhoto"] != DBNull.Value)
                                    {
                                        byte[] imgBytes = (byte[])reader["ProfilePhoto"];
                                        using (var ms = new System.IO.MemoryStream(imgBytes))
                                        {
                                            pictureBoxAssistantProfilePhoto_AdminAssistantTab.Image = System.Drawing.Image.FromStream(ms);
                                        }
                                    }
                                    else
                                    {
                                        pictureBoxAssistantProfilePhoto_AdminAssistantTab.Image = null;
                                    }
                                }
                            }
                        }
                    }
                }

            }
            else
            {
                ComboBoxAssistantTitle_AdminAssistantTab.SelectedIndex = -1;
                txtAssistantFName_AdminAssistantTab.Clear();
                txtAssistantLName_AdminAssistantTab.Clear();
                RadioButtonAssistantMale_AdminAssistantTab.Checked = false;
                RadioButtonAssistantFemale_AdminAssistantTab.Checked = false;
                dateTimePickerAssistantDOB_AdminAssistantTab.Value = DateTime.Now;
                txtAssistantAdress_AdminAssistantTab.Clear();
                txtAssistantPNumber_AdminAssistantTab.Clear();
                txtAssistantEmail_AdminAssistantTab.Clear();
                lblAssistantNICNumber_AdminAssistantTab.Text = string.Empty;
                dateTimePickerAssistantRegisteringDate_AdminAssistantTab.Value = DateTime.Now;
                pictureBoxAssistantProfilePhoto_AdminAssistantTab.Image = null;
                btnAssistantDetailsRegister_AdminAssistantTab.Enabled = true;
            }
        }

        private void btnAssistantDetailsRegister_AdminAssistantTab_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(ComboBoxAssistantTitle_AdminAssistantTab.Text) ||
                    string.IsNullOrWhiteSpace(txtAssistantFName_AdminAssistantTab.Text) ||
                    string.IsNullOrWhiteSpace(txtAssistantLName_AdminAssistantTab.Text) ||
                    (!RadioButtonAssistantMale_AdminAssistantTab.Checked && !RadioButtonAssistantFemale_AdminAssistantTab.Checked) ||
                    string.IsNullOrWhiteSpace(txtAssistantAdress_AdminAssistantTab.Text) ||
                    string.IsNullOrWhiteSpace(txtAssistantPNumber_AdminAssistantTab.Text) ||
                    string.IsNullOrWhiteSpace(txtAssistantEmail_AdminAssistantTab.Text) ||
                    string.IsNullOrWhiteSpace(lblAssistantNICNumber_AdminAssistantTab.Text))
                {
                    MessageBox.Show("Please fill out all required fields.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string gender = RadioButtonAssistantMale_AdminAssistantTab.Checked ? "Male" : "Female";

                byte[] profilePhotoBytes = null;
                if (pictureBoxAssistantProfilePhoto_AdminAssistantTab.Image == null)
                {
                    using (OpenFileDialog ofd = new OpenFileDialog())
                    {
                        ofd.Title = "Select Assistant Profile Photo";
                        ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp";
                        if (ofd.ShowDialog() == DialogResult.OK)
                        {
                            var img = System.Drawing.Image.FromFile(ofd.FileName);
                            pictureBoxAssistantProfilePhoto_AdminAssistantTab.Image = img;
                        }
                        else
                        {
                            MessageBox.Show("Please select a profile photo.", "Missing Photo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                }
                if (pictureBoxAssistantProfilePhoto_AdminAssistantTab.Image != null)
                {
                    using (var ms = new System.IO.MemoryStream())
                    {
                        pictureBoxAssistantProfilePhoto_AdminAssistantTab.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                        profilePhotoBytes = ms.ToArray();
                    }
                }

                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string insertQuery = @"INSERT INTO Assistants (Title, FirstName, LastName, Gender, DateOfBirth, Address, PhoneNumber, Email, NIC, RegisterDate, ProfilePhoto, IsActive)
                        VALUES (@Title, @FirstName, @LastName, @Gender, @DateOfBirth, @Address, @PhoneNumber, @Email, @NIC, @RegisterDate, @ProfilePhoto, @IsActive)";
                    using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@Title", ComboBoxAssistantTitle_AdminAssistantTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@FirstName", txtAssistantFName_AdminAssistantTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@LastName", txtAssistantLName_AdminAssistantTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@Gender", gender);
                        cmd.Parameters.AddWithValue("@DateOfBirth", dateTimePickerAssistantDOB_AdminAssistantTab.Value.Date);
                        cmd.Parameters.AddWithValue("@Address", txtAssistantAdress_AdminAssistantTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@PhoneNumber", txtAssistantPNumber_AdminAssistantTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@Email", txtAssistantEmail_AdminAssistantTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@NIC", lblAssistantNICNumber_AdminAssistantTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@RegisterDate", dateTimePickerAssistantRegisteringDate_AdminAssistantTab.Value);
                        cmd.Parameters.AddWithValue("@ProfilePhoto", (object)profilePhotoBytes ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@IsActive", 1);
                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                        {
                            MessageBox.Show("Assistant registered successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            PopulateAssistantIDComboBox();
                            LoadAssistantsGrid();
                            PopulateMailingAssistantIDComboBox();
                            PopulateAssistantIDComboBox_AdminJobsTab();
                            ComboBoxAssistantID_AdminJobsTab.SelectedIndexChanged += ComboBoxAssistantID_AdminJobsTab_SelectedIndexChanged;
                        }
                        else
                        {
                            MessageBox.Show("Failed to register assistant.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (SqlException ex) when (ex.Message.Contains("UNIQUE"))
            {
                MessageBox.Show("Email or NIC already exists in the system.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error registering assistant: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAssistantDetailsClear_AdminAssistantTab_Click(object sender, EventArgs e)
        {
            ComboBoxAssistantTitle_AdminAssistantTab.SelectedIndex = -1;
            txtAssistantFName_AdminAssistantTab.Clear();
            txtAssistantLName_AdminAssistantTab.Clear();
            RadioButtonAssistantMale_AdminAssistantTab.Checked = false;
            RadioButtonAssistantFemale_AdminAssistantTab.Checked = false;
            dateTimePickerAssistantDOB_AdminAssistantTab.Value = DateTime.Now;
            txtAssistantAdress_AdminAssistantTab.Clear();
            txtAssistantPNumber_AdminAssistantTab.Clear();
            txtAssistantEmail_AdminAssistantTab.Clear();
            lblAssistantNICNumber_AdminAssistantTab.Text = string.Empty;
            dateTimePickerAssistantRegisteringDate_AdminAssistantTab.Value = DateTime.Now;
            pictureBoxAssistantProfilePhoto_AdminAssistantTab.Image = null;
            btnAssistantDetailsRegister_AdminAssistantTab.Enabled = true;
            PopulateAssistantIDComboBox();
            if (ComboBoxAssistantID_AdminAssistantTab.Items.Count > 0)
                ComboBoxAssistantID_AdminAssistantTab.SelectedIndex = 0;

        }

        private void btnAssistantDetailsUpdate_AdminAssistantTab_Click(object sender, EventArgs e)
        {
            if (ComboBoxAssistantID_AdminAssistantTab.SelectedItem == null || ComboBoxAssistantID_AdminAssistantTab.SelectedItem.ToString() == "New register")
            {
                MessageBox.Show("Please select an Assistant ID to update.", "No Assistant Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            int assistantId;
            if (!int.TryParse(ComboBoxAssistantID_AdminAssistantTab.SelectedItem.ToString(), out assistantId))
            {
                MessageBox.Show("Invalid Assistant ID selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                Dictionary<string, object> dbValues = new Dictionary<string, object>();
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string selectQuery = "SELECT * FROM Assistants WHERE AssistantID = @AssistantID";
                    using (SqlCommand selectCmd = new SqlCommand(selectQuery, conn))
                    {
                        selectCmd.Parameters.AddWithValue("@AssistantID", assistantId);
                        using (SqlDataReader reader = selectCmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                dbValues["Title"] = reader["Title"].ToString();
                                dbValues["FirstName"] = reader["FirstName"].ToString();
                                dbValues["LastName"] = reader["LastName"].ToString();
                                dbValues["Gender"] = reader["Gender"].ToString();
                                dbValues["DateOfBirth"] = reader["DateOfBirth"] == DBNull.Value ? null : (object)Convert.ToDateTime(reader["DateOfBirth"]);
                                dbValues["Address"] = reader["Address"].ToString();
                                dbValues["PhoneNumber"] = reader["PhoneNumber"].ToString();
                                dbValues["Email"] = reader["Email"].ToString();
                                dbValues["NIC"] = reader["NIC"].ToString();
                                dbValues["RegisterDate"] = reader["RegisterDate"] == DBNull.Value ? null : (object)Convert.ToDateTime(reader["RegisterDate"]);
                                dbValues["ProfilePhoto"] = reader["ProfilePhoto"] == DBNull.Value ? null : (object)reader["ProfilePhoto"];
                            }
                        }
                    }
                }

                string gender = RadioButtonAssistantMale_AdminAssistantTab.Checked ? "Male" : "Female";
                byte[] profilePhotoBytes = null;
                if (pictureBoxAssistantProfilePhoto_AdminAssistantTab.Image != null)
                {
                    using (var clonedImg = new Bitmap(pictureBoxAssistantProfilePhoto_AdminAssistantTab.Image))
                    using (var ms = new System.IO.MemoryStream())
                    {
                        clonedImg.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                        profilePhotoBytes = ms.ToArray();
                    }
                }

                var updates = new List<string>();
                var parameters = new List<SqlParameter>();
                if (ComboBoxAssistantTitle_AdminAssistantTab.Text.Trim() != (dbValues["Title"]?.ToString() ?? ""))
                {
                    updates.Add("Title = @Title");
                    parameters.Add(new SqlParameter("@Title", ComboBoxAssistantTitle_AdminAssistantTab.Text.Trim()));
                }
                if (txtAssistantFName_AdminAssistantTab.Text.Trim() != (dbValues["FirstName"]?.ToString() ?? ""))
                {
                    updates.Add("FirstName = @FirstName");
                    parameters.Add(new SqlParameter("@FirstName", txtAssistantFName_AdminAssistantTab.Text.Trim()));
                }
                if (txtAssistantLName_AdminAssistantTab.Text.Trim() != (dbValues["LastName"]?.ToString() ?? ""))
                {
                    updates.Add("LastName = @LastName");
                    parameters.Add(new SqlParameter("@LastName", txtAssistantLName_AdminAssistantTab.Text.Trim()));
                }
                if (gender != (dbValues["Gender"]?.ToString() ?? ""))
                {
                    updates.Add("Gender = @Gender");
                    parameters.Add(new SqlParameter("@Gender", gender));
                }
                if (dateTimePickerAssistantDOB_AdminAssistantTab.Value.Date != (dbValues["DateOfBirth"] is DateTime dt ? dt.Date : DateTime.MinValue))
                {
                    updates.Add("DateOfBirth = @DateOfBirth");
                    parameters.Add(new SqlParameter("@DateOfBirth", dateTimePickerAssistantDOB_AdminAssistantTab.Value.Date));
                }
                if (txtAssistantAdress_AdminAssistantTab.Text.Trim() != (dbValues["Address"]?.ToString() ?? ""))
                {
                    updates.Add("Address = @Address");
                    parameters.Add(new SqlParameter("@Address", txtAssistantAdress_AdminAssistantTab.Text.Trim()));
                }
                if (txtAssistantPNumber_AdminAssistantTab.Text.Trim() != (dbValues["PhoneNumber"]?.ToString() ?? ""))
                {
                    updates.Add("PhoneNumber = @PhoneNumber");
                    parameters.Add(new SqlParameter("@PhoneNumber", txtAssistantPNumber_AdminAssistantTab.Text.Trim()));
                }
                if (txtAssistantEmail_AdminAssistantTab.Text.Trim() != (dbValues["Email"]?.ToString() ?? ""))
                {
                    updates.Add("Email = @Email");
                    parameters.Add(new SqlParameter("@Email", txtAssistantEmail_AdminAssistantTab.Text.Trim()));
                }
                if (lblAssistantNICNumber_AdminAssistantTab.Text.Trim() != (dbValues["NIC"]?.ToString() ?? ""))
                {
                    updates.Add("NIC = @NIC");
                    parameters.Add(new SqlParameter("@NIC", lblAssistantNICNumber_AdminAssistantTab.Text.Trim()));
                }
                if (dateTimePickerAssistantRegisteringDate_AdminAssistantTab.Value != (dbValues["RegisterDate"] is DateTime dt2 ? dt2 : DateTime.MinValue))
                {
                    updates.Add("RegisterDate = @RegisterDate");
                    parameters.Add(new SqlParameter("@RegisterDate", dateTimePickerAssistantRegisteringDate_AdminAssistantTab.Value));
                }
                bool photoChanged = false;
                if (profilePhotoBytes != null)
                {
                    var dbPhoto = dbValues["ProfilePhoto"] as byte[];
                    if (dbPhoto == null || !profilePhotoBytes.SequenceEqual(dbPhoto))
                    {
                        photoChanged = true;
                    }
                }
                else if (dbValues["ProfilePhoto"] != null)
                {
                    photoChanged = true;
                }
                if (photoChanged)
                {
                    updates.Add("ProfilePhoto = @ProfilePhoto");
                    parameters.Add(new SqlParameter("@ProfilePhoto", (object)profilePhotoBytes ?? DBNull.Value));
                }

                if (updates.Count == 0)
                {
                    MessageBox.Show("No changes detected. Please modify at least one field to update.", "No Update", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string updateQuery = $"UPDATE Assistants SET {string.Join(", ", updates)} WHERE AssistantID = @AssistantID";
                    using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                    {
                        foreach (var param in parameters)
                            cmd.Parameters.Add(param);
                        cmd.Parameters.AddWithValue("@AssistantID", assistantId);
                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                        {
                            MessageBox.Show("Assistant details updated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            PopulateAssistantIDComboBox();
                            LoadAssistantsGrid();
                        }
                        else
                        {
                            MessageBox.Show("No changes were made. Please check the details and try again.", "No Update", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
            }
            catch (SqlException ex) when (ex.Message.Contains("UNIQUE"))
            {
                MessageBox.Show("Email or NIC already exists in the system.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error updating assistant: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btenAssistantDetailsDelete_AdminAssistantTab_Click(object sender, EventArgs e)
        {
            if (ComboBoxAssistantID_AdminAssistantTab.SelectedItem == null || ComboBoxAssistantID_AdminAssistantTab.SelectedItem.ToString() == "New register")
            {
                MessageBox.Show("Please select an Assistant ID to delete.", "No Assistant Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            int assistantId;
            if (!int.TryParse(ComboBoxAssistantID_AdminAssistantTab.SelectedItem.ToString(), out assistantId))
            {
                MessageBox.Show("Invalid Assistant ID selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var confirm = MessageBox.Show($"Are you sure you want to delete Assistant ID {assistantId}? This action cannot be undone.", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (confirm != DialogResult.Yes)
                return;
            try
            {
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string deleteQuery = "DELETE FROM Assistants WHERE AssistantID = @AssistantID";
                    using (SqlCommand cmd = new SqlCommand(deleteQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@AssistantID", assistantId);
                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                        {
                            MessageBox.Show("Assistant deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            PopulateAssistantIDComboBox();
                            LoadAssistantsGrid();
                        }
                        else
                        {
                            MessageBox.Show("No assistant found to delete. Please check the Assistant ID.", "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                }
                ComboBoxAssistantID_AdminAssistantTab.SelectedIndex = -1;
                ComboBoxAssistantTitle_AdminAssistantTab.SelectedIndex = -1;
                txtAssistantFName_AdminAssistantTab.Clear();
                txtAssistantLName_AdminAssistantTab.Clear();
                RadioButtonAssistantMale_AdminAssistantTab.Checked = false;
                RadioButtonAssistantFemale_AdminAssistantTab.Checked = false;
                dateTimePickerAssistantDOB_AdminAssistantTab.Value = DateTime.Now;
                txtAssistantAdress_AdminAssistantTab.Clear();
                txtAssistantPNumber_AdminAssistantTab.Clear();
                txtAssistantEmail_AdminAssistantTab.Clear();
                lblAssistantNICNumber_AdminAssistantTab.Text = string.Empty;
                dateTimePickerAssistantRegisteringDate_AdminAssistantTab.Value = DateTime.Now;
                pictureBoxAssistantProfilePhoto_AdminAssistantTab.Image = null;
                btnAssistantDetailsRegister_AdminAssistantTab.Enabled = true;
                PopulateAssistantIDComboBox();
                if (ComboBoxAssistantID_AdminAssistantTab.Items.Count > 0)
                    ComboBoxAssistantID_AdminAssistantTab.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error deleting assistant: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadAssistantsGrid()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string query = "SELECT AssistantID, Title, FirstName, LastName, Gender, DateOfBirth, Address, PhoneNumber, Email, NIC, RegisterDate, IsActive FROM Assistants ORDER BY AssistantID";
                    using (SqlDataAdapter da = new SqlDataAdapter(query, conn))
                    {
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dataGridViewAssistantListReport_AdminAssistantTab.DataSource = dt;
                        dataGridViewAssistantListReport_AdminAssistantTab.ClearSelection();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load assistants: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAssistantListReportPrint_AdminAssistantTab_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "PDF files (*.pdf)|*.pdf";
                sfd.Title = "Save Drivers Report as PDF";
                sfd.FileName = "AssistantReport.pdf";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        // Create PDF document in landscape orientation with margins
                        iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 40f, 40f, 60f, 50f);
                        iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, new System.IO.FileStream(sfd.FileName, System.IO.FileMode.Create));
                        doc.Open();

                        // Fonts
                        var titleFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 18);
                        var subTitleFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 12);
                        var bodyFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA, 9);
                        var boldTableFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 8);
                        var endingFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_OBLIQUE, 8);
                        var noteFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.RED);
                        var blueTitleFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 16, iTextSharp.text.Font.BOLD, new iTextSharp.text.BaseColor(0, 102, 204));

                        // === LOGO and HEADER ===
                        iTextSharp.text.pdf.PdfPTable headerTable = new iTextSharp.text.pdf.PdfPTable(2);
                        headerTable.WidthPercentage = 100;
                        headerTable.SetWidths(new float[] { 20, 80 });

                        // Use logo from pictureBoxESHIFTLogo_AdminMyDashboardTab if available
                        if (pictureBoxESHIFTLogo_AdminMyDashboardTab != null && pictureBoxESHIFTLogo_AdminMyDashboardTab.Image != null)
                        {
                            using (var ms = new System.IO.MemoryStream())
                            {
                                pictureBoxESHIFTLogo_AdminMyDashboardTab.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                ms.Position = 0;
                                var logo = iTextSharp.text.Image.GetInstance(ms);
                                logo.ScaleAbsolute(100f, 50f);
                                var logoCell = new iTextSharp.text.pdf.PdfPCell(logo);
                                logoCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER;
                                logoCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
                                logoCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_TOP;
                                headerTable.AddCell(logoCell);
                            }
                        }
                        else
                        {
                            var emptyLogoCell = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(" "));
                            emptyLogoCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER;
                            headerTable.AddCell(emptyLogoCell);
                        }

                        // Company Info
                        var companyInfoCell = new iTextSharp.text.pdf.PdfPCell();
                        companyInfoCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER;
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("E-Shift Company", subTitleFont));
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("123 Business St,\nColombo, Sri Lanka\nPhone: +94 712345678", bodyFont));
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("Report Date: " + DateTime.Now.ToString("yyyy-MM-dd"), bodyFont));
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("Report #: DRIVREP-" + DateTime.Now.ToString("yyyyMMddHHmmss"), bodyFont));
                        headerTable.AddCell(companyInfoCell);

                        doc.Add(headerTable);
                        doc.Add(new iTextSharp.text.Paragraph("\n", bodyFont)); // Spacing

                        // === Blue Drivers Report Title ===
                        var AssistantsReportTitle = new iTextSharp.text.Paragraph("Assistants Report", blueTitleFont);
                        AssistantsReportTitle.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                        doc.Add(AssistantsReportTitle);
                        doc.Add(new iTextSharp.text.Paragraph("\n", bodyFont)); // Spacing

                        // === Table Section ===
                        var pdfTable = new iTextSharp.text.pdf.PdfPTable(dataGridViewAssistantListReport_AdminAssistantTab.Columns.Count);
                        pdfTable.WidthPercentage = 100;
                        pdfTable.SpacingBefore = 10f;
                        pdfTable.SpacingAfter = 10f;

                        float[] columnWidths = new float[dataGridViewAssistantListReport_AdminAssistantTab.Columns.Count];
                        for (int i = 0; i < dataGridViewAssistantListReport_AdminAssistantTab.Columns.Count; i++)
                            columnWidths[i] = 4f;
                        pdfTable.SetWidths(columnWidths);

                        // Table Headers
                        foreach (DataGridViewColumn column in dataGridViewAssistantListReport_AdminAssistantTab.Columns)
                        {
                            var headerCell = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(column.HeaderText, boldTableFont));
                            headerCell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                            headerCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                            headerCell.Padding = 5;
                            pdfTable.AddCell(headerCell);
                        }

                        // Table Rows
                        foreach (DataGridViewRow row in dataGridViewAssistantListReport_AdminAssistantTab.Rows)
                        {
                            if (!row.IsNewRow)
                            {
                                foreach (DataGridViewCell cell in row.Cells)
                                {
                                    var dataCell = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(cell.Value?.ToString() ?? "", bodyFont));
                                    dataCell.Padding = 4;
                                    dataCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                                    pdfTable.AddCell(dataCell);
                                }
                            }
                        }

                        doc.Add(pdfTable);

                        // === Note for Admin/Company Use Only ===
                        var notePara = new iTextSharp.text.Paragraph("Note: For Admin and Company use only", noteFont);
                        notePara.Alignment = iTextSharp.text.Element.ALIGN_LEFT;
                        doc.Add(notePara);

                        doc.Close();
                        writer.Close();

                        MessageBox.Show("Assistants Report PDF exported successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + sfd.FileName + "\"");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failed to export PDF: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void pictureBoxLorryImageAdd_AdminLorryTab_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Select Driver Profile Photo";
                ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var img = System.Drawing.Image.FromFile(ofd.FileName);
                        pictureBoxLorryImageAdd_AdminLorryTab.Image = img;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failed to load image: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void PopulateLorryIDComboBox()
        {
            try
            {
                ComboBoxLorryID_AdminLorryTab.Items.Clear();
                ComboBoxLorryID_AdminLorryTab.Items.Add("New register");
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string query = "SELECT LorryID FROM Lorries ORDER BY LorryID";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            ComboBoxLorryID_AdminLorryTab.Items.Add(reader["LorryID"].ToString());
                        }
                    }
                }
                ComboBoxLorryID_AdminLorryTab.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load Lorry IDs: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ComboBoxLorryID_AdminLorryTab_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ComboBoxLorryID_AdminLorryTab.SelectedItem != null && ComboBoxLorryID_AdminLorryTab.SelectedItem.ToString() != "New register")
            {
                // Fetch lorry data from DB and populate fields
                int lorryId;
                if (int.TryParse(ComboBoxLorryID_AdminLorryTab.SelectedItem.ToString(), out lorryId))
                {
                    using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                    {
                        conn.Open();
                        string query = "SELECT * FROM Lorries WHERE LorryID = @LorryID";
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@LorryID", lorryId);
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    ComboBoxLorryManufacturer_AdminLorryTab.Text = reader["Manufacturer"].ToString();
                                    txtLorryPlateNumber_AdminLorryTab.Text = reader["PlateNumber"].ToString();
                                    ComboBoxLorryFuelType_AdminLorryTab.Text = reader["FuelType"].ToString();
                                    ComboBoxLorryColour_AdminLorryTab.Text = reader["Colour"].ToString();
                                    if (reader["RegisterDate"] != DBNull.Value)
                                        dateTimePickerLorryRegisteringDate_AdminLorryTab.Value = Convert.ToDateTime(reader["RegisterDate"]);
                                    // Vehicle photo
                                    if (reader["VehiclePhoto"] != DBNull.Value)
                                    {
                                        byte[] imgBytes = (byte[])reader["VehiclePhoto"];
                                        using (var ms = new System.IO.MemoryStream(imgBytes))
                                        {
                                            pictureBoxLorryImageAdd_AdminLorryTab.Image = System.Drawing.Image.FromStream(ms);
                                        }
                                    }
                                    else
                                    {
                                        pictureBoxLorryImageAdd_AdminLorryTab.Image = null;
                                    }
                                }
                            }
                        }
                    }
                }
                btnLorryDetailsRegister_AdminLorryTab.Enabled = false;
            }
            else
            {
                ComboBoxLorryManufacturer_AdminLorryTab.SelectedIndex = -1;
                txtLorryPlateNumber_AdminLorryTab.Clear();
                ComboBoxLorryFuelType_AdminLorryTab.SelectedIndex = -1;
                ComboBoxLorryColour_AdminLorryTab.SelectedIndex = -1;
                dateTimePickerLorryRegisteringDate_AdminLorryTab.Value = DateTime.Now;
                pictureBoxLorryImageAdd_AdminLorryTab.Image = null;
                btnLorryDetailsRegister_AdminLorryTab.Enabled = true;
            }
        }

        private void btnLorryDetailsRegister_AdminLorryTab_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(ComboBoxLorryManufacturer_AdminLorryTab.Text) ||
                    string.IsNullOrWhiteSpace(txtLorryPlateNumber_AdminLorryTab.Text) ||
                    string.IsNullOrWhiteSpace(ComboBoxLorryFuelType_AdminLorryTab.Text) ||
                    string.IsNullOrWhiteSpace(ComboBoxLorryColour_AdminLorryTab.Text))
                {
                    MessageBox.Show("Please fill out all required fields.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                byte[] vehiclePhotoBytes = null;
                if (pictureBoxLorryImageAdd_AdminLorryTab.Image == null)
                {
                    // Prompt user to select an image
                    using (OpenFileDialog ofd = new OpenFileDialog())
                    {
                        ofd.Title = "Select Lorry Photo";
                        ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp";
                        if (ofd.ShowDialog() == DialogResult.OK)
                        {
                            var img = System.Drawing.Image.FromFile(ofd.FileName);
                            pictureBoxLorryImageAdd_AdminLorryTab.Image = img;
                        }
                        else
                        {
                            MessageBox.Show("Please select a lorry photo.", "Missing Photo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                }
                if (pictureBoxLorryImageAdd_AdminLorryTab.Image != null)
                {
                    using (var ms = new System.IO.MemoryStream())
                    {
                        pictureBoxLorryImageAdd_AdminLorryTab.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                        vehiclePhotoBytes = ms.ToArray();
                    }
                }

                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string insertQuery = @"INSERT INTO Lorries (PlateNumber, Manufacturer, Colour, FuelType, RegisterDate, VehiclePhoto, IsActive)
                        VALUES (@PlateNumber, @Manufacturer, @Colour, @FuelType, @RegisterDate, @VehiclePhoto, @IsActive)";
                    using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@PlateNumber", txtLorryPlateNumber_AdminLorryTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@Manufacturer", ComboBoxLorryManufacturer_AdminLorryTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@Colour", ComboBoxLorryColour_AdminLorryTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@FuelType", ComboBoxLorryFuelType_AdminLorryTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@RegisterDate", dateTimePickerLorryRegisteringDate_AdminLorryTab.Value);
                        cmd.Parameters.AddWithValue("@VehiclePhoto", (object)vehiclePhotoBytes ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@IsActive", 1);
                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                        {
                            MessageBox.Show("Lorry registered successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            PopulateLorryIDComboBox();
                            LoadLorriesGrid();
                            PopulateLorryIDComboBox_AdminJobsTab();
                            ComboBoxLorryID_AdminJobsTab.SelectedIndexChanged += ComboBoxLorryID_AdminJobsTab_SelectedIndexChanged;
                        }
                        else
                        {
                            MessageBox.Show("Failed to register lorry.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (SqlException ex) when (ex.Message.Contains("UNIQUE"))
            {
                MessageBox.Show("Plate Number already exists in the system.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error registering lorry: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnLorryDetailsClear_AdminLorryTab_Click(object sender, EventArgs e)
        {
            ComboBoxLorryManufacturer_AdminLorryTab.SelectedIndex = -1;
            txtLorryPlateNumber_AdminLorryTab.Clear();
            ComboBoxLorryFuelType_AdminLorryTab.SelectedIndex = -1;
            ComboBoxLorryColour_AdminLorryTab.SelectedIndex = -1;
            dateTimePickerLorryRegisteringDate_AdminLorryTab.Value = DateTime.Now;
            pictureBoxLorryImageAdd_AdminLorryTab.Image = null;
            btnLorryDetailsRegister_AdminLorryTab.Enabled = true;
            PopulateLorryIDComboBox();
            LoadLorriesGrid();
            ComboBoxLorryID_AdminLorryTab.SelectedIndexChanged += ComboBoxLorryID_AdminLorryTab_SelectedIndexChanged;
        }

        private void btnLorryDetailsUpdate_AdminLorryTab_Click(object sender, EventArgs e)
        {
            if (ComboBoxLorryID_AdminLorryTab.SelectedItem == null || ComboBoxLorryID_AdminLorryTab.SelectedItem.ToString() == "New register")
            {
                MessageBox.Show("Please select a Lorry ID to update.", "No Lorry Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            int lorryId;
            if (!int.TryParse(ComboBoxLorryID_AdminLorryTab.SelectedItem.ToString(), out lorryId))
            {
                MessageBox.Show("Invalid Lorry ID selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                Dictionary<string, object> dbValues = new Dictionary<string, object>();
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string selectQuery = "SELECT * FROM Lorries WHERE LorryID = @LorryID";
                    using (SqlCommand selectCmd = new SqlCommand(selectQuery, conn))
                    {
                        selectCmd.Parameters.AddWithValue("@LorryID", lorryId);
                        using (SqlDataReader reader = selectCmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                dbValues["PlateNumber"] = reader["PlateNumber"].ToString();
                                dbValues["Manufacturer"] = reader["Manufacturer"].ToString();
                                dbValues["Colour"] = reader["Colour"].ToString();
                                dbValues["FuelType"] = reader["FuelType"].ToString();
                                dbValues["RegisterDate"] = reader["RegisterDate"] == DBNull.Value ? null : (object)Convert.ToDateTime(reader["RegisterDate"]);
                                dbValues["VehiclePhoto"] = reader["VehiclePhoto"] == DBNull.Value ? null : (object)reader["VehiclePhoto"];
                            }
                        }
                    }
                }

                byte[] vehiclePhotoBytes = null;
                if (pictureBoxLorryImageAdd_AdminLorryTab.Image != null)
                {
                    using (var clonedImg = new Bitmap(pictureBoxLorryImageAdd_AdminLorryTab.Image))
                    using (var ms = new System.IO.MemoryStream())
                    {
                        clonedImg.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                        vehiclePhotoBytes = ms.ToArray();
                    }
                }

                var updates = new List<string>();
                var parameters = new List<SqlParameter>();
                if (txtLorryPlateNumber_AdminLorryTab.Text.Trim() != (dbValues["PlateNumber"]?.ToString() ?? ""))
                {
                    updates.Add("PlateNumber = @PlateNumber");
                    parameters.Add(new SqlParameter("@PlateNumber", txtLorryPlateNumber_AdminLorryTab.Text.Trim()));
                }
                if (ComboBoxLorryManufacturer_AdminLorryTab.Text.Trim() != (dbValues["Manufacturer"]?.ToString() ?? ""))
                {
                    updates.Add("Manufacturer = @Manufacturer");
                    parameters.Add(new SqlParameter("@Manufacturer", ComboBoxLorryManufacturer_AdminLorryTab.Text.Trim()));
                }
                if (ComboBoxLorryColour_AdminLorryTab.Text.Trim() != (dbValues["Colour"]?.ToString() ?? ""))
                {
                    updates.Add("Colour = @Colour");
                    parameters.Add(new SqlParameter("@Colour", ComboBoxLorryColour_AdminLorryTab.Text.Trim()));
                }
                if (ComboBoxLorryFuelType_AdminLorryTab.Text.Trim() != (dbValues["FuelType"]?.ToString() ?? ""))
                {
                    updates.Add("FuelType = @FuelType");
                    parameters.Add(new SqlParameter("@FuelType", ComboBoxLorryFuelType_AdminLorryTab.Text.Trim()));
                }
                if (dateTimePickerLorryRegisteringDate_AdminLorryTab.Value != (dbValues["RegisterDate"] is DateTime dt ? dt : DateTime.MinValue))
                {
                    updates.Add("RegisterDate = @RegisterDate");
                    parameters.Add(new SqlParameter("@RegisterDate", dateTimePickerLorryRegisteringDate_AdminLorryTab.Value));
                }
                bool photoChanged = false;
                if (vehiclePhotoBytes != null)
                {
                    var dbPhoto = dbValues["VehiclePhoto"] as byte[];
                    if (dbPhoto == null || !vehiclePhotoBytes.SequenceEqual(dbPhoto))
                    {
                        photoChanged = true;
                    }
                }
                else if (dbValues["VehiclePhoto"] != null)
                {
                    photoChanged = true;
                }
                if (photoChanged)
                {
                    updates.Add("VehiclePhoto = @VehiclePhoto");
                    parameters.Add(new SqlParameter("@VehiclePhoto", (object)vehiclePhotoBytes ?? DBNull.Value));
                }

                if (updates.Count == 0)
                {
                    MessageBox.Show("No changes detected. Please modify at least one field to update.", "No Update", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string updateQuery = $"UPDATE Lorries SET {string.Join(", ", updates)} WHERE LorryID = @LorryID";
                    using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                    {
                        foreach (var param in parameters)
                            cmd.Parameters.Add(param);
                        cmd.Parameters.AddWithValue("@LorryID", lorryId);
                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                        {
                            MessageBox.Show("Lorry details updated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            PopulateLorryIDComboBox();
                            LoadLorriesGrid();
                        }
                        else
                        {
                            MessageBox.Show("No changes were made. Please check the details and try again.", "No Update", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
            }
            catch (SqlException ex) when (ex.Message.Contains("UNIQUE"))
            {
                MessageBox.Show("Plate Number already exists in the system.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error updating lorry: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnLorryDetailsDelete_AdminLorryTab_Click(object sender, EventArgs e)
        {
            {
                if (ComboBoxLorryID_AdminLorryTab.SelectedItem == null || ComboBoxLorryID_AdminLorryTab.SelectedItem.ToString() == "New register")
                {
                    MessageBox.Show("Please select a Lorry ID to delete.", "No Lorry Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                int lorryId;
                if (!int.TryParse(ComboBoxLorryID_AdminLorryTab.SelectedItem.ToString(), out lorryId))
                {
                    MessageBox.Show("Invalid Lorry ID selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                var confirm = MessageBox.Show($"Are you sure you want to delete Lorry ID {lorryId}? This action cannot be undone.", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (confirm != DialogResult.Yes)
                    return;
                try
                {
                    using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                    {
                        conn.Open();
                        string deleteQuery = "DELETE FROM Lorries WHERE LorryID = @LorryID";
                        using (SqlCommand cmd = new SqlCommand(deleteQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@LorryID", lorryId);
                            int rows = cmd.ExecuteNonQuery();
                            if (rows > 0)
                            {
                                MessageBox.Show("Lorry deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                PopulateLorryIDComboBox();
                                LoadLorriesGrid();
                            }
                            else
                            {
                                MessageBox.Show("No lorry found to delete. Please check the Lorry ID.", "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }
                    }
                    ComboBoxLorryID_AdminLorryTab.SelectedIndex = -1;
                    ComboBoxLorryManufacturer_AdminLorryTab.SelectedIndex = -1;
                    txtLorryPlateNumber_AdminLorryTab.Clear();
                    ComboBoxLorryFuelType_AdminLorryTab.SelectedIndex = -1;
                    ComboBoxLorryColour_AdminLorryTab.SelectedIndex = -1;
                    dateTimePickerLorryRegisteringDate_AdminLorryTab.Value = DateTime.Now;
                    pictureBoxLorryImageAdd_AdminLorryTab.Image = null;
                    btnLorryDetailsRegister_AdminLorryTab.Enabled = true;
                    PopulateLorryIDComboBox();
                    if (ComboBoxLorryID_AdminLorryTab.Items.Count > 0)
                        ComboBoxLorryID_AdminLorryTab.SelectedIndex = 0;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error deleting lorry: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void LoadLorriesGrid()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string query = "SELECT LorryID, PlateNumber, Manufacturer, Colour, FuelType, RegisterDate, IsActive FROM Lorries ORDER BY LorryID";
                    using (SqlDataAdapter da = new SqlDataAdapter(query, conn))
                    {
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dataGridViewLorryListReport_AdminLorryTab.DataSource = dt;
                        dataGridViewLorryListReport_AdminLorryTab.ClearSelection();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load lorries: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnLorryListReportPrint_AdminLorryTab_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "PDF files (*.pdf)|*.pdf";
                sfd.Title = "Save Lorries Report as PDF";
                sfd.FileName = "LorriesReport.pdf";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        // Create PDF document in landscape orientation with margins
                        iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 40f, 40f, 60f, 50f);
                        iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, new System.IO.FileStream(sfd.FileName, System.IO.FileMode.Create));
                        doc.Open();

                        // Fonts
                        var titleFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 18);
                        var subTitleFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 12);
                        var bodyFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA, 9);
                        var boldTableFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 8);
                        var endingFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_OBLIQUE, 8);
                        var noteFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.RED);
                        var blueTitleFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 16, iTextSharp.text.Font.BOLD, new iTextSharp.text.BaseColor(0, 102, 204));

                        // === LOGO and HEADER ===
                        iTextSharp.text.pdf.PdfPTable headerTable = new iTextSharp.text.pdf.PdfPTable(2);
                        headerTable.WidthPercentage = 100;
                        headerTable.SetWidths(new float[] { 20, 80 });

                        // Use logo from pictureBoxESHIFTLogo_AdminMyDashboardTab if available
                        if (pictureBoxESHIFTLogo_AdminMyDashboardTab != null && pictureBoxESHIFTLogo_AdminMyDashboardTab.Image != null)
                        {
                            using (var ms = new System.IO.MemoryStream())
                            {
                                pictureBoxESHIFTLogo_AdminMyDashboardTab.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                ms.Position = 0;
                                var logo = iTextSharp.text.Image.GetInstance(ms);
                                logo.ScaleAbsolute(100f, 50f);
                                var logoCell = new iTextSharp.text.pdf.PdfPCell(logo);
                                logoCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER;
                                logoCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
                                logoCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_TOP;
                                headerTable.AddCell(logoCell);
                            }
                        }
                        else
                        {
                            var emptyLogoCell = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(" "));
                            emptyLogoCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER;
                            headerTable.AddCell(emptyLogoCell);
                        }

                        // Company Info
                        var companyInfoCell = new iTextSharp.text.pdf.PdfPCell();
                        companyInfoCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER;
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("E-Shift Company", subTitleFont));
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("123 Business St,\nColombo, Sri Lanka\nPhone: +94 712345678", bodyFont));
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("Report Date: " + DateTime.Now.ToString("yyyy-MM-dd"), bodyFont));
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("Report #: DRIVREP-" + DateTime.Now.ToString("yyyyMMddHHmmss"), bodyFont));
                        headerTable.AddCell(companyInfoCell);

                        doc.Add(headerTable);
                        doc.Add(new iTextSharp.text.Paragraph("\n", bodyFont)); // Spacing

                        // === Blue Drivers Report Title ===
                        var driversReportTitle = new iTextSharp.text.Paragraph("Lorries Report", blueTitleFont);
                        driversReportTitle.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                        doc.Add(driversReportTitle);
                        doc.Add(new iTextSharp.text.Paragraph("\n", bodyFont)); // Spacing

                        // === Table Section ===
                        var pdfTable = new iTextSharp.text.pdf.PdfPTable(dataGridViewLorryListReport_AdminLorryTab.Columns.Count);
                        pdfTable.WidthPercentage = 100;
                        pdfTable.SpacingBefore = 10f;
                        pdfTable.SpacingAfter = 10f;

                        float[] columnWidths = new float[dataGridViewLorryListReport_AdminLorryTab.Columns.Count];
                        for (int i = 0; i < dataGridViewLorryListReport_AdminLorryTab.Columns.Count; i++)
                            columnWidths[i] = 4f;
                        pdfTable.SetWidths(columnWidths);

                        // Table Headers
                        foreach (DataGridViewColumn column in dataGridViewLorryListReport_AdminLorryTab.Columns)
                        {
                            var headerCell = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(column.HeaderText, boldTableFont));
                            headerCell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                            headerCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                            headerCell.Padding = 5;
                            pdfTable.AddCell(headerCell);
                        }

                        // Table Rows
                        foreach (DataGridViewRow row in dataGridViewLorryListReport_AdminLorryTab.Rows)
                        {
                            if (!row.IsNewRow)
                            {
                                foreach (DataGridViewCell cell in row.Cells)
                                {
                                    var dataCell = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(cell.Value?.ToString() ?? "", bodyFont));
                                    dataCell.Padding = 4;
                                    dataCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                                    pdfTable.AddCell(dataCell);
                                }
                            }
                        }

                        doc.Add(pdfTable);

                        // === Note for Admin/Company Use Only ===
                        var notePara = new iTextSharp.text.Paragraph("Note: For Admin and Company use only", noteFont);
                        notePara.Alignment = iTextSharp.text.Element.ALIGN_LEFT;
                        doc.Add(notePara);

                        doc.Close();
                        writer.Close();

                        MessageBox.Show("Lorries Report PDF exported successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + sfd.FileName + "\"");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failed to export PDF: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void pictureBoxCOntainerPhotoAdd_AdminContainerTab_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Select Driver Profile Photo";
                ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var img = System.Drawing.Image.FromFile(ofd.FileName);
                        pictureBoxCOntainerPhotoAdd_AdminContainerTab.Image = img;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failed to load image: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void PopulateContainerIDComboBox()
        {
            try
            {
                ComboBoxContainerID_AdminContainerTab.Items.Clear();
                ComboBoxContainerID_AdminContainerTab.Items.Add("New register");
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string query = "SELECT ContainerID FROM Containers ORDER BY ContainerID";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            ComboBoxContainerID_AdminContainerTab.Items.Add(reader["ContainerID"].ToString());
                        }
                    }
                }
                ComboBoxContainerID_AdminContainerTab.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load Container IDs: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ComboBoxContainerID_AdminContainerTab_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ComboBoxContainerID_AdminContainerTab.SelectedItem != null && ComboBoxContainerID_AdminContainerTab.SelectedItem.ToString() != "New register")
            {
                int containerId;
                if (int.TryParse(ComboBoxContainerID_AdminContainerTab.SelectedItem.ToString(), out containerId))
                {
                    using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                    {
                        conn.Open();
                        string query = "SELECT * FROM Containers WHERE ContainerID = @ContainerID";
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@ContainerID", containerId);
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    ComboBoxManufactuer_AdminContainerTab.Text = reader["Manufacturer"].ToString();
                                    txtContainerPLateNumber_AdminContainerTab.Text = reader["PlateNumber"].ToString();
                                    ComboBoxCOntainerColour_AdminContainerTab.Text = reader["Colour"].ToString();
                                    txtContainerTarWeight.Text = reader["TarWeight"].ToString();
                                    txtContainerMaxWeight_AdminContainerTab.Text = reader["MaxWeight"].ToString();
                                    txtContainerLength_AdminContainerTab.Text = reader["Length"].ToString();
                                    txtContainerWidth_AdminContainerTab.Text = reader["Width"].ToString();
                                    txtContainerHight_AdminContainerTab.Text = reader["Height"].ToString();
                                    if (reader["RegisterDate"] != DBNull.Value)
                                        dateTimePickerContainerRegisteringDate_AdminContainerTab.Value = Convert.ToDateTime(reader["RegisterDate"]);
                                    if (reader["ContainerPhoto"] != DBNull.Value)
                                    {
                                        byte[] imgBytes = (byte[])reader["ContainerPhoto"];
                                        using (var ms = new System.IO.MemoryStream(imgBytes))
                                        {
                                            pictureBoxCOntainerPhotoAdd_AdminContainerTab.Image = System.Drawing.Image.FromStream(ms);
                                        }
                                    }
                                    else
                                    {
                                        pictureBoxCOntainerPhotoAdd_AdminContainerTab.Image = null;
                                    }
                                }
                            }
                        }
                    }
                }
                btnContainerDetailsRegister_AdminContainerTab.Enabled = false;
            }
            else
            {
                ComboBoxManufactuer_AdminContainerTab.SelectedIndex = -1;
                txtContainerPLateNumber_AdminContainerTab.Clear();
                ComboBoxCOntainerColour_AdminContainerTab.SelectedIndex = -1;
                txtContainerTarWeight.Clear();
                txtContainerMaxWeight_AdminContainerTab.Clear();
                txtContainerLength_AdminContainerTab.Clear();
                txtContainerWidth_AdminContainerTab.Clear();
                txtContainerHight_AdminContainerTab.Clear();
                dateTimePickerContainerRegisteringDate_AdminContainerTab.Value = DateTime.Now;
                pictureBoxCOntainerPhotoAdd_AdminContainerTab.Image = null;
                btnContainerDetailsRegister_AdminContainerTab.Enabled = true;
            }
        }

        private void btnContainerDetailsRegister_AdminContainerTab_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(ComboBoxManufactuer_AdminContainerTab.Text) ||
                    string.IsNullOrWhiteSpace(txtContainerPLateNumber_AdminContainerTab.Text) ||
                    string.IsNullOrWhiteSpace(ComboBoxCOntainerColour_AdminContainerTab.Text) ||
                    string.IsNullOrWhiteSpace(txtContainerLength_AdminContainerTab.Text) ||
                    string.IsNullOrWhiteSpace(txtContainerHight_AdminContainerTab.Text) ||
                    string.IsNullOrWhiteSpace(txtContainerWidth_AdminContainerTab.Text) ||
                    string.IsNullOrWhiteSpace(txtContainerTarWeight.Text) ||
                    string.IsNullOrWhiteSpace(txtContainerMaxWeight_AdminContainerTab.Text))
                {
                    MessageBox.Show("Please fill out all required fields.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                byte[] containerPhotoBytes = null;
                if (pictureBoxCOntainerPhotoAdd_AdminContainerTab.Image == null)
                {
                    using (OpenFileDialog ofd = new OpenFileDialog())
                    {
                        ofd.Title = "Select Container Photo";
                        ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp";
                        if (ofd.ShowDialog() == DialogResult.OK)
                        {
                            var img = System.Drawing.Image.FromFile(ofd.FileName);
                            pictureBoxCOntainerPhotoAdd_AdminContainerTab.Image = img;
                        }
                        else
                        {
                            MessageBox.Show("Please select a container photo.", "Missing Photo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                }
                if (pictureBoxCOntainerPhotoAdd_AdminContainerTab.Image != null)
                {
                    using (var ms = new System.IO.MemoryStream())
                    {
                        pictureBoxCOntainerPhotoAdd_AdminContainerTab.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                        containerPhotoBytes = ms.ToArray();
                    }
                }

                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string insertQuery = @"INSERT INTO Containers 
                (PlateNumber, Manufacturer, Colour, TarWeight, MaxWeight, Length, Width, Height, RegisterDate, ContainerPhoto, IsActive)
                VALUES (@PlateNumber, @Manufacturer, @Colour, @TarWeight, @MaxWeight, @Length, @Width, @Height, @RegisterDate, @ContainerPhoto, @IsActive)";
                    using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@PlateNumber", txtContainerPLateNumber_AdminContainerTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@Manufacturer", ComboBoxManufactuer_AdminContainerTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@Colour", ComboBoxCOntainerColour_AdminContainerTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@TarWeight", decimal.Parse(txtContainerTarWeight.Text.Trim()));
                        cmd.Parameters.AddWithValue("@MaxWeight", decimal.Parse(txtContainerMaxWeight_AdminContainerTab.Text.Trim()));
                        cmd.Parameters.AddWithValue("@Length", decimal.Parse(txtContainerLength_AdminContainerTab.Text.Trim()));
                        cmd.Parameters.AddWithValue("@Width", decimal.Parse(txtContainerWidth_AdminContainerTab.Text.Trim()));
                        cmd.Parameters.AddWithValue("@Height", decimal.Parse(txtContainerHight_AdminContainerTab.Text.Trim()));
                        cmd.Parameters.AddWithValue("@RegisterDate", dateTimePickerContainerRegisteringDate_AdminContainerTab.Value);
                        cmd.Parameters.AddWithValue("@ContainerPhoto", (object)containerPhotoBytes ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@IsActive", 1);
                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                        {
                            MessageBox.Show("Container registered successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            PopulateContainerIDComboBox();
                            LoadContainersGrid();
                            PopulateContainerIDComboBox_AdminJobsTab();
                            ComboBoxContainerID_AdminJobsTab.SelectedIndexChanged += ComboBoxContainerID_AdminJobsTab_SelectedIndexChanged;
                        }
                        else
                        {
                            MessageBox.Show("Failed to register container.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (SqlException ex) when (ex.Message.Contains("UNIQUE"))
            {
                MessageBox.Show("Plate Number already exists in the system.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error registering container: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnContainerDetailsUpdate_AdminContainerTab_Click(object sender, EventArgs e)
        {
            if (ComboBoxContainerID_AdminContainerTab.SelectedItem == null || ComboBoxContainerID_AdminContainerTab.SelectedItem.ToString() == "New register")
            {
                MessageBox.Show("Please select a Container ID to update.", "No Container Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            int containerId;
            if (!int.TryParse(ComboBoxContainerID_AdminContainerTab.SelectedItem.ToString(), out containerId))
            {
                MessageBox.Show("Invalid Container ID selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                // Fetch current DB values for the selected container
                Dictionary<string, object> dbValues = new Dictionary<string, object>();
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string selectQuery = "SELECT * FROM Containers WHERE ContainerID = @ContainerID";
                    using (SqlCommand selectCmd = new SqlCommand(selectQuery, conn))
                    {
                        selectCmd.Parameters.AddWithValue("@ContainerID", containerId);
                        using (SqlDataReader reader = selectCmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                dbValues["PlateNumber"] = reader["PlateNumber"].ToString();
                                dbValues["Manufacturer"] = reader["Manufacturer"].ToString();
                                dbValues["Colour"] = reader["Colour"].ToString();
                                dbValues["TarWeight"] = reader["TarWeight"];
                                dbValues["MaxWeight"] = reader["MaxWeight"];
                                dbValues["Length"] = reader["Length"];
                                dbValues["Width"] = reader["Width"];
                                dbValues["Height"] = reader["Height"];
                                dbValues["RegisterDate"] = reader["RegisterDate"] == DBNull.Value ? null : (object)Convert.ToDateTime(reader["RegisterDate"]);
                                dbValues["ContainerPhoto"] = reader["ContainerPhoto"] == DBNull.Value ? null : (object)reader["ContainerPhoto"];
                            }
                        }
                    }
                }

                // Prepare new values from form
                byte[] containerPhotoBytes = null;
                if (pictureBoxCOntainerPhotoAdd_AdminContainerTab.Image != null)
                {
                    using (var clonedImg = new Bitmap(pictureBoxCOntainerPhotoAdd_AdminContainerTab.Image))
                    using (var ms = new System.IO.MemoryStream())
                    {
                        clonedImg.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                        containerPhotoBytes = ms.ToArray();
                    }
                }

                // Compare and build update query only for changed fields
                var updates = new List<string>();
                var parameters = new List<SqlParameter>();
                if (txtContainerPLateNumber_AdminContainerTab.Text.Trim() != (dbValues["PlateNumber"]?.ToString() ?? ""))
                {
                    updates.Add("PlateNumber = @PlateNumber");
                    parameters.Add(new SqlParameter("@PlateNumber", txtContainerPLateNumber_AdminContainerTab.Text.Trim()));
                }
                if (ComboBoxManufactuer_AdminContainerTab.Text.Trim() != (dbValues["Manufacturer"]?.ToString() ?? ""))
                {
                    updates.Add("Manufacturer = @Manufacturer");
                    parameters.Add(new SqlParameter("@Manufacturer", ComboBoxManufactuer_AdminContainerTab.Text.Trim()));
                }
                if (ComboBoxCOntainerColour_AdminContainerTab.Text.Trim() != (dbValues["Colour"]?.ToString() ?? ""))
                {
                    updates.Add("Colour = @Colour");
                    parameters.Add(new SqlParameter("@Colour", ComboBoxCOntainerColour_AdminContainerTab.Text.Trim()));
                }
                if (decimal.Parse(txtContainerTarWeight.Text.Trim()) != Convert.ToDecimal(dbValues["TarWeight"]))
                {
                    updates.Add("TarWeight = @TarWeight");
                    parameters.Add(new SqlParameter("@TarWeight", decimal.Parse(txtContainerTarWeight.Text.Trim())));
                }
                if (decimal.Parse(txtContainerMaxWeight_AdminContainerTab.Text.Trim()) != Convert.ToDecimal(dbValues["MaxWeight"]))
                {
                    updates.Add("MaxWeight = @MaxWeight");
                    parameters.Add(new SqlParameter("@MaxWeight", decimal.Parse(txtContainerMaxWeight_AdminContainerTab.Text.Trim())));
                }
                if (decimal.Parse(txtContainerLength_AdminContainerTab.Text.Trim()) != Convert.ToDecimal(dbValues["Length"]))
                {
                    updates.Add("Length = @Length");
                    parameters.Add(new SqlParameter("@Length", decimal.Parse(txtContainerLength_AdminContainerTab.Text.Trim())));
                }
                if (decimal.Parse(txtContainerWidth_AdminContainerTab.Text.Trim()) != Convert.ToDecimal(dbValues["Width"]))
                {
                    updates.Add("Width = @Width");
                    parameters.Add(new SqlParameter("@Width", decimal.Parse(txtContainerWidth_AdminContainerTab.Text.Trim())));
                }
                if (decimal.Parse(txtContainerHight_AdminContainerTab.Text.Trim()) != Convert.ToDecimal(dbValues["Height"]))
                {
                    updates.Add("Height = @Height");
                    parameters.Add(new SqlParameter("@Height", decimal.Parse(txtContainerHight_AdminContainerTab.Text.Trim())));
                }
                if (dateTimePickerContainerRegisteringDate_AdminContainerTab.Value != (dbValues["RegisterDate"] is DateTime dt ? dt : DateTime.MinValue))
                {
                    updates.Add("RegisterDate = @RegisterDate");
                    parameters.Add(new SqlParameter("@RegisterDate", dateTimePickerContainerRegisteringDate_AdminContainerTab.Value));
                }
                // Profile photo: compare byte arrays
                bool photoChanged = false;
                if (containerPhotoBytes != null)
                {
                    var dbPhoto = dbValues["ContainerPhoto"] as byte[];
                    if (dbPhoto == null || !containerPhotoBytes.SequenceEqual(dbPhoto))
                    {
                        photoChanged = true;
                    }
                }
                else if (dbValues["ContainerPhoto"] != null)
                {
                    photoChanged = true;
                }
                if (photoChanged)
                {
                    updates.Add("ContainerPhoto = @ContainerPhoto");
                    parameters.Add(new SqlParameter("@ContainerPhoto", (object)containerPhotoBytes ?? DBNull.Value));
                }

                if (updates.Count == 0)
                {
                    MessageBox.Show("No changes detected. Please modify at least one field to update.", "No Update", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string updateQuery = $"UPDATE Containers SET {string.Join(", ", updates)} WHERE ContainerID = @ContainerID";
                    using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                    {
                        foreach (var param in parameters)
                            cmd.Parameters.Add(param);
                        cmd.Parameters.AddWithValue("@ContainerID", containerId);
                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                        {
                            MessageBox.Show("Container details updated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            PopulateContainerIDComboBox();
                            LoadContainersGrid();
                        }
                        else
                        {
                            MessageBox.Show("No changes were made. Please check the details and try again.", "No Update", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
            }
            catch (SqlException ex) when (ex.Message.Contains("UNIQUE"))
            {
                MessageBox.Show("Plate Number already exists in the system.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error updating container: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void guna2ButtonContainerDetailsDelete_AdminContainerTab_Click(object sender, EventArgs e)
        {
            if (ComboBoxContainerID_AdminContainerTab.SelectedItem == null || ComboBoxContainerID_AdminContainerTab.SelectedItem.ToString() == "New register")
            {
                MessageBox.Show("Please select a Container ID to delete.", "No Container Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            int containerId;
            if (!int.TryParse(ComboBoxContainerID_AdminContainerTab.SelectedItem.ToString(), out containerId))
            {
                MessageBox.Show("Invalid Container ID selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var confirm = MessageBox.Show($"Are you sure you want to delete Container ID {containerId}? This action cannot be undone.", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (confirm != DialogResult.Yes)
                return;
            try
            {
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string deleteQuery = "DELETE FROM Containers WHERE ContainerID = @ContainerID";
                    using (SqlCommand cmd = new SqlCommand(deleteQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@ContainerID", containerId);
                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                        {
                            MessageBox.Show("Container deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            PopulateContainerIDComboBox();
                            LoadContainersGrid();
                        }
                        else
                        {
                            MessageBox.Show("No container found to delete. Please check the Container ID.", "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                }
                ComboBoxContainerID_AdminContainerTab.SelectedIndex = -1;
                ComboBoxManufactuer_AdminContainerTab.SelectedIndex = -1;
                txtContainerPLateNumber_AdminContainerTab.Clear();
                ComboBoxCOntainerColour_AdminContainerTab.SelectedIndex = -1;
                txtContainerTarWeight.Clear();
                txtContainerMaxWeight_AdminContainerTab.Clear();
                txtContainerLength_AdminContainerTab.Clear();
                txtContainerWidth_AdminContainerTab.Clear();
                txtContainerHight_AdminContainerTab.Clear();
                dateTimePickerContainerRegisteringDate_AdminContainerTab.Value = DateTime.Now;
                pictureBoxCOntainerPhotoAdd_AdminContainerTab.Image = null;
                btnContainerDetailsRegister_AdminContainerTab.Enabled = true;
                PopulateContainerIDComboBox();
                if (ComboBoxContainerID_AdminContainerTab.Items.Count > 0)
                    ComboBoxContainerID_AdminContainerTab.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error deleting container: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnContainerDetailsClear_AdminContainerTab_Click(object sender, EventArgs e)
        {
            ComboBoxContainerID_AdminContainerTab.SelectedIndex = -1;
            ComboBoxManufactuer_AdminContainerTab.SelectedIndex = -1;
            txtContainerPLateNumber_AdminContainerTab.Clear();
            ComboBoxCOntainerColour_AdminContainerTab.SelectedIndex = -1;
            txtContainerTarWeight.Clear();
            txtContainerMaxWeight_AdminContainerTab.Clear();
            txtContainerLength_AdminContainerTab.Clear();
            txtContainerWidth_AdminContainerTab.Clear();
            txtContainerHight_AdminContainerTab.Clear();
            dateTimePickerContainerRegisteringDate_AdminContainerTab.Value = DateTime.Now;
            pictureBoxCOntainerPhotoAdd_AdminContainerTab.Image = null;
            btnContainerDetailsRegister_AdminContainerTab.Enabled = true;
            PopulateContainerIDComboBox();
            if (ComboBoxContainerID_AdminContainerTab.Items.Count > 0)
                ComboBoxContainerID_AdminContainerTab.SelectedIndex = 0;
        }

        private void LoadContainersGrid()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string query = "SELECT ContainerID, PlateNumber, Manufacturer, Colour, TarWeight, MaxWeight, Length, Width, Height, RegisterDate, IsActive FROM Containers ORDER BY ContainerID";
                    using (SqlDataAdapter da = new SqlDataAdapter(query, conn))
                    {
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dataGridViewContainerListReport_AdminContainerTab.DataSource = dt;
                        dataGridViewContainerListReport_AdminContainerTab.ClearSelection();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load containers: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnContainerListReportPrint_AdminContainerTab_Click(object sender, EventArgs e)
        {
            // Export dataGridViewContainerListReport_AdminContainerTab to PDF using SaveFileDialog and iTextSharp
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "PDF files (*.pdf)|*.pdf";
                sfd.Title = "Save Container Report as PDF";
                sfd.FileName = "ContainerReport.pdf";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        // Create PDF document in landscape orientation with margins
                        iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 40f, 40f, 60f, 50f);
                        iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, new System.IO.FileStream(sfd.FileName, System.IO.FileMode.Create));
                        doc.Open();

                        // Fonts
                        var titleFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 18);
                        var subTitleFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 12);
                        var bodyFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA, 9);
                        var boldTableFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 8);
                        var endingFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_OBLIQUE, 8);
                        var noteFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.RED);
                        var blueTitleFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 16, iTextSharp.text.Font.BOLD, new iTextSharp.text.BaseColor(0, 102, 204));

                        // === LOGO and HEADER ===
                        iTextSharp.text.pdf.PdfPTable headerTable = new iTextSharp.text.pdf.PdfPTable(2);
                        headerTable.WidthPercentage = 100;
                        headerTable.SetWidths(new float[] { 20, 80 });

                        // Use logo from pictureBoxESHIFTLogo_AdminMyDashboardTab if available
                        if (pictureBoxESHIFTLogo_AdminMyDashboardTab != null && pictureBoxESHIFTLogo_AdminMyDashboardTab.Image != null)
                        {
                            using (var ms = new System.IO.MemoryStream())
                            {
                                pictureBoxESHIFTLogo_AdminMyDashboardTab.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                ms.Position = 0;
                                var logo = iTextSharp.text.Image.GetInstance(ms);
                                logo.ScaleAbsolute(100f, 50f);
                                var logoCell = new iTextSharp.text.pdf.PdfPCell(logo);
                                logoCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER;
                                logoCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
                                logoCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_TOP;
                                headerTable.AddCell(logoCell);
                            }
                        }
                        else
                        {
                            var emptyLogoCell = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(" "));
                            emptyLogoCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER;
                            headerTable.AddCell(emptyLogoCell);
                        }

                        // Company Info
                        var companyInfoCell = new iTextSharp.text.pdf.PdfPCell();
                        companyInfoCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER;
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("E-Shift Company", subTitleFont));
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("123 Business St,\nColombo, Sri Lanka\nPhone: +94 712345678", bodyFont));
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("Report Date: " + DateTime.Now.ToString("yyyy-MM-dd"), bodyFont));
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("Report #: DRIVREP-" + DateTime.Now.ToString("yyyyMMddHHmmss"), bodyFont));
                        headerTable.AddCell(companyInfoCell);

                        doc.Add(headerTable);
                        doc.Add(new iTextSharp.text.Paragraph("\n", bodyFont)); // Spacing

                        // === Blue Drivers Report Title ===
                        var driversReportTitle = new iTextSharp.text.Paragraph("Container Report", blueTitleFont);
                        driversReportTitle.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                        doc.Add(driversReportTitle);
                        doc.Add(new iTextSharp.text.Paragraph("\n", bodyFont)); // Spacing

                        // === Table Section ===
                        var pdfTable = new iTextSharp.text.pdf.PdfPTable(dataGridViewContainerListReport_AdminContainerTab.Columns.Count);
                        pdfTable.WidthPercentage = 100;
                        pdfTable.SpacingBefore = 10f;
                        pdfTable.SpacingAfter = 10f;

                        float[] columnWidths = new float[dataGridViewContainerListReport_AdminContainerTab.Columns.Count];
                        for (int i = 0; i < dataGridViewContainerListReport_AdminContainerTab.Columns.Count; i++)
                            columnWidths[i] = 4f;
                        pdfTable.SetWidths(columnWidths);

                        // Table Headers
                        foreach (DataGridViewColumn column in dataGridViewContainerListReport_AdminContainerTab.Columns)
                        {
                            var headerCell = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(column.HeaderText, boldTableFont));
                            headerCell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                            headerCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                            headerCell.Padding = 5;
                            pdfTable.AddCell(headerCell);
                        }

                        // Table Rows
                        foreach (DataGridViewRow row in dataGridViewContainerListReport_AdminContainerTab.Rows)
                        {
                            if (!row.IsNewRow)
                            {
                                foreach (DataGridViewCell cell in row.Cells)
                                {
                                    var dataCell = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(cell.Value?.ToString() ?? "", bodyFont));
                                    dataCell.Padding = 4;
                                    dataCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                                    pdfTable.AddCell(dataCell);
                                }
                            }
                        }

                        doc.Add(pdfTable);

                        // === Note for Admin/Company Use Only ===
                        var notePara = new iTextSharp.text.Paragraph("Note: For Admin and Company use only", noteFont);
                        notePara.Alignment = iTextSharp.text.Element.ALIGN_LEFT;
                        doc.Add(notePara);

                        doc.Close();
                        writer.Close();

                        MessageBox.Show("Container Report PDF exported successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + sfd.FileName + "\"");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failed to export PDF: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnAdd_AdminProductsTab_Click(object sender, EventArgs e)
        {
            // Validate input
            if (string.IsNullOrWhiteSpace(txtProductName_AdminProductsTab.Text))
            {
                MessageBox.Show("Please enter a product name.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string insertQuery = "INSERT INTO ProductTypes (ProductName) VALUES (@ProductName)";
                    using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@ProductName", txtProductName_AdminProductsTab.Text.Trim());
                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                        {
                            MessageBox.Show("Product type added successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            txtProductName_AdminProductsTab.Clear();
                            LoadProductTypesGrid();
                        }
                        else
                        {
                            MessageBox.Show("Failed to add product type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (SqlException ex) when (ex.Message.Contains("UNIQUE"))
            {
                MessageBox.Show("Product name already exists.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error adding product type: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDelete_AdminProductsTab_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtProductName_AdminProductsTab.Text))
            {
                MessageBox.Show("Please enter the product name to delete.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var confirm = MessageBox.Show($"Are you sure you want to delete the product type '{txtProductName_AdminProductsTab.Text.Trim()}'? This action cannot be undone.", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (confirm != DialogResult.Yes)
                return;
            try
            {
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string deleteQuery = "DELETE FROM ProductTypes WHERE ProductName = @ProductName";
                    using (SqlCommand cmd = new SqlCommand(deleteQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@ProductName", txtProductName_AdminProductsTab.Text.Trim());
                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                        {
                            MessageBox.Show("Product type deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            txtProductName_AdminProductsTab.Clear();
                            LoadProductTypesGrid();
                        }
                        else
                        {
                            MessageBox.Show("No product type found with that name.", "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error deleting product type: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnClearlblProducts_AdminProductsTab_Click(object sender, EventArgs e)
        {
            txtProductName_AdminProductsTab.Clear();

        }

        private void LoadProductTypesGrid()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string query = "SELECT ProductTypeID, ProductName FROM ProductTypes ORDER BY ProductTypeID";
                    using (SqlDataAdapter da = new SqlDataAdapter(query, conn))
                    {
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dataGridViewProducts_AdminProductsTab.DataSource = dt;
                        dataGridViewProducts_AdminProductsTab.ClearSelection();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load product types: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnProductsPrintReportlblProducts_AdminProductsTab_Click(object sender, EventArgs e)
        {
            // Export dataGridViewProducts_AdminProductsTab to PDF using SaveFileDialog and iTextSharp
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "PDF files (*.pdf)|*.pdf";
                sfd.Title = "Save Products Report as PDF";
                sfd.FileName = "ProductsReport.pdf";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        // Create PDF document in landscape orientation with margins
                        iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 40f, 40f, 60f, 50f);
                        iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, new System.IO.FileStream(sfd.FileName, System.IO.FileMode.Create));
                        doc.Open();

                        // Fonts
                        var titleFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 18);
                        var subTitleFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 12);
                        var bodyFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA, 9);
                        var boldTableFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_BOLD, 8);
                        var endingFont = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA_OBLIQUE, 8);
                        var noteFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.RED);
                        var blueTitleFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 16, iTextSharp.text.Font.BOLD, new iTextSharp.text.BaseColor(0, 102, 204));

                        // === LOGO and HEADER ===
                        iTextSharp.text.pdf.PdfPTable headerTable = new iTextSharp.text.pdf.PdfPTable(2);
                        headerTable.WidthPercentage = 100;
                        headerTable.SetWidths(new float[] { 20, 80 });

                        // Use logo from pictureBoxESHIFTLogo_AdminMyDashboardTab if available
                        if (pictureBoxESHIFTLogo_AdminMyDashboardTab != null && pictureBoxESHIFTLogo_AdminMyDashboardTab.Image != null)
                        {
                            using (var ms = new System.IO.MemoryStream())
                            {
                                pictureBoxESHIFTLogo_AdminMyDashboardTab.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                ms.Position = 0;
                                var logo = iTextSharp.text.Image.GetInstance(ms);
                                logo.ScaleAbsolute(100f, 50f);
                                var logoCell = new iTextSharp.text.pdf.PdfPCell(logo);
                                logoCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER;
                                logoCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
                                logoCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_TOP;
                                headerTable.AddCell(logoCell);
                            }
                        }
                        else
                        {
                            var emptyLogoCell = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(" "));
                            emptyLogoCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER;
                            headerTable.AddCell(emptyLogoCell);
                        }

                        // Company Info
                        var companyInfoCell = new iTextSharp.text.pdf.PdfPCell();
                        companyInfoCell.Border = iTextSharp.text.pdf.PdfPCell.NO_BORDER;
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("E-Shift Company", subTitleFont));
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("123 Business St,\nColombo, Sri Lanka\nPhone: +94 712345678", bodyFont));
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("Report Date: " + DateTime.Now.ToString("yyyy-MM-dd"), bodyFont));
                        companyInfoCell.AddElement(new iTextSharp.text.Paragraph("Report #: DRIVREP-" + DateTime.Now.ToString("yyyyMMddHHmmss"), bodyFont));
                        headerTable.AddCell(companyInfoCell);

                        doc.Add(headerTable);
                        doc.Add(new iTextSharp.text.Paragraph("\n", bodyFont)); // Spacing

                        // === Blue Drivers Report Title ===
                        var driversReportTitle = new iTextSharp.text.Paragraph("Products Report", blueTitleFont);
                        driversReportTitle.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                        doc.Add(driversReportTitle);
                        doc.Add(new iTextSharp.text.Paragraph("\n", bodyFont)); // Spacing

                        // === Table Section ===
                        var pdfTable = new iTextSharp.text.pdf.PdfPTable(dataGridViewProducts_AdminProductsTab.Columns.Count);
                        pdfTable.WidthPercentage = 100;
                        pdfTable.SpacingBefore = 10f;
                        pdfTable.SpacingAfter = 10f;

                        float[] columnWidths = new float[dataGridViewProducts_AdminProductsTab.Columns.Count];
                        for (int i = 0; i < dataGridViewProducts_AdminProductsTab.Columns.Count; i++)
                            columnWidths[i] = 4f;
                        pdfTable.SetWidths(columnWidths);

                        // Table Headers
                        foreach (DataGridViewColumn column in dataGridViewProducts_AdminProductsTab.Columns)
                        {
                            var headerCell = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(column.HeaderText, boldTableFont));
                            headerCell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                            headerCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                            headerCell.Padding = 5;
                            pdfTable.AddCell(headerCell);
                        }

                        // Table Rows
                        foreach (DataGridViewRow row in dataGridViewProducts_AdminProductsTab.Rows)
                        {
                            if (!row.IsNewRow)
                            {
                                foreach (DataGridViewCell cell in row.Cells)
                                {
                                    var dataCell = new iTextSharp.text.pdf.PdfPCell(new iTextSharp.text.Phrase(cell.Value?.ToString() ?? "", bodyFont));
                                    dataCell.Padding = 4;
                                    dataCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                                    pdfTable.AddCell(dataCell);
                                }
                            }
                        }

                        doc.Add(pdfTable);

                        // === Note for Admin/Company Use Only ===
                        var notePara = new iTextSharp.text.Paragraph("Note: For Admin and Company use only", noteFont);
                        notePara.Alignment = iTextSharp.text.Element.ALIGN_LEFT;
                        doc.Add(notePara);

                        doc.Close();
                        writer.Close();

                        MessageBox.Show("Products Report PDF exported successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + sfd.FileName + "\"");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failed to export PDF: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void PopulateMailingCustomerIDComboBox()
        {
            try
            {
                ComboBoxMailingCustomerID_AdminMailingTab.Items.Clear();
                ComboBoxMailingCustomerID_AdminMailingTab.Items.Add("Select Customer");
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string query = "SELECT CustomerID FROM Customers WHERE IsActive = 1 ORDER BY CustomerID";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            ComboBoxMailingCustomerID_AdminMailingTab.Items.Add(reader["CustomerID"].ToString());
                        }
                    }
                }
                ComboBoxMailingCustomerID_AdminMailingTab.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load Customer IDs: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ComboBoxMailingCustomerID_AdminMailingTab_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ComboBoxMailingCustomerID_AdminMailingTab.SelectedIndex <= 0)
            {
                txtMailingCustomerFName_AdminMailingTab.Clear();
                txtMailingCustomerLName_AdminMailingTab.Clear();
                txtMailingCustomerEmail_AdminMailingTab.Clear();
                return;
            }
            int customerId;
            if (!int.TryParse(ComboBoxMailingCustomerID_AdminMailingTab.SelectedItem.ToString(), out customerId))
            {
                txtMailingCustomerFName_AdminMailingTab.Clear();
                txtMailingCustomerLName_AdminMailingTab.Clear();
                txtMailingCustomerEmail_AdminMailingTab.Clear();
                return;
            }
            try
            {
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string query = "SELECT FirstName, LastName, Email FROM Customers WHERE CustomerID = @CustomerID";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@CustomerID", customerId);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                txtMailingCustomerFName_AdminMailingTab.Text = reader["FirstName"].ToString();
                                txtMailingCustomerLName_AdminMailingTab.Text = reader["LastName"].ToString();
                                txtMailingCustomerEmail_AdminMailingTab.Text = reader["Email"].ToString();
                            }
                            else
                            {
                                txtMailingCustomerFName_AdminMailingTab.Clear();
                                txtMailingCustomerLName_AdminMailingTab.Clear();
                                txtMailingCustomerEmail_AdminMailingTab.Clear();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to fetch customer details: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnMailingCustomerEmailDetailsSend_AdminMailingTab_Click(object sender, EventArgs e)
        {
            if (ComboBoxMailingCustomerID_AdminMailingTab.SelectedIndex <= 0 ||
                string.IsNullOrWhiteSpace(txtMailingCustomerFName_AdminMailingTab.Text) ||
                string.IsNullOrWhiteSpace(txtMailingCustomerLName_AdminMailingTab.Text) ||
                string.IsNullOrWhiteSpace(txtMailingCustomerEmail_AdminMailingTab.Text) ||
                string.IsNullOrWhiteSpace(txtMailingCustomerEmailSubject_AdminMailingTab.Text) ||
                string.IsNullOrWhiteSpace(txtMailingCustomerEmailBody.Text))
            {
                MessageBox.Show("Please fill out all required fields before sending.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int customerId;
            if (!int.TryParse(ComboBoxMailingCustomerID_AdminMailingTab.SelectedItem.ToString(), out customerId))
            {
                MessageBox.Show("Invalid Customer ID selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int? adminId = null;
            string adminEmail = null;
            try
            {
                using (SqlConnection conn = new SqlConnection(this.con.ConnectionString))
                {
                    conn.Open();
                    string adminQuery = "SELECT TOP 1 AdminID, Email FROM Admins WHERE IsActive = 1 AND Email IS NOT NULL AND Email <> ''";
                    using (SqlCommand adminCmd = new SqlCommand(adminQuery, conn))
                    {
                        using (SqlDataReader reader = adminCmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                adminId = Convert.ToInt32(reader["AdminID"]);
                                adminEmail = reader["Email"].ToString();
                            }
                        }
                    }
                }
            }
            catch { }

            string customerEmail = txtMailingCustomerEmail_AdminMailingTab.Text.Trim();
            string customerFirstName = txtMailingCustomerFName_AdminMailingTab.Text.Trim();
            string subject = txtMailingCustomerEmailSubject_AdminMailingTab.Text.Trim();
            string body = txtMailingCustomerEmailBody.Text.Trim();

            string fullBody = $"Dear {customerFirstName},\n\n" +
                body +
                "\n\nThank you for being a valued customer.\nBest regards,\nE-Shift Team";

            bool emailSent = false;
            string emailError = null;
            if (!string.IsNullOrEmpty(adminEmail))
            {
                try
                {
                    System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
                    mail.From = new System.Net.Mail.MailAddress(adminEmail);
                    mail.To.Add(customerEmail);
                    mail.Subject = subject;
                    mail.Body = fullBody;

                    System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("smtp.gmail.com", 587);
                    smtp.EnableSsl = true;
                    smtp.Credentials = new System.Net.NetworkCredential(adminEmail, Environment.GetEnvironmentVariable("ESHIFT_SMTP_APP_PASSWORD"));
                    smtp.Send(mail);
                    emailSent = true;
                }
                catch (Exception ex)
                {
                    emailError = ex.Message;
                }
            }
            else
            {
                emailError = "Admin email not found.";
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(this.con.ConnectionString))
                {
                    conn.Open();
                    string insertQuery = @"INSERT INTO CustomerMailingLogs 
                        (SentByAdminID, SentToCustomerID, CustomerFirstName, CustomerLastName, Subject, Body, SentDate)
                        VALUES (@SentByAdminID, @SentToCustomerID, @CustomerFirstName, @CustomerLastName, @Subject, @Body, GETDATE())";
                    using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                    {
                        if (adminId.HasValue)
                            cmd.Parameters.AddWithValue("@SentByAdminID", adminId.Value);
                        else
                            cmd.Parameters.AddWithValue("@SentByAdminID", DBNull.Value);
                        cmd.Parameters.AddWithValue("@SentToCustomerID", customerId);
                        cmd.Parameters.AddWithValue("@CustomerFirstName", customerFirstName);
                        cmd.Parameters.AddWithValue("@CustomerLastName", txtMailingCustomerLName_AdminMailingTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@Subject", subject);
                        cmd.Parameters.AddWithValue("@Body", body);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch { /* Ignore DB errors for now */ }

            if (emailSent)
            {
                MessageBox.Show("Email sent successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ComboBoxMailingCustomerID_AdminMailingTab.SelectedIndex = 0;
                txtMailingCustomerFName_AdminMailingTab.Clear();
                txtMailingCustomerLName_AdminMailingTab.Clear();
                txtMailingCustomerEmail_AdminMailingTab.Clear();
                txtMailingCustomerEmailSubject_AdminMailingTab.Clear();
                txtMailingCustomerEmailBody.Clear();
            }
            else
            {
                MessageBox.Show("Failed to send email: " + (emailError ?? "Unknown error."), "Email Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnMailingCustomerEmailDetailsClear_AdminMailingTab_Click(object sender, EventArgs e)
        {
            txtMailingCustomerFName_AdminMailingTab.Clear();
            txtMailingCustomerLName_AdminMailingTab.Clear();
            txtMailingCustomerEmail_AdminMailingTab.Clear();
            txtMailingCustomerEmailSubject_AdminMailingTab.Clear();
            txtMailingCustomerEmailBody.Clear();
            ComboBoxMailingCustomerID_AdminMailingTab.SelectedIndex = 0;
        }

        private void PopulateMailingDriverIDComboBox()
        {
            try
            {
                ComboBoxMailingDriverID_AdminMailingTab.Items.Clear();
                ComboBoxMailingDriverID_AdminMailingTab.Items.Add("Select Driver");
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string query = "SELECT DriverID FROM Drivers WHERE IsActive = 1 ORDER BY DriverID";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            ComboBoxMailingDriverID_AdminMailingTab.Items.Add(reader["DriverID"].ToString());
                        }
                    }
                }
                ComboBoxMailingDriverID_AdminMailingTab.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load Driver IDs: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ComboBoxMailingDriverID_AdminMailingTab_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ComboBoxMailingDriverID_AdminMailingTab.SelectedIndex <= 0)
            {
                txtMailingDriverFName.Clear();
                txtMailingDriverLName.Clear();
                txtMailingDriverEmail_AdminMailingTab.Clear();
                return;
            }
            int driverId;
            if (!int.TryParse(ComboBoxMailingDriverID_AdminMailingTab.SelectedItem.ToString(), out driverId))
            {
                txtMailingDriverFName.Clear();
                txtMailingDriverLName.Clear();
                txtMailingDriverEmail_AdminMailingTab.Clear();
                return;
            }
            try
            {
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string query = "SELECT FirstName, LastName, Email FROM Drivers WHERE DriverID = @DriverID";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@DriverID", driverId);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                txtMailingDriverFName.Text = reader["FirstName"].ToString();
                                txtMailingDriverLName.Text = reader["LastName"].ToString();
                                txtMailingDriverEmail_AdminMailingTab.Text = reader["Email"].ToString();
                            }
                            else
                            {
                                txtMailingDriverFName.Clear();
                                txtMailingDriverLName.Clear();
                                txtMailingDriverEmail_AdminMailingTab.Clear();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to fetch driver details: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnMaliningDriverEmailDetailsSend_AdminMailingTab_Click(object sender, EventArgs e)
        {
            if (ComboBoxMailingDriverID_AdminMailingTab.SelectedIndex <= 0 ||
                string.IsNullOrWhiteSpace(txtMailingDriverFName.Text) ||
                string.IsNullOrWhiteSpace(txtMailingDriverLName.Text) ||
                string.IsNullOrWhiteSpace(txtMailingDriverEmail_AdminMailingTab.Text) ||
                string.IsNullOrWhiteSpace(txtMailingDriverEmailSubject_AdminMailingTab.Text) ||
                string.IsNullOrWhiteSpace(txtMailingDriverEmailBody_AdminMailingTab.Text))
            {
                MessageBox.Show("Please fill out all required fields before sending.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int driverId;
            if (!int.TryParse(ComboBoxMailingDriverID_AdminMailingTab.SelectedItem.ToString(), out driverId))
            {
                MessageBox.Show("Invalid Driver ID selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int? adminId = null;
            string adminEmail = null;
            try
            {
                using (SqlConnection conn = new SqlConnection(this.con.ConnectionString))
                {
                    conn.Open();
                    string adminQuery = "SELECT TOP 1 AdminID, Email FROM Admins WHERE IsActive = 1 AND Email IS NOT NULL AND Email <> ''";
                    using (SqlCommand adminCmd = new SqlCommand(adminQuery, conn))
                    {
                        using (SqlDataReader reader = adminCmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                adminId = Convert.ToInt32(reader["AdminID"]);
                                adminEmail = reader["Email"].ToString();
                            }
                        }
                    }
                }
            }
            catch { }

            string driverEmail = txtMailingDriverEmail_AdminMailingTab.Text.Trim();
            string driverFirstName = txtMailingDriverFName.Text.Trim();
            string subject = txtMailingDriverEmailSubject_AdminMailingTab.Text.Trim();
            string body = txtMailingDriverEmailBody_AdminMailingTab.Text.Trim();

            string fullBody = $"Dear {driverFirstName},\n\n" +
                body +
                "\n\nThank you for your service.\nBest regards,\nE-Shift Team";

            bool emailSent = false;
            string emailError = null;
            if (!string.IsNullOrEmpty(adminEmail))
            {
                try
                {
                    System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
                    mail.From = new System.Net.Mail.MailAddress(adminEmail);
                    mail.To.Add(driverEmail);
                    mail.Subject = subject;
                    mail.Body = fullBody;

                    System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("smtp.gmail.com", 587);
                    smtp.EnableSsl = true;
                    smtp.Credentials = new System.Net.NetworkCredential(adminEmail, Environment.GetEnvironmentVariable("ESHIFT_SMTP_APP_PASSWORD"));
                    smtp.Send(mail);
                    emailSent = true;
                }
                catch (Exception ex)
                {
                    emailError = ex.Message;
                }
            }
            else
            {
                emailError = "Admin email not found.";
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string insertQuery = @"INSERT INTO DriverMailingLogs 
                        (SentByAdminID, SentToDriverID, DriverFirstName, DriverLastName, Subject, Body, SentDate)
                        VALUES (@SentByAdminID, @SentToDriverID, @DriverFirstName, @DriverLastName, @Subject, @Body, GETDATE())";
                    using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                    {
                        if (adminId.HasValue)
                            cmd.Parameters.AddWithValue("@SentByAdminID", adminId.Value);
                        else
                            cmd.Parameters.AddWithValue("@SentByAdminID", DBNull.Value);
                        cmd.Parameters.AddWithValue("@SentToDriverID", driverId);
                        cmd.Parameters.AddWithValue("@DriverFirstName", driverFirstName);
                        cmd.Parameters.AddWithValue("@DriverLastName", txtMailingDriverLName.Text.Trim());
                        cmd.Parameters.AddWithValue("@Subject", subject);
                        cmd.Parameters.AddWithValue("@Body", body);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch { /* Ignore DB errors for now */ }

            if (emailSent)
            {
                MessageBox.Show("Email sent successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ComboBoxMailingDriverID_AdminMailingTab.SelectedIndex = 0;
                txtMailingDriverFName.Clear();
                txtMailingDriverLName.Clear();
                txtMailingDriverEmail_AdminMailingTab.Clear();
                txtMailingDriverEmailSubject_AdminMailingTab.Clear();
                txtMailingDriverEmailBody_AdminMailingTab.Clear();
            }
            else
            {
                MessageBox.Show("Failed to send email: " + (emailError ?? "Unknown error."), "Email Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnMailingDriverEmailDetailsClear_AdminMailingTab_Click(object sender, EventArgs e)
        {
            ComboBoxMailingDriverID_AdminMailingTab.SelectedIndex = 0;
            txtMailingDriverFName.Clear();
            txtMailingDriverLName.Clear();
            txtMailingDriverEmail_AdminMailingTab.Clear();
            txtMailingDriverEmailSubject_AdminMailingTab.Clear();
            txtMailingDriverEmailBody_AdminMailingTab.Clear();
        }

        private void PopulateMailingAssistantIDComboBox()
        {
            try
            {
                ComboBoxMailingAssistantID.Items.Clear();
                ComboBoxMailingAssistantID.Items.Add("Select Assistant");
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string query = "SELECT AssistantID FROM Assistants WHERE IsActive = 1 ORDER BY AssistantID";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            ComboBoxMailingAssistantID.Items.Add(reader["AssistantID"].ToString());
                        }
                    }
                }
                ComboBoxMailingAssistantID.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load Assistant IDs: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ComboBoxMailingAssistantID_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ComboBoxMailingAssistantID.SelectedIndex <= 0)
            {
                txtMailingAssistantFName_AdminMailingTab.Clear();
                txtMailingAssistantLName_AdminMailingTab.Clear();
                txtMailingAssistantEmail_AdminMailingTab.Clear();
                return;
            }
            int assistantId;
            if (!int.TryParse(ComboBoxMailingAssistantID.SelectedItem.ToString(), out assistantId))
            {
                txtMailingAssistantFName_AdminMailingTab.Clear();
                txtMailingAssistantLName_AdminMailingTab.Clear();
                txtMailingAssistantEmail_AdminMailingTab.Clear();
                return;
            }
            try
            {
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string query = "SELECT FirstName, LastName, Email FROM Assistants WHERE AssistantID = @AssistantID";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@AssistantID", assistantId);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                txtMailingAssistantFName_AdminMailingTab.Text = reader["FirstName"].ToString();
                                txtMailingAssistantLName_AdminMailingTab.Text = reader["LastName"].ToString();
                                txtMailingAssistantEmail_AdminMailingTab.Text = reader["Email"].ToString();
                            }
                            else
                            {
                                txtMailingAssistantFName_AdminMailingTab.Clear();
                                txtMailingAssistantLName_AdminMailingTab.Clear();
                                txtMailingAssistantEmail_AdminMailingTab.Clear();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to fetch assistant details: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnMailingAssistantEmailDetailsSend_AdminMailingTab_Click(object sender, EventArgs e)
        {
            if (ComboBoxMailingAssistantID.SelectedIndex <= 0 ||
                string.IsNullOrWhiteSpace(txtMailingAssistantFName_AdminMailingTab.Text) ||
                string.IsNullOrWhiteSpace(txtMailingAssistantLName_AdminMailingTab.Text) ||
                string.IsNullOrWhiteSpace(txtMailingAssistantEmail_AdminMailingTab.Text) ||
                string.IsNullOrWhiteSpace(txtMailingAssistantEmailSubject_AdminMailingTab.Text) ||
                string.IsNullOrWhiteSpace(txtMailingAssistantEmailBody_AdminMailingTab.Text))
            {
                MessageBox.Show("Please fill out all required fields before sending.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int assistantId;
            if (!int.TryParse(ComboBoxMailingAssistantID.SelectedItem.ToString(), out assistantId))
            {
                MessageBox.Show("Invalid Assistant ID selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int? adminId = null;
            string adminEmail = null;
            try
            {
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string adminQuery = "SELECT TOP 1 AdminID, Email FROM Admins WHERE IsActive = 1 AND Email IS NOT NULL AND Email <> ''";
                    using (SqlCommand adminCmd = new SqlCommand(adminQuery, conn))
                    {
                        using (SqlDataReader reader = adminCmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                adminId = Convert.ToInt32(reader["AdminID"]);
                                adminEmail = reader["Email"].ToString();
                            }
                        }
                    }
                }
            }
            catch { }

            string assistantEmail = txtMailingAssistantEmail_AdminMailingTab.Text.Trim();
            string assistantFirstName = txtMailingAssistantFName_AdminMailingTab.Text.Trim();
            string assistantLastName = txtMailingAssistantLName_AdminMailingTab.Text.Trim();
            string subject = txtMailingAssistantEmailSubject_AdminMailingTab.Text.Trim();
            string body = txtMailingAssistantEmailBody_AdminMailingTab.Text.Trim();

            string fullBody = $"Dear {assistantFirstName},\n\n" +
                body +
                "\n\nThank you for your support.\nBest regards,\nE-Shift Team";

            bool emailSent = false;
            string emailError = null;
            if (!string.IsNullOrEmpty(adminEmail))
            {
                try
                {
                    System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
                    mail.From = new System.Net.Mail.MailAddress(adminEmail);
                    mail.To.Add(assistantEmail);
                    mail.Subject = subject;
                    mail.Body = fullBody;

                    System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("smtp.gmail.com", 587);
                    smtp.EnableSsl = true;
                    smtp.Credentials = new System.Net.NetworkCredential(adminEmail, Environment.GetEnvironmentVariable("ESHIFT_SMTP_APP_PASSWORD"));
                    smtp.Send(mail);
                    emailSent = true;
                }
                catch (Exception ex)
                {
                    emailError = ex.Message;
                }
            }
            else
            {
                emailError = "Admin email not found.";
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string insertQuery = @"INSERT INTO AssistantMailingLogs 
                        (SentByAdminID, SentToAssistantID, AssistantFirstName, AssistantLastName, Subject, Body, SentDate)
                        VALUES (@SentByAdminID, @SentToAssistantID, @AssistantFirstName, @AssistantLastName, @Subject, @Body, GETDATE())";
                    using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                    {
                        if (adminId.HasValue)
                            cmd.Parameters.AddWithValue("@SentByAdminID", adminId.Value);
                        else
                            cmd.Parameters.AddWithValue("@SentByAdminID", DBNull.Value);
                        cmd.Parameters.AddWithValue("@SentToAssistantID", assistantId);
                        cmd.Parameters.AddWithValue("@AssistantFirstName", assistantFirstName);
                        cmd.Parameters.AddWithValue("@AssistantLastName", assistantLastName);
                        cmd.Parameters.AddWithValue("@Subject", subject);
                        cmd.Parameters.AddWithValue("@Body", body);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch { /* Ignore DB errors for now */ }

            if (emailSent)
            {
                MessageBox.Show("Email sent successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ComboBoxMailingAssistantID.SelectedIndex = 0;
                txtMailingAssistantFName_AdminMailingTab.Clear();
                txtMailingAssistantLName_AdminMailingTab.Clear();
                txtMailingAssistantEmail_AdminMailingTab.Clear();
                txtMailingAssistantEmailSubject_AdminMailingTab.Clear();
                txtMailingAssistantEmailBody_AdminMailingTab.Clear();
            }
            else
            {
                MessageBox.Show("Failed to send email: " + (emailError ?? "Unknown error."), "Email Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnMailingAssistantEmailDetailsClewar_AdminMailingTab_Click(object sender, EventArgs e)
        {
            ComboBoxMailingAssistantID.SelectedIndex = 0;
            txtMailingAssistantFName_AdminMailingTab.Clear();
            txtMailingAssistantLName_AdminMailingTab.Clear();
            txtMailingAssistantEmail_AdminMailingTab.Clear();
            txtMailingAssistantEmailSubject_AdminMailingTab.Clear();
            txtMailingAssistantEmailBody_AdminMailingTab.Clear();
        }

        private void pictureBoxSettingsChnageProfilePhoto_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Select Assistant Profile Photo";
                ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var img = System.Drawing.Image.FromFile(ofd.FileName);
                        pictureBoxSettingsChnageProfilePhoto.Image = img;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failed to load image: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnSettingsDetailsUpdate_Click(object sender, EventArgs e)
        {
            if (ComboBoxSettingstitle_AdminSettingsTab.SelectedIndex < 0 ||
                string.IsNullOrWhiteSpace(txtSettingFName_AdminSettingsTab.Text) ||
                string.IsNullOrWhiteSpace(txtSettingsLName_AdminSettingsTab.Text) ||
                string.IsNullOrWhiteSpace(txtSettingsNIC_AdminSettingsTab.Text) ||
                string.IsNullOrWhiteSpace(txtSettingsPNumber_AdminSettingsTab.Text) ||
                string.IsNullOrWhiteSpace(txtSettingsAdress_AdminSettingsTab.Text) ||
                string.IsNullOrWhiteSpace(txtSettingsUsername_AdminSettingsTab.Text))
            {
                MessageBox.Show("Please fill out all required fields (Title, First Name, Last Name, NIC, Phone Number, Address, Username).", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string profilePicPath = null;
            if (pictureBoxSettingsChnageProfilePhoto.Image != null)
            {
                try
                {
                    string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    string eShiftFolder = System.IO.Path.Combine(appDataPath, "E-Shift-ProfilePics");
                    if (!System.IO.Directory.Exists(eShiftFolder))
                        System.IO.Directory.CreateDirectory(eShiftFolder);
                    string fileName = "admin_profile_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".png";
                    string savePath = System.IO.Path.Combine(eShiftFolder, fileName);
                    pictureBoxSettingsChnageProfilePhoto.Image.Save(savePath, System.Drawing.Imaging.ImageFormat.Png);
                    profilePicPath = savePath;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to save profile photo: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    profilePicPath = null;
                }
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string updateQuery = @"UPDATE Admins SET
                Title = @Title,
                FirstName = @FirstName,
                LastName = @LastName,
                NIC = @NIC,
                Username = @Username,
                PhoneNumber = @PhoneNumber,
                Address = @Address,
                ProfilePicturePath = @ProfilePicturePath";

                    using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@Title", ComboBoxSettingstitle_AdminSettingsTab.SelectedItem.ToString());
                        cmd.Parameters.AddWithValue("@FirstName", txtSettingFName_AdminSettingsTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@LastName", txtSettingsLName_AdminSettingsTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@NIC", txtSettingsNIC_AdminSettingsTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@Username", txtSettingsUsername_AdminSettingsTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@PhoneNumber", txtSettingsPNumber_AdminSettingsTab.Text.Trim());
                        cmd.Parameters.AddWithValue("@Address", txtSettingsAdress_AdminSettingsTab.Text.Trim());
                        if (!string.IsNullOrEmpty(profilePicPath))
                            cmd.Parameters.AddWithValue("@ProfilePicturePath", profilePicPath);
                        else
                            cmd.Parameters.AddWithValue("@ProfilePicturePath", DBNull.Value);

                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                        {
                            MessageBox.Show("Admin details updated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            LoadAdminProfilePicture();
                        }
                        else
                        {
                            MessageBox.Show("No admin record updated.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to update admin details: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadAdminProfilePicture()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string query = "SELECT ProfilePicturePath FROM Admins WHERE Username = @Username AND IsActive = 1";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@Username", _adminUsername);
                        object result = cmd.ExecuteScalar();
                        if (result != null && result != DBNull.Value)
                        {
                            string profilePicPath = result.ToString();
                            if (!string.IsNullOrEmpty(profilePicPath) && System.IO.File.Exists(profilePicPath))
                            {
                                using (var img = System.Drawing.Image.FromFile(profilePicPath))
                                {
                                    // Clone the image to avoid file lock issues
                                    CirclePictureBoxAdminProfilePic_AdminDashboard.Image = new Bitmap(img);
                                }
                            }
                            else
                            {
                                CirclePictureBoxAdminProfilePic_AdminDashboard.Image = null; // or set a default resource
                            }
                        }
                        else
                        {
                            CirclePictureBoxAdminProfilePic_AdminDashboard.Image = null; 
                        }
                    }
                }
            }
            catch (Exception)
            {
                CirclePictureBoxAdminProfilePic_AdminDashboard.Image = null; 
            }
        }

        private void dataGridViewDriverList_AdminDriverTab_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void PopulateDriverIDComboBox_AdminJobsTab()
        {
            try
            {
                ComboBoxDriverID_AdminJobsTab.Items.Clear();
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string query = "SELECT DriverID FROM Drivers WHERE IsActive = 1 ORDER BY DriverID";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            ComboBoxDriverID_AdminJobsTab.Items.Add(reader["DriverID"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load Driver IDs: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ComboBoxDriverID_AdminJobsTab_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ComboBoxDriverID_AdminJobsTab.SelectedItem != null)
            {
                int driverId;
                if (int.TryParse(ComboBoxDriverID_AdminJobsTab.SelectedItem.ToString(), out driverId))
                {
                    using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                    {
                        conn.Open();
                        string query = "SELECT FirstName, LastName, Email, PhoneNumber FROM Drivers WHERE DriverID = @DriverID";
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@DriverID", driverId);
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    txtDriverFName_AdminJobsTab.Text = reader["FirstName"].ToString();
                                    txtDriverLName.Text = reader["LastName"].ToString();
                                    txtDriverEmail_AdminJobsTab.Text = reader["Email"].ToString();
                                    txtDriverPNumber_AdminJobsTab.Text = reader["PhoneNumber"].ToString();
                                }
                                else
                                {
                                    txtDriverFName_AdminJobsTab.Clear();
                                    txtDriverLName.Clear();
                                    txtDriverEmail_AdminJobsTab.Clear();
                                    txtDriverPNumber_AdminJobsTab.Clear();
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                txtDriverFName_AdminJobsTab.Clear();
                txtDriverLName.Clear();
                txtDriverEmail_AdminJobsTab.Clear();
                txtDriverPNumber_AdminJobsTab.Clear();
            }
        }

        private void PopulateAssistantIDComboBox_AdminJobsTab()
        {
            try
            {
                ComboBoxAssistantID_AdminJobsTab.Items.Clear();
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string query = "SELECT AssistantID FROM Assistants WHERE IsActive = 1 ORDER BY AssistantID";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            ComboBoxAssistantID_AdminJobsTab.Items.Add(reader["AssistantID"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load Assistant IDs: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ComboBoxAssistantID_AdminJobsTab_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ComboBoxAssistantID_AdminJobsTab.SelectedItem != null)
            {
                int assistantId;
                if (int.TryParse(ComboBoxAssistantID_AdminJobsTab.SelectedItem.ToString(), out assistantId))
                {
                    using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                    {
                        conn.Open();
                        string query = "SELECT FirstName, LastName, Email, PhoneNumber FROM Assistants WHERE AssistantID = @AssistantID";
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@AssistantID", assistantId);
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    txtAssistantFName_AdminJobsTab.Text = reader["FirstName"].ToString();
                                    txtAssistantLName_AdminJobsTab.Text = reader["LastName"].ToString();
                                    txtAssistantEmail_AdminJobsTab.Text = reader["Email"].ToString();
                                    txtAssistantPNumber_AdminJobsTab.Text = reader["PhoneNumber"].ToString();
                                }
                                else
                                {
                                    txtAssistantFName_AdminJobsTab.Clear();
                                    txtAssistantLName_AdminJobsTab.Clear();
                                    txtAssistantEmail_AdminJobsTab.Clear();
                                    txtAssistantPNumber_AdminJobsTab.Clear();
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                txtAssistantFName_AdminJobsTab.Clear();
                txtAssistantLName_AdminJobsTab.Clear();
                txtAssistantEmail_AdminJobsTab.Clear();
                txtAssistantPNumber_AdminJobsTab.Clear();
            }
        }

        private void PopulateLorryIDComboBox_AdminJobsTab()
        {
            try
            {
                ComboBoxLorryID_AdminJobsTab.Items.Clear();
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string query = "SELECT LorryID FROM Lorries WHERE IsActive = 1 ORDER BY LorryID";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            ComboBoxLorryID_AdminJobsTab.Items.Add(reader["LorryID"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load Lorry IDs: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ComboBoxLorryID_AdminJobsTab_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ComboBoxLorryID_AdminJobsTab.SelectedItem != null)
            {
                int lorryId;
                if (int.TryParse(ComboBoxLorryID_AdminJobsTab.SelectedItem.ToString(), out lorryId))
                {
                    using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                    {
                        conn.Open();
                        string query = "SELECT PlateNumber, FuelType FROM Lorries WHERE LorryID = @LorryID";
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@LorryID", lorryId);
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    txtLorryPlateNumber_AdminJobsTab.Text = reader["PlateNumber"].ToString();
                                    txtLorryFuelType_AdminJobsTab.Text = reader["FuelType"].ToString();
                                }
                                else
                                {
                                    txtLorryPlateNumber_AdminJobsTab.Clear();
                                    txtLorryFuelType_AdminJobsTab.Clear();
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                txtLorryPlateNumber_AdminJobsTab.Clear();
                txtLorryFuelType_AdminJobsTab.Clear();
            }
        }

        private void PopulateContainerIDComboBox_AdminJobsTab()
        {
            try
            {
                ComboBoxContainerID_AdminJobsTab.Items.Clear();
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string query = "SELECT ContainerID FROM Containers WHERE IsActive = 1 ORDER BY ContainerID";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            ComboBoxContainerID_AdminJobsTab.Items.Add(reader["ContainerID"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load Container IDs: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ComboBoxContainerID_AdminJobsTab_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ComboBoxContainerID_AdminJobsTab.SelectedItem != null)
            {
                int containerId;
                if (int.TryParse(ComboBoxContainerID_AdminJobsTab.SelectedItem.ToString(), out containerId))
                {
                    using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                    {
                        conn.Open();
                        string query = "SELECT PlateNumber, Colour, Length, Width, Height, MaxWeight, TarWeight FROM Containers WHERE ContainerID = @ContainerID";
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@ContainerID", containerId);
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    txtContainerPlateNumber_AdminJobsTab.Text = reader["PlateNumber"].ToString();
                                    txtContainerColour.Text = reader["Colour"].ToString();
                                    txtContainerLength_AdminJobsTab.Text = reader["Length"].ToString();
                                    txtContainerHight_AdminJobsTab.Text = reader["Height"].ToString();
                                    txtContainerWidth_AdminJobsTab.Text = reader["Width"].ToString();
                                    txtContainerMaxWeight_AdminJobsTab.Text = reader["MaxWeight"].ToString();
                                    txtContainerTarWeight_AdminJobsTab.Text = reader["TarWeight"].ToString();
                                }
                                else
                                {
                                    txtContainerPlateNumber_AdminJobsTab.Clear();
                                    txtContainerColour.Clear();
                                    txtContainerLength_AdminJobsTab.Clear();
                                    txtContainerHight_AdminJobsTab.Clear();
                                    txtContainerWidth_AdminJobsTab.Clear();
                                    txtContainerMaxWeight_AdminJobsTab.Clear();
                                    txtContainerTarWeight_AdminJobsTab.Clear();
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                txtContainerPlateNumber_AdminJobsTab.Clear();
                txtContainerColour.Clear();
                txtContainerLength_AdminJobsTab.Clear();
                txtContainerHight_AdminJobsTab.Clear();
                txtContainerWidth_AdminJobsTab.Clear();
                txtContainerMaxWeight_AdminJobsTab.Clear();
                txtContainerTarWeight_AdminJobsTab.Clear();
            }
        }

        private void btnAcceptedRequestedJob_AdminMyDashboardTab_Click(object sender, EventArgs e)
        {
            if (dataGridViewMain_AdminMyDashboardTab.SelectedRows.Count > 0)
            {
                var row = dataGridViewMain_AdminMyDashboardTab.SelectedRows[0];
                var jobIdCell = row.Cells["JobID"];

                if (jobIdCell != null && jobIdCell.Value != null && int.TryParse(jobIdCell.Value.ToString(), out int jobId))
                {
                    try
                    {
                        con.Open();
                        string updateQuery = "UPDATE Jobs SET Status = @Status WHERE JobID = @JobID";
                        using (SqlCommand cmd = new SqlCommand(updateQuery, con))
                        {
                            cmd.Parameters.AddWithValue("@Status", "Accepted");
                            cmd.Parameters.AddWithValue("@JobID", jobId);
                            int rowsAffected = cmd.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Job request accepted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                LoadJobsGrid();
                                PopulateJobIDComboBox_AdminJobsTab();
                                ComboBoxJobID_AdminJobsTab.SelectedIndexChanged += ComboBoxJobID_AdminJobsTabb_SelectedIndexChanged;
                            }
                            else
                            {
                                MessageBox.Show("Failed to accept job request. No record found or status already accepted.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error accepting job request: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        con.Close();
                    }
                }
                else
                {
                    MessageBox.Show("Could not retrieve Job ID from the selected row.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Please select a job to accept.", "No Job Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void PopulateJobIDComboBox_AdminJobsTab()
        {
            try
            {
                ComboBoxJobID_AdminJobsTab.Items.Clear();
                using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                {
                    conn.Open();
                    string query = "SELECT JobID FROM Jobs WHERE Status = 'Accepted' ORDER BY JobID DESC";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            ComboBoxJobID_AdminJobsTab.Items.Add(reader["JobID"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load Job IDs: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ComboBoxJobID_AdminJobsTabb_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtCustomerID_AdminJobsTab.Clear();
            txtCustomerFName_AdminJobsTab.Clear();
            txtCustomerLName.Clear();
            txtCustomerPNumber_AdminJobsTab.Clear();
            txtCustomerEmail_AdminJobsTab.Clear();
            txtCustomerPickupAddress_AdminJobsTab.Clear();
            txtCustomerPickupCity_AdminJobsTab.Clear();
            txtCustomerDeliveryAddress_AdminJobsTab.Clear();
            txtCustomerDeliveryCity_AdminJobsTab.Clear();
            txtCustomerLocationType_AdminJobsTab.Clear();
            txtFirstProduct_AdminJobsTab.Clear();
            txtFirstProductMaxAmount_RequestAJob_AdminDashboard.Clear(); // Using the name provided
            txtSecondProduct_AdminJobsTab.Clear();
            txtSecondProductMaxAmount_AdminJobsTab.Clear();
            RadioButtonNotPickedUp_AdminJobsTab.Checked = false;
            RadioButtonPickUp_AdminJobsTab.Checked = false;
            RadioButtonDelivered_AdminJobsTab.Checked = false;

            if (ComboBoxJobID_AdminJobsTab.SelectedItem != null)
            {
                int jobId;
                if (int.TryParse(ComboBoxJobID_AdminJobsTab.SelectedItem.ToString(), out jobId))
                {
                    try
                    {
                        using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                        {
                            conn.Open();
                            string query = @"SELECT 
                                CustomerID, LorryID, DriverID, AssistantID, ContainerID, LoadStatus, FirstProduct, FirstProductQuantity, SecondProduct, SecondProductQuantity,
                                FirstName, LastName, PhoneNumber, Email, PickupAddress, PickupCity,
                                DeliveryAddress, DeliveryCity, LocationType, PickupDate, TotalDistance, EstimatedTime, JobCostPerDistance, TotalCost
                                FROM Jobs WHERE JobID = @JobID";
                            using (SqlCommand cmd = new SqlCommand(query, conn))
                            {
                                cmd.Parameters.AddWithValue("@JobID", jobId);
                                using (SqlDataReader reader = cmd.ExecuteReader())
                                {
                                    if (reader.Read())
                                    {
                                        txtCustomerID_AdminJobsTab.Text = reader["CustomerID"].ToString();
                                        txtCustomerFName_AdminJobsTab.Text = reader["FirstName"].ToString();
                                        txtCustomerLName.Text = reader["LastName"].ToString();
                                        txtCustomerPNumber_AdminJobsTab.Text = reader["PhoneNumber"].ToString();
                                        txtCustomerEmail_AdminJobsTab.Text = reader["Email"].ToString();
                                        txtCustomerPickupAddress_AdminJobsTab.Text = reader["PickupAddress"].ToString();
                                        txtCustomerPickupCity_AdminJobsTab.Text = reader["PickupCity"].ToString();
                                        txtCustomerDeliveryAddress_AdminJobsTab.Text = reader["DeliveryAddress"].ToString();
                                        txtJobPickupCity_AdminJobsTab.Text = reader["PickupCity"].ToString();
                                        txtJobDeliveryCity_AdminJobsTab.Text = reader["DeliveryCity"].ToString();
                                        txtCustomerDeliveryCity_AdminJobsTab.Text = reader["DeliveryCity"].ToString();
                                        txtCustomerLocationType_AdminJobsTab.Text = reader["LocationType"].ToString();
                                        txtFirstProduct_AdminJobsTab.Text = reader["FirstProduct"].ToString();
                                        txtFirstProductMaxAmount_RequestAJob_AdminDashboard.Text = reader["FirstProductQuantity"].ToString();
                                        txtSecondProduct_AdminJobsTab.Text = reader["SecondProduct"].ToString();
                                        txtSecondProductMaxAmount_AdminJobsTab.Text = reader["SecondProductQuantity"].ToString();

                                        txtJobTotalDistance_AdminJobsTab.Text = reader["TotalDistance"].ToString();
                                        txtEstimatedTime_AdminJobsTab.Text = reader["EstimatedTime"].ToString();
                                        txtJobCostPerdistance_AdminJobsTab.Text = reader["JobCostPerDistance"].ToString();
                                        txtJobTotalCost.Text = reader["TotalCost"].ToString();

                                        string loadStatus = reader["LoadStatus"].ToString();
                                        if (loadStatus == "Not PickedUp")
                                        {
                                            RadioButtonNotPickedUp_AdminJobsTab.Checked = true;
                                        }
                                        else if (loadStatus == "PickedUp")
                                        {
                                            RadioButtonPickUp_AdminJobsTab.Checked = true;
                                        }
                                        else if (loadStatus == "Delivered")
                                        {
                                            RadioButtonDelivered_AdminJobsTab.Checked = true;
                                        }
                                        if (reader["DriverID"] != DBNull.Value && !string.IsNullOrEmpty(reader["DriverID"].ToString()))
                                            ComboBoxDriverID_AdminJobsTab.SelectedItem = reader["DriverID"].ToString();
                                        if (reader["AssistantID"] != DBNull.Value && !string.IsNullOrEmpty(reader["AssistantID"].ToString()))
                                            ComboBoxAssistantID_AdminJobsTab.SelectedItem = reader["AssistantID"].ToString();
                                        if (reader["LorryID"] != DBNull.Value && !string.IsNullOrEmpty(reader["LorryID"].ToString()))
                                            ComboBoxLorryID_AdminJobsTab.SelectedItem = reader["LorryID"].ToString();
                                        if (reader["ContainerID"] != DBNull.Value && !string.IsNullOrEmpty(reader["ContainerID"].ToString()))
                                            ComboBoxContainerID_AdminJobsTab.SelectedItem = reader["ContainerID"].ToString();
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failed to load job details: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnDriverDetailsClear_AdminJobsTab_Click(object sender, EventArgs e)
        {
            ComboBoxDriverID_AdminJobsTab.SelectedIndex = -1;
            txtDriverFName_AdminJobsTab.Clear();
            txtDriverLName.Clear();
            txtDriverEmail_AdminJobsTab.Clear();
            txtDriverPNumber_AdminJobsTab.Clear();
        }

        private void btnAssistantDetailsClear_AdminJobsTab_Click(object sender, EventArgs e)
        {
            ComboBoxAssistantID_AdminJobsTab.SelectedIndex = -1;
            txtAssistantFName_AdminJobsTab.Clear();
            txtAssistantLName_AdminJobsTab.Clear();
            txtAssistantEmail_AdminJobsTab.Clear();
            txtAssistantPNumber_AdminJobsTab.Clear();
        }

        private void btnLorryDetailsClear_AdminJobsTab_Click(object sender, EventArgs e)
        {
            ComboBoxLorryID_AdminJobsTab.SelectedIndex = -1;
            txtLorryPlateNumber_AdminJobsTab.Clear();
            txtLorryFuelType_AdminJobsTab.Clear();
        }

        private void btnContainerDetailsClear_AdminJobsTab_Click(object sender, EventArgs e)
        {
            ComboBoxContainerID_AdminJobsTab.SelectedIndex = -1;
            txtContainerPlateNumber_AdminJobsTab.Clear();
            txtContainerColour.Clear();
            txtContainerLength_AdminJobsTab.Clear();
            txtContainerHight_AdminJobsTab.Clear();
            txtContainerWidth_AdminJobsTab.Clear();
            txtContainerMaxWeight_AdminJobsTab.Clear();
            txtContainerTarWeight_AdminJobsTab.Clear();
        }

        private void btnFullClear_AdminJobsTab_Click(object sender, EventArgs e)
        {
            ComboBoxContainerID_AdminJobsTab.SelectedIndex = -1;
            txtContainerPlateNumber_AdminJobsTab.Clear();
            txtContainerColour.Clear();
            txtContainerLength_AdminJobsTab.Clear();
            txtContainerHight_AdminJobsTab.Clear();
            txtContainerWidth_AdminJobsTab.Clear();
            txtContainerMaxWeight_AdminJobsTab.Clear();
            txtContainerTarWeight_AdminJobsTab.Clear();
            ComboBoxLorryID_AdminJobsTab.SelectedIndex = -1;
            txtLorryPlateNumber_AdminJobsTab.Clear();
            txtLorryFuelType_AdminJobsTab.Clear();
            ComboBoxAssistantID_AdminJobsTab.SelectedIndex = -1;
            txtAssistantFName_AdminJobsTab.Clear();
            txtAssistantLName_AdminJobsTab.Clear();
            txtAssistantEmail_AdminJobsTab.Clear();
            txtAssistantPNumber_AdminJobsTab.Clear();
            ComboBoxDriverID_AdminJobsTab.SelectedIndex = -1;
            txtDriverFName_AdminJobsTab.Clear();
            txtDriverLName.Clear();
            txtDriverEmail_AdminJobsTab.Clear();
            txtDriverPNumber_AdminJobsTab.Clear();
            txtJobTotalDistance_AdminJobsTab.Clear();
            txtEstimatedTime_AdminJobsTab.Clear();
            txtJobCostPerdistance_AdminJobsTab.Clear();
            txtJobTotalCost.Clear();
            txtCustomerFName_AdminJobsTab.Clear();
            txtCustomerLName.Clear();
            txtCustomerPNumber_AdminJobsTab.Clear();
            txtCustomerEmail_AdminJobsTab.Clear();
            txtCustomerPickupAddress_AdminJobsTab.Clear();
            txtCustomerPickupCity_AdminJobsTab.Clear();
            txtCustomerDeliveryAddress_AdminJobsTab.Clear();
            txtCustomerDeliveryCity_AdminJobsTab.Clear();
            txtCustomerLocationType_AdminJobsTab.Clear();
            txtFirstProduct_AdminJobsTab.Clear();
            txtSecondProduct_AdminJobsTab.Clear();
            txtFirstProductMaxAmount_RequestAJob_AdminDashboard.Clear();
            txtSecondProductMaxAmount_AdminJobsTab.Clear();
            ComboBoxJobID_AdminJobsTab.SelectedIndex = -1;
            txtJobPickupCity_AdminJobsTab.Clear();
            txtJobDeliveryCity_AdminJobsTab.Clear();
            txtJobTotalDistance_AdminJobsTab.Clear();
            txtEstimatedTime_AdminJobsTab.Clear();
            txtJobCostPerdistance_AdminJobsTab.Clear();
            txtJobTotalCost.Clear();
            if (webViewMap != null)
            {
                webViewMap.Source = new Uri("about:blank");
            }

        }

        private void btnFullProceed_AdminJobsTab_Click(object sender, EventArgs e)
        {
            if (ComboBoxJobID_AdminJobsTab.SelectedItem == null)
            {
                MessageBox.Show("Please select a Job ID to update.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int jobId;
            if (!int.TryParse(ComboBoxJobID_AdminJobsTab.SelectedItem.ToString(), out jobId))
            {
                MessageBox.Show("Invalid Job ID selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                int? lorryId = ComboBoxLorryID_AdminJobsTab.SelectedItem != null && int.TryParse(ComboBoxLorryID_AdminJobsTab.SelectedItem.ToString(), out int lId) ? lId : (int?)null;
                int? driverId = ComboBoxDriverID_AdminJobsTab.SelectedItem != null && int.TryParse(ComboBoxDriverID_AdminJobsTab.SelectedItem.ToString(), out int dId) ? dId : (int?)null;
                int? assistantId = ComboBoxAssistantID_AdminJobsTab.SelectedItem != null && int.TryParse(ComboBoxAssistantID_AdminJobsTab.SelectedItem.ToString(), out int aId) ? aId : (int?)null;
                int? containerId = ComboBoxContainerID_AdminJobsTab.SelectedItem != null && int.TryParse(ComboBoxContainerID_AdminJobsTab.SelectedItem.ToString(), out int cId) ? cId : (int?)null;

                string firstProduct = txtFirstProduct_AdminJobsTab.Text.Trim();
                int firstProductQuantity = int.TryParse(txtFirstProductMaxAmount_RequestAJob_AdminDashboard.Text.Trim(), out int fpq) ? fpq : 0;
                string secondProduct = txtSecondProduct_AdminJobsTab.Text.Trim();
                int secondProductQuantity = int.TryParse(txtSecondProductMaxAmount_AdminJobsTab.Text.Trim(), out int spq) ? spq : 0;

                string firstName = txtCustomerFName_AdminJobsTab.Text.Trim();
                string lastName = txtCustomerLName.Text.Trim();
                string phoneNumber = txtCustomerPNumber_AdminJobsTab.Text.Trim();
                string email = txtCustomerEmail_AdminJobsTab.Text.Trim();
                string pickupAddress = txtCustomerPickupAddress_AdminJobsTab.Text.Trim();
                string pickupCity = txtCustomerPickupCity_AdminJobsTab.Text.Trim();
                string deliveryAddress = txtCustomerDeliveryAddress_AdminJobsTab.Text.Trim();
                string deliveryCity = txtCustomerDeliveryCity_AdminJobsTab.Text.Trim();
                string locationType = txtCustomerLocationType_AdminJobsTab.Text.Trim();

                decimal totalDistance = decimal.TryParse(txtJobTotalDistance_AdminJobsTab.Text.Trim(), out decimal td) ? td : 0.00m;
                string estimatedTime = txtEstimatedTime_AdminJobsTab.Text.Trim();
                decimal jobCostPerDistance = decimal.TryParse(txtJobCostPerdistance_AdminJobsTab.Text.Trim(), out decimal jcpd) ? jcpd : 0.00m;
                decimal totalCost = decimal.TryParse(txtJobTotalCost.Text.Trim(), out decimal tc) ? tc : 0.00m;

                string loadStatus = "Not PickedUp"; 
                if (RadioButtonPickUp_AdminJobsTab.Checked)
                    loadStatus = "PickedUp";
                else if (RadioButtonDelivered_AdminJobsTab.Checked)
                    loadStatus = "Delivered";

                con.Open();
                string driverIdQuery = "SELECT DriverID FROM Jobs WHERE JobID = @JobID";
                SqlCommand driverCmd = new SqlCommand(driverIdQuery, con);
                driverCmd.Parameters.AddWithValue("@JobID", jobId);
                object existingDriverIdObj = driverCmd.ExecuteScalar();

                if (existingDriverIdObj != null && Convert.ToInt32(existingDriverIdObj) > 0)
                {
                    MessageBox.Show("This job request has already been processed (Accepted). If you wish to modify its details, please use the 'Update' button.", "Already Processed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string updateQuery = @"UPDATE Jobs SET
                    LorryID = @LorryID,
                    DriverID = @DriverID,
                    AssistantID = @AssistantID,
                    ContainerID = @ContainerID,
                    FirstProduct = @FirstProduct,
                    FirstProductQuantity = @FirstProductQuantity,
                    SecondProduct = @SecondProduct,
                    SecondProductQuantity = @SecondProductQuantity,
                    FirstName = @FirstName,
                    LastName = @LastName,
                    PhoneNumber = @PhoneNumber,
                    Email = @Email,
                    PickupAddress = @PickupAddress,
                    PickupCity = @PickupCity,
                    DeliveryAddress = @DeliveryAddress,
                    DeliveryCity = @DeliveryCity,
                    LocationType = @LocationType,
                    TotalDistance = @TotalDistance,
                    EstimatedTime = @EstimatedTime,
                    JobCostPerDistance = @JobCostPerDistance,
                    TotalCost = @TotalCost,
                    LoadStatus = @LoadStatus
                    WHERE JobID = @JobID";

                using (SqlCommand cmd = new SqlCommand(updateQuery, con))
                {
                    cmd.Parameters.AddWithValue("@LorryID", (object)lorryId ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@DriverID", (object)driverId ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@AssistantID", (object)assistantId ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@ContainerID", (object)containerId ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@FirstProduct", firstProduct);
                    cmd.Parameters.AddWithValue("@FirstProductQuantity", firstProductQuantity);
                    cmd.Parameters.AddWithValue("@SecondProduct", secondProduct);
                    cmd.Parameters.AddWithValue("@SecondProductQuantity", secondProductQuantity);
                    cmd.Parameters.AddWithValue("@FirstName", firstName);
                    cmd.Parameters.AddWithValue("@LastName", lastName);
                    cmd.Parameters.AddWithValue("@PhoneNumber", phoneNumber);
                    cmd.Parameters.AddWithValue("@Email", email);
                    cmd.Parameters.AddWithValue("@PickupAddress", pickupAddress);
                    cmd.Parameters.AddWithValue("@PickupCity", pickupCity);
                    cmd.Parameters.AddWithValue("@DeliveryAddress", deliveryAddress);
                    cmd.Parameters.AddWithValue("@DeliveryCity", deliveryCity);
                    cmd.Parameters.AddWithValue("@LocationType", locationType);
                    cmd.Parameters.AddWithValue("@TotalDistance", totalDistance);
                    cmd.Parameters.AddWithValue("@EstimatedTime", estimatedTime);
                    cmd.Parameters.AddWithValue("@JobCostPerDistance", jobCostPerDistance);
                    cmd.Parameters.AddWithValue("@TotalCost", totalCost);
                    cmd.Parameters.AddWithValue("@LoadStatus", loadStatus);
                    cmd.Parameters.AddWithValue("@JobID", jobId);

                    int rowsAffected = cmd.ExecuteNonQuery();

                    if (rowsAffected > 0)
                    {
                        MessageBox.Show("Job details successfully assigned!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadJobsGrid();
                    }
                    else
                    {
                        MessageBox.Show("No job found with the selected ID or no changes were made.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error updating job: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
            }
        }

        private void btnJobDetailsClear_AdminJobsTab_Click(object sender, EventArgs e)
        {
            txtJobTotalDistance_AdminJobsTab.Clear();
            txtEstimatedTime_AdminJobsTab.Clear();
            txtJobCostPerdistance_AdminJobsTab.Clear();
            txtJobTotalCost.Clear();
            if (webViewMap != null)
            {
                webViewMap.Source = new Uri("about:blank");
            }
        }

        private void btnFullUpdate_AdminJobsTab_Click(object sender, EventArgs e)
        {
            if (ComboBoxJobID_AdminJobsTab.SelectedItem == null)
            {
                MessageBox.Show("Please select a Job ID to update.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int jobId;
            if (!int.TryParse(ComboBoxJobID_AdminJobsTab.SelectedItem.ToString(), out jobId))
            {
                MessageBox.Show("Invalid Job ID selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                int? lorryId = ComboBoxLorryID_AdminJobsTab.SelectedItem != null && int.TryParse(ComboBoxLorryID_AdminJobsTab.SelectedItem.ToString(), out int lId) ? lId : (int?)null;
                int? driverId = ComboBoxDriverID_AdminJobsTab.SelectedItem != null && int.TryParse(ComboBoxDriverID_AdminJobsTab.SelectedItem.ToString(), out int dId) ? dId : (int?)null;
                int? assistantId = ComboBoxAssistantID_AdminJobsTab.SelectedItem != null && int.TryParse(ComboBoxAssistantID_AdminJobsTab.SelectedItem.ToString(), out int aId) ? aId : (int?)null;
                int? containerId = ComboBoxContainerID_AdminJobsTab.SelectedItem != null && int.TryParse(ComboBoxContainerID_AdminJobsTab.SelectedItem.ToString(), out int cId) ? cId : (int?)null;

                string firstProduct = txtFirstProduct_AdminJobsTab.Text.Trim();
                int firstProductQuantity = int.TryParse(txtFirstProductMaxAmount_RequestAJob_AdminDashboard.Text.Trim(), out int fpq) ? fpq : 0;
                string secondProduct = txtSecondProduct_AdminJobsTab.Text.Trim();
                int secondProductQuantity = int.TryParse(txtSecondProductMaxAmount_AdminJobsTab.Text.Trim(), out int spq) ? spq : 0;

                string firstName = txtCustomerFName_AdminJobsTab.Text.Trim();
                string lastName = txtCustomerLName.Text.Trim();
                string phoneNumber = txtCustomerPNumber_AdminJobsTab.Text.Trim();
                string email = txtCustomerEmail_AdminJobsTab.Text.Trim();
                string pickupAddress = txtCustomerPickupAddress_AdminJobsTab.Text.Trim();
                string pickupCity = txtCustomerPickupCity_AdminJobsTab.Text.Trim();
                string deliveryAddress = txtCustomerDeliveryAddress_AdminJobsTab.Text.Trim();
                string deliveryCity = txtCustomerDeliveryCity_AdminJobsTab.Text.Trim();
                string locationType = txtCustomerLocationType_AdminJobsTab.Text.Trim();

                decimal totalDistance = decimal.TryParse(txtJobTotalDistance_AdminJobsTab.Text.Trim(), out decimal td) ? td : 0.00m;
                string estimatedTime = txtEstimatedTime_AdminJobsTab.Text.Trim();
                decimal jobCostPerDistance = decimal.TryParse(txtJobCostPerdistance_AdminJobsTab.Text.Trim(), out decimal jcpd) ? jcpd : 0.00m;
                decimal totalCost = decimal.TryParse(txtJobTotalCost.Text.Trim(), out decimal tc) ? tc : 0.00m;

                string loadStatus = "Not PickedUp"; 
                if (RadioButtonPickUp_AdminJobsTab.Checked)
                    loadStatus = "PickedUp";
                else if (RadioButtonDelivered_AdminJobsTab.Checked)
                    loadStatus = "Delivered";

                con.Open();
                string updateQuery = @"UPDATE Jobs SET
                    LorryID = @LorryID,
                    DriverID = @DriverID,
                    AssistantID = @AssistantID,
                    ContainerID = @ContainerID,
                    FirstProduct = @FirstProduct,
                    FirstProductQuantity = @FirstProductQuantity,
                    SecondProduct = @SecondProduct,
                    SecondProductQuantity = @SecondProductQuantity,
                    FirstName = @FirstName,
                    LastName = @LastName,
                    PhoneNumber = @PhoneNumber,
                    Email = @Email,
                    PickupAddress = @PickupAddress,
                    PickupCity = @PickupCity,
                    DeliveryAddress = @DeliveryAddress,
                    DeliveryCity = @DeliveryCity,
                    LocationType = @LocationType,
                    TotalDistance = @TotalDistance,
                    EstimatedTime = @EstimatedTime,
                    JobCostPerDistance = @JobCostPerDistance,
                    TotalCost = @TotalCost,
                    LoadStatus = @LoadStatus
                    WHERE JobID = @JobID";

                using (SqlCommand cmd = new SqlCommand(updateQuery, con))
                {
                    cmd.Parameters.AddWithValue("@LorryID", (object)lorryId ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@DriverID", (object)driverId ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@AssistantID", (object)assistantId ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@ContainerID", (object)containerId ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@FirstProduct", firstProduct);
                    cmd.Parameters.AddWithValue("@FirstProductQuantity", firstProductQuantity);
                    cmd.Parameters.AddWithValue("@SecondProduct", secondProduct);
                    cmd.Parameters.AddWithValue("@SecondProductQuantity", secondProductQuantity);
                    cmd.Parameters.AddWithValue("@FirstName", firstName);
                    cmd.Parameters.AddWithValue("@LastName", lastName);
                    cmd.Parameters.AddWithValue("@PhoneNumber", phoneNumber);
                    cmd.Parameters.AddWithValue("@Email", email);
                    cmd.Parameters.AddWithValue("@PickupAddress", pickupAddress);
                    cmd.Parameters.AddWithValue("@PickupCity", pickupCity);
                    cmd.Parameters.AddWithValue("@DeliveryAddress", deliveryAddress);
                    cmd.Parameters.AddWithValue("@DeliveryCity", deliveryCity);
                    cmd.Parameters.AddWithValue("@LocationType", locationType);
                    cmd.Parameters.AddWithValue("@TotalDistance", totalDistance);
                    cmd.Parameters.AddWithValue("@EstimatedTime", estimatedTime);
                    cmd.Parameters.AddWithValue("@JobCostPerDistance", jobCostPerDistance);
                    cmd.Parameters.AddWithValue("@TotalCost", totalCost);
                    cmd.Parameters.AddWithValue("@LoadStatus", loadStatus);
                    cmd.Parameters.AddWithValue("@JobID", jobId);

                    int rowsAffected = cmd.ExecuteNonQuery();

                    if (rowsAffected > 0)
                    {
                        MessageBox.Show("Job details updated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadJobsGrid();
                    }
                    else
                    {
                        MessageBox.Show("No job found with the selected ID or no changes were made.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error updating job: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
            }
        }

        private void btnFullCancel_AdminJobsTab_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Proceeding the job is cancelled!", "Job Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Information);
            ComboBoxContainerID_AdminJobsTab.SelectedIndex = -1;
            txtContainerPlateNumber_AdminJobsTab.Clear();
            txtContainerColour.Clear();
            txtContainerLength_AdminJobsTab.Clear();
            txtContainerHight_AdminJobsTab.Clear();
            txtContainerWidth_AdminJobsTab.Clear();
            txtContainerMaxWeight_AdminJobsTab.Clear();
            txtContainerTarWeight_AdminJobsTab.Clear();
            ComboBoxLorryID_AdminJobsTab.SelectedIndex = -1;
            txtLorryPlateNumber_AdminJobsTab.Clear();
            txtLorryFuelType_AdminJobsTab.Clear();
            ComboBoxAssistantID_AdminJobsTab.SelectedIndex = -1;
            txtAssistantFName_AdminJobsTab.Clear();
            txtAssistantLName_AdminJobsTab.Clear();
            txtAssistantEmail_AdminJobsTab.Clear();
            txtAssistantPNumber_AdminJobsTab.Clear();
            ComboBoxDriverID_AdminJobsTab.SelectedIndex = -1;
            txtDriverFName_AdminJobsTab.Clear();
            txtDriverLName.Clear();
            txtDriverEmail_AdminJobsTab.Clear();
            txtDriverPNumber_AdminJobsTab.Clear();
            txtJobTotalDistance_AdminJobsTab.Clear();
            txtEstimatedTime_AdminJobsTab.Clear();
            txtJobCostPerdistance_AdminJobsTab.Clear();
            txtJobTotalCost.Clear();
            txtCustomerFName_AdminJobsTab.Clear();
            txtCustomerLName.Clear();
            txtCustomerPNumber_AdminJobsTab.Clear();
            txtCustomerEmail_AdminJobsTab.Clear();
            txtCustomerPickupAddress_AdminJobsTab.Clear();
            txtCustomerPickupCity_AdminJobsTab.Clear();
            txtCustomerDeliveryAddress_AdminJobsTab.Clear();
            txtCustomerDeliveryCity_AdminJobsTab.Clear();
            txtCustomerLocationType_AdminJobsTab.Clear();
            txtFirstProduct_AdminJobsTab.Clear();
            txtSecondProduct_AdminJobsTab.Clear();
            txtFirstProductMaxAmount_RequestAJob_AdminDashboard.Clear();
            txtSecondProductMaxAmount_AdminJobsTab.Clear();
            ComboBoxJobID_AdminJobsTab.SelectedIndex = -1;
            txtJobPickupCity_AdminJobsTab.Clear();
            txtJobDeliveryCity_AdminJobsTab.Clear();
            txtJobTotalDistance_AdminJobsTab.Clear();
            txtEstimatedTime_AdminJobsTab.Clear();
            txtJobCostPerdistance_AdminJobsTab.Clear();
            txtJobTotalCost.Clear();
        }

        // Event handler for updating cost per distance when total distance changes
        private void txtJobTotalDistance_AdminJobsTab_TextChanged(object sender, EventArgs e)
        {
            if (decimal.TryParse(txtJobTotalDistance_AdminJobsTab.Text.Trim(), out decimal distance))
            {
                decimal cost = distance * 100;
                txtJobCostPerdistance_AdminJobsTab.Text = cost.ToString("N2");
                // Calculate total cost with 2% tax
                decimal totalCost = cost + (cost * 0.02m);
                txtJobTotalCost.Text = totalCost.ToString("N2");
            }
            else
            {
                txtJobCostPerdistance_AdminJobsTab.Text = string.Empty;
                txtJobTotalCost.Text = string.Empty;
            }
        }

        private void btnViewJob_AdminMyDashboardTab_Click(object sender, EventArgs e)
        {
            if (dataGridViewMain_AdminMyDashboardTab.SelectedRows.Count > 0)
            {
                var row = dataGridViewMain_AdminMyDashboardTab.SelectedRows[0];
                // Restriction: Only allow if Status is 'Accepted'
                var statusObj = row.Cells["Status"].Value;
                if (statusObj == null || statusObj.ToString() != "Accepted")
                {
                    MessageBox.Show("Accept the specific job in order to view its details.", "Job Not Accepted", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                // Switch to Jobs tab
                TabControlAdminDashboard.SelectedTab = tabPageAdminJobsTab;

                // Get JobID from selected row
                var jobIdObj = row.Cells["JobID"].Value;
                if (jobIdObj == null || !int.TryParse(jobIdObj.ToString(), out int jobId))
                    return;

                try
                {
                    using (SqlConnection conn = new SqlConnection(con.ConnectionString))
                    {
                        conn.Open();
                        string query = @"SELECT 
                            CustomerID, LorryID, DriverID, AssistantID, ContainerID, LoadStatus, FirstProduct, FirstProductQuantity, SecondProduct, SecondProductQuantity,
                            FirstName, LastName, PhoneNumber, Email, PickupAddress, PickupCity,
                            DeliveryAddress, DeliveryCity, LocationType, PickupDate, TotalDistance, EstimatedTime, JobCostPerDistance, TotalCost
                            FROM Jobs WHERE JobID = @JobID";
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@JobID", jobId);
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    txtCustomerID_AdminJobsTab.Text = reader["CustomerID"].ToString();
                                    txtCustomerFName_AdminJobsTab.Text = reader["FirstName"].ToString();
                                    txtCustomerLName.Text = reader["LastName"].ToString();
                                    txtCustomerPNumber_AdminJobsTab.Text = reader["PhoneNumber"].ToString();
                                    txtCustomerEmail_AdminJobsTab.Text = reader["Email"].ToString();
                                    txtCustomerPickupAddress_AdminJobsTab.Text = reader["PickupAddress"].ToString();
                                    txtCustomerPickupCity_AdminJobsTab.Text = reader["PickupCity"].ToString();
                                    txtCustomerDeliveryAddress_AdminJobsTab.Text = reader["DeliveryAddress"].ToString();
                                    txtJobPickupCity_AdminJobsTab.Text = reader["PickupCity"].ToString();
                                    txtJobDeliveryCity_AdminJobsTab.Text = reader["DeliveryCity"].ToString();
                                    txtCustomerDeliveryCity_AdminJobsTab.Text = reader["DeliveryCity"].ToString();
                                    txtCustomerLocationType_AdminJobsTab.Text = reader["LocationType"].ToString();
                                    txtFirstProduct_AdminJobsTab.Text = reader["FirstProduct"].ToString();
                                    txtFirstProductMaxAmount_RequestAJob_AdminDashboard.Text = reader["FirstProductQuantity"].ToString();
                                    txtSecondProduct_AdminJobsTab.Text = reader["SecondProduct"].ToString();
                                    txtSecondProductMaxAmount_AdminJobsTab.Text = reader["SecondProductQuantity"].ToString();

                                    txtJobTotalDistance_AdminJobsTab.Text = reader["TotalDistance"].ToString();
                                    txtEstimatedTime_AdminJobsTab.Text = reader["EstimatedTime"].ToString();
                                    txtJobCostPerdistance_AdminJobsTab.Text = reader["JobCostPerDistance"].ToString();
                                    txtJobTotalCost.Text = reader["TotalCost"].ToString();

                                    string loadStatus = reader["LoadStatus"].ToString();
                                    if (loadStatus == "Not PickedUp")
                                        RadioButtonNotPickedUp_AdminJobsTab.Checked = true;
                                    else if (loadStatus == "PickedUp")
                                        RadioButtonPickUp_AdminJobsTab.Checked = true;
                                    else if (loadStatus == "Delivered")
                                        RadioButtonDelivered_AdminJobsTab.Checked = true;

                                    if (reader["DriverID"] != DBNull.Value && !string.IsNullOrEmpty(reader["DriverID"].ToString()))
                                    {
                                        ComboBoxDriverID_AdminJobsTab.SelectedItem = reader["DriverID"].ToString();
                                        ComboBoxDriverID_AdminJobsTab_SelectedIndexChanged(ComboBoxDriverID_AdminJobsTab, EventArgs.Empty);
                                    }
                                    if (reader["AssistantID"] != DBNull.Value && !string.IsNullOrEmpty(reader["AssistantID"].ToString()))
                                    {
                                        ComboBoxAssistantID_AdminJobsTab.SelectedItem = reader["AssistantID"].ToString();
                                        ComboBoxAssistantID_AdminJobsTab_SelectedIndexChanged(ComboBoxAssistantID_AdminJobsTab, EventArgs.Empty);
                                    }
                                    if (reader["LorryID"] != DBNull.Value && !string.IsNullOrEmpty(reader["LorryID"].ToString()))
                                    {
                                        ComboBoxLorryID_AdminJobsTab.SelectedItem = reader["LorryID"].ToString();
                                        ComboBoxLorryID_AdminJobsTab_SelectedIndexChanged(ComboBoxLorryID_AdminJobsTab, EventArgs.Empty);
                                    }
                                    if (reader["ContainerID"] != DBNull.Value && !string.IsNullOrEmpty(reader["ContainerID"].ToString()))
                                    {
                                        ComboBoxContainerID_AdminJobsTab.SelectedItem = reader["ContainerID"].ToString();
                                        ComboBoxContainerID_AdminJobsTab_SelectedIndexChanged(ComboBoxContainerID_AdminJobsTab, EventArgs.Empty);
                                    }
                                    ComboBoxJobID_AdminJobsTab.SelectedItem = jobId.ToString();
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to load job details: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnSettingsDetailsClear_AdminSettingsTab_Click(object sender, EventArgs e)
        {
            ComboBoxSettingstitle_AdminSettingsTab.SelectedIndex = -1;
            txtSettingFName_AdminSettingsTab.Clear();
            txtSettingsLName_AdminSettingsTab.Clear();
            txtSettingsNIC_AdminSettingsTab.Clear();
            txtSettingsPNumber_AdminSettingsTab.Clear();
            txtSettingsAdress_AdminSettingsTab.Clear();
            txtSettingsUsername_AdminSettingsTab.Clear();
            pictureBoxSettingsChnageProfilePhoto.Image = null;
        }

        private void Admin_Dashboard_MapInit()
        {
            if (webViewMap != null)
            {
                webViewMap.WebMessageReceived += webViewMap_WebMessageReceived;
            }
            txtJobPickupCity_AdminJobsTab.TextChanged += AddressFields_TextChanged;
            txtJobDeliveryCity_AdminJobsTab.TextChanged += AddressFields_TextChanged;
        }

        private void AddressFields_TextChanged(object sender, EventArgs e)
        {
            UpdateMap();
        }

        private void webViewMap_WebMessageReceived(object sender, Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                var options = new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                var data = System.Text.Json.JsonSerializer.Deserialize<MapResult>(e.WebMessageAsJson, options);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error receiving map data: " + ex.Message);
            }
        }

        private class MapResult
        {
            public string distance { get; set; }
            public string duration { get; set; }
        }

        private void UpdateMap()
        {
            string pickup = txtJobPickupCity_AdminJobsTab.Text.Trim();
            string drop = txtJobDeliveryCity_AdminJobsTab.Text.Trim();

            if (!string.IsNullOrEmpty(pickup) && !string.IsNullOrEmpty(drop))
            {
                string html = GenerateLeafletMapHtml(pickup + ", Sri Lanka", drop + ", Sri Lanka");
                string tempFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "map.html");
                System.IO.File.WriteAllText(tempFilePath, html);
                webViewMap.Source = new Uri(tempFilePath);
            }
        }

        private string GenerateLeafletMapHtml(string origin, string destination)
        {
            return $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8' />
    <title>OpenStreetMap Route</title>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <style>
        html, body, #map {{ height: 100%; margin: 0; padding: 0; }}
        #info {{ padding: 10px; font-family: Arial, sans-serif; }}
    </style>
    <link rel='stylesheet' href='https://unpkg.com/leaflet@1.9.3/dist/leaflet.css' />
    <script src='https://unpkg.com/leaflet@1.9.3/dist/leaflet.js'></script>
    <script src='https://unpkg.com/axios/dist/axios.min.js'></script>
</head>
<body>
    <div id='map'></div>
    <div id='info'></div>
    <script>
        var map = L.map('map').setView([7.8731, 80.7718], 7);

        L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
            maxZoom: 19,
            attribution: '© OpenStreetMap contributors'
        }}).addTo(map);

        function getCoordinates(address) {{
            var fullAddress = address.trim();
            if (!/sri lanka$/i.test(fullAddress)) {{
                fullAddress += ', Sri Lanka';
            }}
            return axios.get('https://nominatim.openstreetmap.org/search', {{
                params: {{
                    q: fullAddress,
                    format: 'json',
                    limit: 1,
                    countrycodes: 'LK'
                }},
                headers: {{
                    'Accept-Language': 'en'
                }}
            }}).then(function(response) {{
                if (response.data.length > 0) {{
                    return {{
                        lat: parseFloat(response.data[0].lat),
                        lon: parseFloat(response.data[0].lon)
                    }};
                }} else {{
                    throw new Error('No results for ' + fullAddress);
                }}
            }});
        }}

        function drawRoute(from, to) {{
            if (from.lat === to.lat && from.lon === to.lon) {{
                document.getElementById('info').innerText = 'Pickup and Drop-off are the same.';
                return;
            }}

            axios.get(`https://router.project-osrm.org/route/v1/driving/${{from.lon}},${{from.lat}};${{to.lon}},${{to.lat}}?overview=full&geometries=geojson`)
                .then(function(res) {{
                    var route = res.data.routes[0];
                    var coords = route.geometry.coordinates.map(p => [p[1], p[0]]);
                    var line = L.polyline(coords, {{ color: 'blue' }}).addTo(map);
                    map.fitBounds(line.getBounds());
                    L.marker([from.lat, from.lon]).addTo(map).bindPopup('Pickup').openPopup();
                    L.marker([to.lat, to.lon]).addTo(map).bindPopup('Drop-off').openPopup();

                    var distanceKm = route.distance / 1000;
                    var speedKmh = 50;
                    var durationMin = (distanceKm / speedKmh) * 60;

                    document.getElementById('info').innerHTML =
                        'Distance: ' + distanceKm.toFixed(2) + ' km<br>' +
                        'Duration: ' + durationMin.toFixed(1) + ' minutes (at 50 km/h)';

                    if (window.chrome && window.chrome.webview) {{
                        window.chrome.webview.postMessage({{
                            distance: distanceKm.toFixed(2) + ' km',
                            duration: durationMin.toFixed(1) + ' minutes (at 50 km/h)'
                        }});
                    }}
                }})
                .catch(function(err) {{
                    document.getElementById('info').innerText = 'Routing error: ' + err.message;
                }});
        }}

        const origin = '{origin}';
        const destination = '{destination}';

        Promise.all([getCoordinates(origin), getCoordinates(destination)])
            .then(function(results) {{
                drawRoute(results[0], results[1]);
            }})
            .catch(function(error) {{
                document.getElementById('info').innerText = 'Geocoding error: ' + error.message;
            }});
    </script>
</body>
</html>";
        }
    }
}