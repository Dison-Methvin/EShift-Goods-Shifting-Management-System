using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace E_Shift_V02
{
    public partial class Customer_Dashboard : Form
    {
        // Store JobID for deletion after RequestJobAgain is used
        private int? _pendingJobIdForDelete = null;
        private string _username;

        private Timer timerDateTime;

        public Customer_Dashboard(string username)
        {
            InitializeComponent();
            txtOldPassword_Settings_CustomerDashboard.PasswordChar = '●';
            txtNewPassword_Settings_CustomerDashboard.PasswordChar = '●';
            txtConfirmPassword_Settings_CustomerDashboard.PasswordChar = '●';

            _username = username;
            lblCustomerUsername_CustomerDashboard.Text = _username;
            this.Load += Customer_Dashboard_Load;

            timerDateTime = new Timer();
            timerDateTime.Interval = 1000; // 1 second
            timerDateTime.Tick += TimerDateTime_Tick;
            timerDateTime.Tick += (s, e) => FetchAndDisplayAcceptedJobsCount(); 
            timerDateTime.Tick += (s, e) => FetchAndDisplayRejectedJobsCount(); 
            timerDateTime.Start();

            btnRequestJobAgain_MyDashboard_CustomerDashboard.Enabled = false;
            dataGridViewMainGrid_MyDashboard_CustomerDashboard.SelectionChanged += DataGridViewMainGrid_MyDashboard_CustomerDashboard_SelectionChanged;
            btnRequestJobAgain_MyDashboard_CustomerDashboard.Click += btnRequestJobAgain_MyDashboard_CustomerDashboard_Click;

            FetchAndDisplayAcceptedJobsCount();
            FetchAndDisplayRejectedJobsCount();
        }
        SqlConnection con = new SqlConnection("Data Source=DISON_PC\\SQLEXPRESS;Initial Catalog=ESHIFT;Integrated Security=True;Encrypt=False");


        private void DataGridViewMainGrid_MyDashboard_CustomerDashboard_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridViewMainGrid_MyDashboard_CustomerDashboard.SelectedRows.Count > 0)
            {
                var row = dataGridViewMainGrid_MyDashboard_CustomerDashboard.SelectedRows[0];
                var statusCell = row.Cells["Status"];
                if (statusCell != null && statusCell.Value != null && statusCell.Value.ToString() == "Pending")
                {
                    btnRequestJobAgain_MyDashboard_CustomerDashboard.Enabled = true;
                }
                else
                {
                    btnRequestJobAgain_MyDashboard_CustomerDashboard.Enabled = false;
                }
            }
            else
            {
                btnRequestJobAgain_MyDashboard_CustomerDashboard.Enabled = false;
            }
        }

        private void TimerDateTime_Tick(object sender, EventArgs e)
        {
            TimeZoneInfo sriLankaTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Sri Lanka Standard Time");
            DateTime sriLankaTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, sriLankaTimeZone);
            lblDateTime_CustomerDashboard.Text = sriLankaTime.ToString("HH:mm:ss");
        }

        private void Customer_Dashboard_Load(object sender, EventArgs e)
        {
            LoadDashboardData();
            FetchAndDisplayAcceptedJobsCount();
            FetchAndDisplayRejectedJobsCount();
        }

        private void FetchAndDisplayAcceptedJobsCount()
        {
            try
            {
                con.Open();
                int customerId = -1;
                using (SqlCommand getIdCmd = new SqlCommand("SELECT CustomerID FROM Customers WHERE Username=@username", con))
                {
                    getIdCmd.Parameters.AddWithValue("@username", _username);
                    object idResult = getIdCmd.ExecuteScalar();
                    if (idResult != null && idResult != DBNull.Value)
                        customerId = Convert.ToInt32(idResult);
                }
                int acceptedCount = 0;
                if (customerId != -1)
                {
                    using (SqlCommand countCmd = new SqlCommand("SELECT COUNT(*) FROM Jobs WHERE CustomerID = @CustomerID AND Status = 'Accepted'", con))
                    {
                        countCmd.Parameters.AddWithValue("@CustomerID", customerId);
                        object countResult = countCmd.ExecuteScalar();
                        if (countResult != null && countResult != DBNull.Value)
                            acceptedCount = Convert.ToInt32(countResult);
                    }
                }
                lblJobsAcceptedNumber_MyDashboard_CustomerDashboard.Text = acceptedCount.ToString("D2");
            }
            catch
            {
                lblJobsAcceptedNumber_MyDashboard_CustomerDashboard.Text = "00";
            }
            finally
            {
                con.Close();
            }
        }

        private void FetchAndDisplayRejectedJobsCount()
        {
            try
            {
                con.Open();
                int customerId = -1;
                using (SqlCommand getIdCmd = new SqlCommand("SELECT CustomerID FROM Customers WHERE Username=@username", con))
                {
                    getIdCmd.Parameters.AddWithValue("@username", _username);
                    object idResult = getIdCmd.ExecuteScalar();
                    if (idResult != null && idResult != DBNull.Value)
                        customerId = Convert.ToInt32(idResult);
                }
                int rejectedCount = 0;
                if (customerId != -1)
                {
                    using (SqlCommand countCmd = new SqlCommand("SELECT COUNT(*) FROM Jobs WHERE CustomerID = @CustomerID AND Status = 'Rejected'", con))
                    {
                        countCmd.Parameters.AddWithValue("@CustomerID", customerId);
                        object countResult = countCmd.ExecuteScalar();
                        if (countResult != null && countResult != DBNull.Value)
                            rejectedCount = Convert.ToInt32(countResult);
                    }
                }
                lblJobsRejectedNumber_MyDashboard_CustomerDashboard.Text = rejectedCount.ToString("D2");
            }
            catch
            {
                lblJobsRejectedNumber_MyDashboard_CustomerDashboard.Text = "00";
            }
            finally
            {
                con.Close();
            }
        }

        private void LoadDashboardData()
        {
            try
            {
                con.Open();
                // Load profile photo
                SqlCommand cmd = new SqlCommand("SELECT ProfilePhoto FROM Customers WHERE Username=@username", con);
                cmd.Parameters.AddWithValue("@username", _username);
                object result = cmd.ExecuteScalar();
                if (result != DBNull.Value && result != null)
                {
                    byte[] imgBytes = (byte[])result;
                    using (var ms = new System.IO.MemoryStream(imgBytes))
                    {
                        CirclePictureBoxCustomerUsername_CustomerDashboard.Image = System.Drawing.Image.FromStream(ms);
                    }
                }
                else
                {
                    CirclePictureBoxCustomerUsername_CustomerDashboard.Image = null;
                }

                // Load pickup cities into ComboBox
                SqlCommand cityCmd = new SqlCommand("SELECT CityName FROM PickUpCity", con);
                using (SqlDataReader reader = cityCmd.ExecuteReader())
                {
                    ComboBoxPickupCity_RequestAJob_CustomerDashboard.Items.Clear();
                    while (reader.Read())
                    {
                        ComboBoxPickupCity_RequestAJob_CustomerDashboard.Items.Add(reader["CityName"].ToString());
                    }
                }

                // Load delivery cities into ComboBox
                SqlCommand deliveryCityCmd = new SqlCommand("SELECT CityName FROM DeliveryCity", con);
                using (SqlDataReader reader = deliveryCityCmd.ExecuteReader())
                {
                    ComboBoxDeliveryCity_RequestAJob_CustomerDashboard.Items.Clear();
                    while (reader.Read())
                    {
                        ComboBoxDeliveryCity_RequestAJob_CustomerDashboard.Items.Add(reader["CityName"].ToString());
                    }
                }

                // Load location types into ComboBox
                SqlCommand locationTypeCmd = new SqlCommand("SELECT TypeName FROM LocationType", con);
                using (SqlDataReader reader = locationTypeCmd.ExecuteReader())
                {
                    ComboBoxLocationType_RequestAJob_CustomerDashboard.Items.Clear();
                    while (reader.Read())
                    {
                        ComboBoxLocationType_RequestAJob_CustomerDashboard.Items.Add(reader["TypeName"].ToString());
                    }
                }

                // Load product names into FirstProduct and SecondProduct ComboBoxes
                SqlCommand productCmd = new SqlCommand("SELECT ProductName FROM ProductTypes", con);
                using (SqlDataReader reader = productCmd.ExecuteReader())
                {
                    ComboBoxFirstProduct_RequestAJob_CustomerDashboard.Items.Clear();
                    ComboBoxSecondProduct_RequestAJob_CustomerDashboard.Items.Clear();
                    while (reader.Read())
                    {
                        string productName = reader["ProductName"].ToString();
                        ComboBoxFirstProduct_RequestAJob_CustomerDashboard.Items.Add(productName);
                        ComboBoxSecondProduct_RequestAJob_CustomerDashboard.Items.Add(productName);
                    }
                }

                // Load job requests for this customer into DataGridView
                int customerId = -1;
                using (SqlCommand getIdCmd = new SqlCommand("SELECT CustomerID FROM Customers WHERE Username=@username", con))
                {
                    getIdCmd.Parameters.AddWithValue("@username", _username);
                    object idResult = getIdCmd.ExecuteScalar();
                    if (idResult != null && idResult != DBNull.Value)
                        customerId = Convert.ToInt32(idResult);
                }
                if (customerId != -1)
                {
                    // Fetch all required columns from Jobs table for this customer
                    SqlCommand jobsCmd = new SqlCommand(@"
                        SELECT JobID, CustomerID, FirstProduct, FirstProductQuantity, SecondProduct, SecondProductQuantity, FirstName, LastName, PhoneNumber, Email, PickupAddress, PickupCity, DeliveryAddress, DeliveryCity, LocationType, PickupDate, TotalDistance, EstimatedTime, JobCostPerDistance, TotalCost, Status, CreatedDate, LoadStatus
                        FROM Jobs
                        WHERE CustomerID = @CustomerID
                        ORDER BY CreatedDate DESC", con);
                    jobsCmd.Parameters.AddWithValue("@CustomerID", customerId);
                    SqlDataAdapter da = new SqlDataAdapter(jobsCmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridViewMainGrid_MyDashboard_CustomerDashboard.DataSource = dt;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load dashboard data: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
            }
        }

        private void btnLogout_RequestAJob_CustomerDashboard_Click(object sender, EventArgs e)
        {
            Role_Selection signUpForm = new Role_Selection();
            signUpForm.Show();
            this.Hide();
        }

        private void btnRequest_RequestAJob_CustomerDashboard_Click_1(object sender, EventArgs e)
        {
            // Validate required fields
            if (string.IsNullOrWhiteSpace(txtFname_RequestAJob_CustomerDashboard.Text) ||
                string.IsNullOrWhiteSpace(txtLName_RequestAJob_CustomerDashboard.Text) ||
                string.IsNullOrWhiteSpace(txtPNumber_RequestAJob_CustomerDashboard.Text) ||
                string.IsNullOrWhiteSpace(txtEmail_RequestAJob_CustomerDashboard.Text) ||
                string.IsNullOrWhiteSpace(txtPuckupAddress_RequestAJob_CustomerDashboard.Text) ||
                ComboBoxPickupCity_RequestAJob_CustomerDashboard.SelectedIndex == -1 ||
                string.IsNullOrWhiteSpace(txtDeliveryAddress_RequestAJob_CustomerDashboard.Text) ||
                ComboBoxDeliveryCity_RequestAJob_CustomerDashboard.SelectedIndex == -1 ||
                ComboBoxLocationType_RequestAJob_CustomerDashboard.SelectedIndex == -1 ||
                string.IsNullOrWhiteSpace(ComboBoxFirstProduct_RequestAJob_CustomerDashboard.Text) ||
                string.IsNullOrWhiteSpace(txtFirstProductMaxAmount_RequestAJob_CustomerDashboard.Text) ||
                string.IsNullOrWhiteSpace(ComboBoxSecondProduct_RequestAJob_CustomerDashboard.Text) ||
                string.IsNullOrWhiteSpace(txtSecondProductMaxAmount_RequestAJob_CustomerDashboard.Text))
            {
                MessageBox.Show("Please fill out all required fields before submitting the job request.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                con.Open();
                // Get CustomerID from username
                int customerId = -1;
                using (SqlCommand getIdCmd = new SqlCommand("SELECT CustomerID FROM Customers WHERE Username=@username", con))
                {
                    getIdCmd.Parameters.AddWithValue("@username", _username);
                    object idResult = getIdCmd.ExecuteScalar();
                    if (idResult != null && idResult != DBNull.Value)
                        customerId = Convert.ToInt32(idResult);
                    else
                        throw new Exception("Customer not found.");
                }

                if (dataGridViewMainGrid_MyDashboard_CustomerDashboard.SelectedRows.Count > 0)
                {
                    var selectedRow = dataGridViewMainGrid_MyDashboard_CustomerDashboard.SelectedRows[0];
                    var statusCell = selectedRow.Cells["Status"];
                    if (statusCell != null && statusCell.Value != null && statusCell.Value.ToString() == "Pending")
                    {
                        int jobIdToDelete = Convert.ToInt32(selectedRow.Cells["JobID"].Value);
                        using (SqlCommand deleteCmd = new SqlCommand("DELETE FROM Jobs WHERE JobID = @JobID AND CustomerID = @CustomerID", con))
                        {
                            deleteCmd.Parameters.AddWithValue("@JobID", jobIdToDelete);
                            deleteCmd.Parameters.AddWithValue("@CustomerID", customerId);
                            deleteCmd.ExecuteNonQuery();
                        }
                    }
                }

                // Collect values from controls
                string firstName = txtFname_RequestAJob_CustomerDashboard.Text.Trim();
                string lastName = txtLName_RequestAJob_CustomerDashboard.Text.Trim();
                string phoneNumber = txtPNumber_RequestAJob_CustomerDashboard.Text.Trim();
                string email = txtEmail_RequestAJob_CustomerDashboard.Text.Trim();
                string pickupAddress = txtPuckupAddress_RequestAJob_CustomerDashboard.Text.Trim();
                string pickupCity = ComboBoxPickupCity_RequestAJob_CustomerDashboard.SelectedItem?.ToString() ?? "";
                string deliveryAddress = txtDeliveryAddress_RequestAJob_CustomerDashboard.Text.Trim();
                string deliveryCity = ComboBoxDeliveryCity_RequestAJob_CustomerDashboard.SelectedItem?.ToString() ?? "";
                string locationType = ComboBoxLocationType_RequestAJob_CustomerDashboard.SelectedItem?.ToString() ?? "";
                DateTime pickupDate = dateTimePickerPickUpDate_RequestAJob_CustomerDashboard.Value;
                string firstProduct = ComboBoxFirstProduct_RequestAJob_CustomerDashboard.Text.Trim();
                int firstProductQuantity = int.TryParse(txtFirstProductMaxAmount_RequestAJob_CustomerDashboard.Text.Trim(), out int fpq) ? fpq : 0;
                string secondProduct = ComboBoxSecondProduct_RequestAJob_CustomerDashboard.Text.Trim();
                int secondProductQuantity = int.TryParse(txtSecondProductMaxAmount_RequestAJob_CustomerDashboard.Text.Trim(), out int spq) ? spq : 0;

                // Set default values for columns not yet assigned
                decimal totalDistance = 0;
                string estimatedTime = "Not Yet assigned";
                decimal jobCostPerDistance = 0;
                decimal totalCost = 0;

                string insertQuery = @"INSERT INTO Jobs (
                    CustomerID, LorryID, DriverID, AssistantID, ContainerID,
                    FirstProduct, FirstProductQuantity, SecondProduct, SecondProductQuantity,
                    FirstName, LastName, PhoneNumber, Email, PickupAddress, PickupCity,
                    DeliveryAddress, DeliveryCity, LocationType, PickupDate,
                    TotalDistance, EstimatedTime, JobCostPerDistance, TotalCost,
                    Status, CreatedDate, LoadStatus
                ) VALUES (
                    @CustomerID, @LorryID, @DriverID, @AssistantID, @ContainerID,
                    @FirstProduct, @FirstProductQuantity, @SecondProduct, @SecondProductQuantity,
                    @FirstName, @LastName, @PhoneNumber, @Email, @PickupAddress, @PickupCity,
                    @DeliveryAddress, @DeliveryCity, @LocationType, @PickupDate,
                    @TotalDistance, @EstimatedTime, @JobCostPerDistance, @TotalCost,
                    @Status, GETDATE(), @LoadStatus
                )";
                using (SqlCommand insertCmd = new SqlCommand(insertQuery, con))
                {
                    insertCmd.Parameters.AddWithValue("@CustomerID", customerId);
                    insertCmd.Parameters.AddWithValue("@LorryID", 0);      // Default, to be assigned by admin
                    insertCmd.Parameters.AddWithValue("@DriverID", 0);     // Default, to be assigned by admin
                    insertCmd.Parameters.AddWithValue("@AssistantID", 0);  // Default, to be assigned by admin
                    insertCmd.Parameters.AddWithValue("@ContainerID", 0);  // Default, to be assigned by admin
                    insertCmd.Parameters.AddWithValue("@FirstProduct", firstProduct);
                    insertCmd.Parameters.AddWithValue("@FirstProductQuantity", firstProductQuantity);
                    insertCmd.Parameters.AddWithValue("@SecondProduct", secondProduct);
                    insertCmd.Parameters.AddWithValue("@SecondProductQuantity", secondProductQuantity);
                    insertCmd.Parameters.AddWithValue("@FirstName", firstName);
                    insertCmd.Parameters.AddWithValue("@LastName", lastName);
                    insertCmd.Parameters.AddWithValue("@PhoneNumber", phoneNumber);
                    insertCmd.Parameters.AddWithValue("@Email", email);
                    insertCmd.Parameters.AddWithValue("@PickupAddress", pickupAddress);
                    insertCmd.Parameters.AddWithValue("@PickupCity", pickupCity);
                    insertCmd.Parameters.AddWithValue("@DeliveryAddress", deliveryAddress);
                    insertCmd.Parameters.AddWithValue("@DeliveryCity", deliveryCity);
                    insertCmd.Parameters.AddWithValue("@LocationType", locationType);
                    insertCmd.Parameters.AddWithValue("@PickupDate", pickupDate);
                    insertCmd.Parameters.AddWithValue("@TotalDistance", totalDistance);
                    insertCmd.Parameters.AddWithValue("@EstimatedTime", estimatedTime);
                    insertCmd.Parameters.AddWithValue("@JobCostPerDistance", jobCostPerDistance);
                    insertCmd.Parameters.AddWithValue("@TotalCost", totalCost);
                    insertCmd.Parameters.AddWithValue("@Status", "Pending");
                    insertCmd.Parameters.AddWithValue("@LoadStatus", "Not PickedUp");
                    int rows = insertCmd.ExecuteNonQuery();
                    if (rows > 0)
                    {
                        // Get the JobID of the newly inserted job
                        int newJobId = -1;
                        using (SqlCommand getJobIdCmd = new SqlCommand("SELECT TOP 1 JobID FROM Jobs WHERE CustomerID = @CustomerID ORDER BY CreatedDate DESC", con))
                        {
                            getJobIdCmd.Parameters.AddWithValue("@CustomerID", customerId);
                            object jobIdResult = getJobIdCmd.ExecuteScalar();
                            if (jobIdResult != null && jobIdResult != DBNull.Value)
                                newJobId = Convert.ToInt32(jobIdResult);
                        }

                        // Send confirmation email to the customer
                        try
                        {
                            string senderEmail = "";
                            // Get sender (admin) email
                            using (SqlCommand getAdminEmailCmd = new SqlCommand("SELECT TOP 1 Email FROM Admins WHERE Email IS NOT NULL AND Email <> ''", con))
                            {
                                using (SqlDataReader reader = getAdminEmailCmd.ExecuteReader())
                                {
                                    if (reader.Read())
                                    {
                                        senderEmail = reader["Email"].ToString();
                                    }
                                }
                            }
                            if (!string.IsNullOrEmpty(email) && !string.IsNullOrEmpty(senderEmail))
                            {
                                System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
                                mail.From = new System.Net.Mail.MailAddress(senderEmail);
                                mail.To.Add(email);
                                mail.Subject = "E-Shift - Job Request Submitted";
                                mail.Body = $"Dear {firstName},\n\nYour job request (Job ID: {newJobId}) was sent to E-Shift Company. You will receive an email on further updates.\n\nThank you for using E-Shift!\n\nBest regards,\nE-Shift Team";

                                System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("smtp.gmail.com", 587);
                                smtp.EnableSsl = true;
                                smtp.Credentials = new System.Net.NetworkCredential(senderEmail, Environment.GetEnvironmentVariable("ESHIFT_SMTP_APP_PASSWORD"));
                                smtp.Send(mail);
                            }
                        }
                        catch (Exception emailEx)
                        {
                            MessageBox.Show("Job request submitted, but failed to send confirmation email: " + emailEx.Message, "Email Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }

                        MessageBox.Show("Job request submitted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        txtFname_RequestAJob_CustomerDashboard.Text = string.Empty;
                        txtLName_RequestAJob_CustomerDashboard.Text = string.Empty;
                        txtPNumber_RequestAJob_CustomerDashboard.Text = string.Empty;
                        txtEmail_RequestAJob_CustomerDashboard.Text = string.Empty;
                        txtPuckupAddress_RequestAJob_CustomerDashboard.Text = string.Empty;
                        ComboBoxPickupCity_RequestAJob_CustomerDashboard.SelectedIndex = -1;
                        txtDeliveryAddress_RequestAJob_CustomerDashboard.Text = string.Empty;
                        ComboBoxDeliveryCity_RequestAJob_CustomerDashboard.SelectedIndex = -1;
                        ComboBoxLocationType_RequestAJob_CustomerDashboard.SelectedIndex = -1;
                        dateTimePickerPickUpDate_RequestAJob_CustomerDashboard.Value = DateTime.Now;
                        ComboBoxFirstProduct_RequestAJob_CustomerDashboard.SelectedIndex = -1;
                        txtFirstProductMaxAmount_RequestAJob_CustomerDashboard.Text = string.Empty;
                        ComboBoxSecondProduct_RequestAJob_CustomerDashboard.SelectedIndex = -1;
                        txtSecondProductMaxAmount_RequestAJob_CustomerDashboard.Text = string.Empty;
                        btnDelete_RequestAJob_CustomerDashboard.Enabled = true;
                        btnRequest_RequestAJob_CustomerDashboard.Enabled = true;
                        int refreshedCustomerId = customerId;
                        if (refreshedCustomerId != -1)
                        {
                            SqlCommand jobsCmd = new SqlCommand(@"
                                SELECT JobID, CustomerID, FirstProduct, FirstProductQuantity, SecondProduct, SecondProductQuantity, FirstName, LastName, PhoneNumber, Email, PickupAddress, PickupCity, DeliveryAddress, DeliveryCity, LocationType, PickupDate, TotalDistance, EstimatedTime, JobCostPerDistance, TotalCost, Status, CreatedDate, LoadStatus
                                FROM Jobs
                                WHERE CustomerID = @CustomerID
                                ORDER BY CreatedDate DESC", con);
                            jobsCmd.Parameters.AddWithValue("@CustomerID", refreshedCustomerId);
                            SqlDataAdapter da = new SqlDataAdapter(jobsCmd);
                            DataTable dt = new DataTable();
                            da.Fill(dt);
                            dataGridViewMainGrid_MyDashboard_CustomerDashboard.DataSource = dt;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Failed to submit job request.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error submitting job request: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
            }
        }

        private void btnClear_RequestAJob_CustomerDashboard_Click(object sender, EventArgs e)
        {
            txtFname_RequestAJob_CustomerDashboard.Text = string.Empty;
            txtLName_RequestAJob_CustomerDashboard.Text = string.Empty;
            txtPNumber_RequestAJob_CustomerDashboard.Text = string.Empty;
            txtEmail_RequestAJob_CustomerDashboard.Text = string.Empty;
            txtPuckupAddress_RequestAJob_CustomerDashboard.Text = string.Empty;
            ComboBoxPickupCity_RequestAJob_CustomerDashboard.SelectedIndex = -1;
            txtDeliveryAddress_RequestAJob_CustomerDashboard.Text = string.Empty;
            ComboBoxDeliveryCity_RequestAJob_CustomerDashboard.SelectedIndex = -1;
            ComboBoxLocationType_RequestAJob_CustomerDashboard.SelectedIndex = -1;
            dateTimePickerPickUpDate_RequestAJob_CustomerDashboard.Value = DateTime.Now;
            ComboBoxFirstProduct_RequestAJob_CustomerDashboard.SelectedIndex = -1;
            txtFirstProductMaxAmount_RequestAJob_CustomerDashboard.Text = string.Empty;
            ComboBoxSecondProduct_RequestAJob_CustomerDashboard.SelectedIndex = -1;
            txtSecondProductMaxAmount_RequestAJob_CustomerDashboard.Text = string.Empty;
            btnDelete_RequestAJob_CustomerDashboard.Enabled = true;
            btnRequest_RequestAJob_CustomerDashboard.Enabled = true;
        }

        private void txtEmail_RequestAJob_CustomerDashboard_TextChanged(object sender, EventArgs e)
        {
        }

        private void ImageButtonSearch_MyDashboard_CustomerDashboard_Click(object sender, EventArgs e)
        {
            string jobIdText = txtSearchByJobID_MyDashboard_CustomerDashboard.Text.Trim();
            if (string.IsNullOrWhiteSpace(jobIdText))
            {
                MessageBox.Show("Please enter a Job ID to search.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!int.TryParse(jobIdText, out int jobId))
            {
                MessageBox.Show("Please enter a valid numeric Job ID.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                con.Open();
                int customerId = -1;
                using (SqlCommand getIdCmd = new SqlCommand("SELECT CustomerID FROM Customers WHERE Username=@username", con))
                {
                    getIdCmd.Parameters.AddWithValue("@username", _username);
                    object idResult = getIdCmd.ExecuteScalar();
                    if (idResult != null && idResult != DBNull.Value)
                        customerId = Convert.ToInt32(idResult);
                    else
                        throw new Exception("Customer not found.");
                }

                SqlCommand searchCmd = new SqlCommand(@"
                    SELECT JobID, CustomerID, FirstProduct, FirstProductQuantity, SecondProduct, SecondProductQuantity, FirstName, LastName, PhoneNumber, Email, PickupAddress, PickupCity, DeliveryAddress, DeliveryCity, LocationType, PickupDate, TotalDistance, EstimatedTime, JobCostPerDistance, TotalCost, Status, CreatedDate, LoadStatus
                    FROM Jobs
                    WHERE CustomerID = @CustomerID AND JobID = @JobID", con);
                searchCmd.Parameters.AddWithValue("@CustomerID", customerId);
                searchCmd.Parameters.AddWithValue("@JobID", jobId);
                SqlDataAdapter da = new SqlDataAdapter(searchCmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridViewMainGrid_MyDashboard_CustomerDashboard.DataSource = dt;

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
        }

        private void ImageButtonRefresh_MyDashboard_CustomerDashboard_Click(object sender, EventArgs e)
        {
            LoadDashboardData();
        }

        private void RadioButtonAll_MyDashboard_CustomerDashboard_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonAll_MyDashboard_CustomerDashboard.Checked)
            {
                LoadDashboardData();
            }
        }

        private void RadioButtonNotPickedUp_MyDashboard_CustomerDashboard_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonNotPickedUp_MyDashboard_CustomerDashboard.Checked)
            {
                try
                {
                    con.Open();
                    int customerId = -1;
                    using (SqlCommand getIdCmd = new SqlCommand("SELECT CustomerID FROM Customers WHERE Username=@username", con))
                    {
                        getIdCmd.Parameters.AddWithValue("@username", _username);
                        object idResult = getIdCmd.ExecuteScalar();
                        if (idResult != null && idResult != DBNull.Value)
                            customerId = Convert.ToInt32(idResult);
                    }
                    if (customerId != -1)
                    {
                        SqlCommand jobsCmd = new SqlCommand(@"
                            SELECT JobID, CustomerID, FirstProduct, FirstProductQuantity, SecondProduct, SecondProductQuantity, FirstName, LastName, PhoneNumber, Email, PickupAddress, PickupCity, DeliveryAddress, DeliveryCity, LocationType, PickupDate, TotalDistance, EstimatedTime, JobCostPerDistance, TotalCost, Status, CreatedDate, LoadStatus
                            FROM Jobs
                            WHERE CustomerID = @CustomerID AND LoadStatus = @LoadStatus
                            ORDER BY CreatedDate DESC", con);
                        jobsCmd.Parameters.AddWithValue("@CustomerID", customerId);
                        jobsCmd.Parameters.AddWithValue("@LoadStatus", "Not PickedUp");
                        SqlDataAdapter da = new SqlDataAdapter(jobsCmd);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dataGridViewMainGrid_MyDashboard_CustomerDashboard.DataSource = dt;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to filter jobs: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    con.Close();
                }
            }
        }

        private void RadioButtonPickedUp_MyDashboard_CustomerDashboard_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonPickedUp_MyDashboard_CustomerDashboard.Checked)
            {
                try
                {
                    con.Open();
                    int customerId = -1;
                    using (SqlCommand getIdCmd = new SqlCommand("SELECT CustomerID FROM Customers WHERE Username=@username", con))
                    {
                        getIdCmd.Parameters.AddWithValue("@username", _username);
                        object idResult = getIdCmd.ExecuteScalar();
                        if (idResult != null && idResult != DBNull.Value)
                            customerId = Convert.ToInt32(idResult);
                    }
                    if (customerId != -1)
                    {
                        SqlCommand jobsCmd = new SqlCommand(@"
                            SELECT JobID, CustomerID, FirstProduct, FirstProductQuantity, SecondProduct, SecondProductQuantity, FirstName, LastName, PhoneNumber, Email, PickupAddress, PickupCity, DeliveryAddress, DeliveryCity, LocationType, PickupDate, TotalDistance, EstimatedTime, JobCostPerDistance, TotalCost, Status, CreatedDate, LoadStatus
                            FROM Jobs
                            WHERE CustomerID = @CustomerID AND LoadStatus = @LoadStatus
                            ORDER BY CreatedDate DESC", con);
                        jobsCmd.Parameters.AddWithValue("@CustomerID", customerId);
                        jobsCmd.Parameters.AddWithValue("@LoadStatus", "PickedUp");
                        SqlDataAdapter da = new SqlDataAdapter(jobsCmd);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dataGridViewMainGrid_MyDashboard_CustomerDashboard.DataSource = dt;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to filter jobs: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    con.Close();
                }
            }
        }

        private void RadioButtonDelivered_MyDashboard_CustomerDashboard_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonDelivered_MyDashboard_CustomerDashboard.Checked)
            {
                try
                {
                    con.Open();
                    int customerId = -1;
                    using (SqlCommand getIdCmd = new SqlCommand("SELECT CustomerID FROM Customers WHERE Username=@username", con))
                    {
                        getIdCmd.Parameters.AddWithValue("@username", _username);
                        object idResult = getIdCmd.ExecuteScalar();
                        if (idResult != null && idResult != DBNull.Value)
                            customerId = Convert.ToInt32(idResult);
                    }
                    if (customerId != -1)
                    {
                        SqlCommand jobsCmd = new SqlCommand(@"
                            SELECT JobID, CustomerID, FirstProduct, FirstProductQuantity, SecondProduct, SecondProductQuantity, FirstName, LastName, PhoneNumber, Email, PickupAddress, PickupCity, DeliveryAddress, DeliveryCity, LocationType, PickupDate, TotalDistance, EstimatedTime, JobCostPerDistance, TotalCost, Status, CreatedDate, LoadStatus
                            FROM Jobs
                            WHERE CustomerID = @CustomerID AND LoadStatus = @LoadStatus
                            ORDER BY CreatedDate DESC", con);
                        jobsCmd.Parameters.AddWithValue("@CustomerID", customerId);
                        jobsCmd.Parameters.AddWithValue("@LoadStatus", "Delivered");
                        SqlDataAdapter da = new SqlDataAdapter(jobsCmd);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dataGridViewMainGrid_MyDashboard_CustomerDashboard.DataSource = dt;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to filter jobs: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    con.Close();
                }
            }
        }

        private void RadioButtonAccepted_MyDashboard_CustomerDashboard_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonAccepted_MyDashboard_CustomerDashboard.Checked)
            {
                try
                {
                    con.Open();
                    int customerId = -1;
                    using (SqlCommand getIdCmd = new SqlCommand("SELECT CustomerID FROM Customers WHERE Username=@username", con))
                    {
                        getIdCmd.Parameters.AddWithValue("@username", _username);
                        object idResult = getIdCmd.ExecuteScalar();
                        if (idResult != null && idResult != DBNull.Value)
                            customerId = Convert.ToInt32(idResult);
                    }
                    if (customerId != -1)
                    {
                        SqlCommand jobsCmd = new SqlCommand(@"
                            SELECT JobID, CustomerID, FirstProduct, FirstProductQuantity, SecondProduct, SecondProductQuantity, FirstName, LastName, PhoneNumber, Email, PickupAddress, PickupCity, DeliveryAddress, DeliveryCity, LocationType, PickupDate, TotalDistance, EstimatedTime, JobCostPerDistance, TotalCost, Status, CreatedDate, LoadStatus
                            FROM Jobs
                            WHERE CustomerID = @CustomerID AND Status = @Status
                            ORDER BY CreatedDate DESC", con);
                        jobsCmd.Parameters.AddWithValue("@CustomerID", customerId);
                        jobsCmd.Parameters.AddWithValue("@Status", "Accepted");
                        SqlDataAdapter da = new SqlDataAdapter(jobsCmd);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dataGridViewMainGrid_MyDashboard_CustomerDashboard.DataSource = dt;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to filter jobs: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    con.Close();
                }
            }
        }

        private void RadioButtonRejected_MyDashboard_CustomerDashboard_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonRejected_MyDashboard_CustomerDashboard.Checked)
            {
                try
                {
                    con.Open();
                    int customerId = -1;
                    using (SqlCommand getIdCmd = new SqlCommand("SELECT CustomerID FROM Customers WHERE Username=@username", con))
                    {
                        getIdCmd.Parameters.AddWithValue("@username", _username);
                        object idResult = getIdCmd.ExecuteScalar();
                        if (idResult != null && idResult != DBNull.Value)
                            customerId = Convert.ToInt32(idResult);
                    }
                    if (customerId != -1)
                    {
                        SqlCommand jobsCmd = new SqlCommand(@"
                            SELECT JobID, CustomerID, FirstProduct, FirstProductQuantity, SecondProduct, SecondProductQuantity, FirstName, LastName, PhoneNumber, Email, PickupAddress, PickupCity, DeliveryAddress, DeliveryCity, LocationType, PickupDate, TotalDistance, EstimatedTime, JobCostPerDistance, TotalCost, Status, CreatedDate, LoadStatus
                            FROM Jobs
                            WHERE CustomerID = @CustomerID AND Status = @Status
                            ORDER BY CreatedDate DESC", con);
                        jobsCmd.Parameters.AddWithValue("@CustomerID", customerId);
                        jobsCmd.Parameters.AddWithValue("@Status", "Rejected");
                        SqlDataAdapter da = new SqlDataAdapter(jobsCmd);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dataGridViewMainGrid_MyDashboard_CustomerDashboard.DataSource = dt;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to filter jobs: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    con.Close();
                }
            }
        }

        private void RadioButtonPendingRequests_MyDashboard_CustomerDashboard_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonPendingRequests_MyDashboard_CustomerDashboard.Checked)
            {
                try
                {
                    con.Open();
                    int customerId = -1;
                    using (SqlCommand getIdCmd = new SqlCommand("SELECT CustomerID FROM Customers WHERE Username=@username", con))
                    {
                        getIdCmd.Parameters.AddWithValue("@username", _username);
                        object idResult = getIdCmd.ExecuteScalar();
                        if (idResult != null && idResult != DBNull.Value)
                            customerId = Convert.ToInt32(idResult);
                    }
                    if (customerId != -1)
                    {
                        SqlCommand jobsCmd = new SqlCommand(@"
                            SELECT JobID, CustomerID, FirstProduct, FirstProductQuantity, SecondProduct, SecondProductQuantity, FirstName, LastName, PhoneNumber, Email, PickupAddress, PickupCity, DeliveryAddress, DeliveryCity, LocationType, PickupDate, TotalDistance, EstimatedTime, JobCostPerDistance, TotalCost, Status, CreatedDate, LoadStatus
                            FROM Jobs
                            WHERE CustomerID = @CustomerID AND Status = @Status
                            ORDER BY CreatedDate DESC", con);
                        jobsCmd.Parameters.AddWithValue("@CustomerID", customerId);
                        jobsCmd.Parameters.AddWithValue("@Status", "Pending");
                        SqlDataAdapter da = new SqlDataAdapter(jobsCmd);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dataGridViewMainGrid_MyDashboard_CustomerDashboard.DataSource = dt;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to filter jobs: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    con.Close();
                }
            }
        }

        private void btnLogout_MyDashboard_CustomerDashboard_Click(object sender, EventArgs e)
        {
            Role_Selection signUpForm = new Role_Selection();
            signUpForm.Show();
            this.Hide();
        }

        private void btnRequestJobAgain_MyDashboard_CustomerDashboard_Click(object sender, EventArgs e)
        // When user clicks Request Job Again, fill form and enable delete for that job
        {
            if (dataGridViewMainGrid_MyDashboard_CustomerDashboard.SelectedRows.Count > 0)
            {
                var row = dataGridViewMainGrid_MyDashboard_CustomerDashboard.SelectedRows[0];
                // Switch to the Request A Job tab
                tabPagRequestAJob_CustomerDashboard.Select();

                // Fill controls with selected job data
                txtFname_RequestAJob_CustomerDashboard.Text = row.Cells["FirstName"].Value?.ToString() ?? "";
                txtLName_RequestAJob_CustomerDashboard.Text = row.Cells["LastName"].Value?.ToString() ?? "";
                txtPNumber_RequestAJob_CustomerDashboard.Text = row.Cells["PhoneNumber"].Value?.ToString() ?? "";
                txtEmail_RequestAJob_CustomerDashboard.Text = row.Cells["Email"].Value?.ToString() ?? "";
                txtPuckupAddress_RequestAJob_CustomerDashboard.Text = row.Cells["PickupAddress"].Value?.ToString() ?? "";
                ComboBoxPickupCity_RequestAJob_CustomerDashboard.SelectedItem = row.Cells["PickupCity"].Value?.ToString() ?? null;
                txtDeliveryAddress_RequestAJob_CustomerDashboard.Text = row.Cells["DeliveryAddress"].Value?.ToString() ?? "";
                ComboBoxDeliveryCity_RequestAJob_CustomerDashboard.SelectedItem = row.Cells["DeliveryCity"].Value?.ToString() ?? null;
                ComboBoxLocationType_RequestAJob_CustomerDashboard.SelectedItem = row.Cells["LocationType"].Value?.ToString() ?? null;
                if (row.Cells["PickupDate"].Value != null && DateTime.TryParse(row.Cells["PickupDate"].Value.ToString(), out DateTime pickupDate))
                    dateTimePickerPickUpDate_RequestAJob_CustomerDashboard.Value = pickupDate;
                else
                    dateTimePickerPickUpDate_RequestAJob_CustomerDashboard.Value = DateTime.Now;
                ComboBoxFirstProduct_RequestAJob_CustomerDashboard.SelectedItem = row.Cells["FirstProduct"].Value?.ToString() ?? null;
                txtFirstProductMaxAmount_RequestAJob_CustomerDashboard.Text = row.Cells["FirstProductQuantity"].Value?.ToString() ?? "";
                ComboBoxSecondProduct_RequestAJob_CustomerDashboard.SelectedItem = row.Cells["SecondProduct"].Value?.ToString() ?? null;
                txtSecondProductMaxAmount_RequestAJob_CustomerDashboard.Text = row.Cells["SecondProductQuantity"].Value?.ToString() ?? "";

                // Enable delete button and store JobID for deletion
                if (row.Cells["JobID"].Value != null && int.TryParse(row.Cells["JobID"].Value.ToString(), out int jobId))
                {
                    _pendingJobIdForDelete = jobId;
                    btnDelete_RequestAJob_CustomerDashboard.Enabled = true;
                }
                else
                {
                    _pendingJobIdForDelete = null;
                    btnDelete_RequestAJob_CustomerDashboard.Enabled = true;
                }
            }
        }

        // Delete the pending job if set
        private void btnDelete_RequestAJob_CustomerDashboard_Click(object sender, EventArgs e)
        {
            if (_pendingJobIdForDelete.HasValue)
            {
                try
                {
                    con.Open();
                    // Get CustomerID from username
                    int customerId = -1;
                    using (SqlCommand getIdCmd = new SqlCommand("SELECT CustomerID FROM Customers WHERE Username=@username", con))
                    {
                        getIdCmd.Parameters.AddWithValue("@username", _username);
                        object idResult = getIdCmd.ExecuteScalar();
                        if (idResult != null && idResult != DBNull.Value)
                            customerId = Convert.ToInt32(idResult);
                        else
                            throw new Exception("Customer not found.");
                    }
                    using (SqlCommand deleteCmd = new SqlCommand("DELETE FROM Jobs WHERE JobID = @JobID AND CustomerID = @CustomerID", con))
                    {
                        deleteCmd.Parameters.AddWithValue("@JobID", _pendingJobIdForDelete.Value);
                        deleteCmd.Parameters.AddWithValue("@CustomerID", customerId);
                        int rows = deleteCmd.ExecuteNonQuery();
                        if (rows > 0)
                        {
                            MessageBox.Show("Job deleted successfully.", "Deleted", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            LoadDashboardData();
                            // After deletion, clear the form and disable delete button
                            btnClear_RequestAJob_CustomerDashboard_Click(null, null);
                        }
                        else
                        {
                            MessageBox.Show("Failed to delete job.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error deleting job: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    con.Close();
                    btnDelete_RequestAJob_CustomerDashboard.Enabled = false;
                    _pendingJobIdForDelete = null;
                }
            }
        }

        private void btnDelete_RequestAJob_CustomerDashboard_Click_1(object sender, EventArgs e)
        {
            if (_pendingJobIdForDelete.HasValue)
            {
                try
                {
                    con.Open();
                    int customerId = -1;
                    using (SqlCommand getIdCmd = new SqlCommand("SELECT CustomerID FROM Customers WHERE Username=@username", con))
                    {
                        getIdCmd.Parameters.AddWithValue("@username", _username);
                        object idResult = getIdCmd.ExecuteScalar();
                        if (idResult != null && idResult != DBNull.Value)
                            customerId = Convert.ToInt32(idResult);
                        else
                            throw new Exception("Customer not found.");
                    }
                    using (SqlCommand deleteCmd = new SqlCommand("DELETE FROM Jobs WHERE JobID = @JobID AND CustomerID = @CustomerID", con))
                    {
                        deleteCmd.Parameters.AddWithValue("@JobID", _pendingJobIdForDelete.Value);
                        deleteCmd.Parameters.AddWithValue("@CustomerID", customerId);
                        int rows = deleteCmd.ExecuteNonQuery();
                        if (rows > 0)
                        {
                            MessageBox.Show("Job deleted successfully.", "Deleted", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            LoadDashboardData();
                            btnClear_RequestAJob_CustomerDashboard_Click(null, null);
                        }
                        else
                        {
                            MessageBox.Show("Failed to delete job.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error deleting job: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    con.Close();
                    btnDelete_RequestAJob_CustomerDashboard.Enabled = false;
                    _pendingJobIdForDelete = null;
                }
            }
        }

        private void btnPrintInvoice_Click(object sender, EventArgs e)
        {
            // Export DataGridView to PDF using SaveFileDialog and iTextSharp
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "PDF files (*.pdf)|*.pdf";
                sfd.Title = "Save Invoice as PDF";
                sfd.FileName = "Invoice.pdf";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        // Create PDF document in landscape orientation with margins
                        Document doc = new Document(PageSize.A4.Rotate(), 40f, 40f, 60f, 50f);
                        PdfWriter writer = PdfWriter.GetInstance(doc, new System.IO.FileStream(sfd.FileName, System.IO.FileMode.Create));
                        doc.Open();

                        var titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 18);
                        var subTitleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12);
                        var bodyFont = FontFactory.GetFont(FontFactory.HELVETICA, 9);
                        var boldTableFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8);
                        var endingFont = FontFactory.GetFont(FontFactory.HELVETICA_OBLIQUE, 8);

                        PdfPTable headerTable = new PdfPTable(2);
                        headerTable.WidthPercentage = 100;
                        headerTable.SetWidths(new float[] { 20, 80 });

                        if (pictureBoxESHIFTLogo_MyDashboard_CustomerDashboard != null && pictureBoxESHIFTLogo_MyDashboard_CustomerDashboard.Image != null)
                        {
                            using (var ms = new System.IO.MemoryStream())
                            {
                                pictureBoxESHIFTLogo_MyDashboard_CustomerDashboard.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                ms.Position = 0;
                                iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(ms);
                                logo.ScaleAbsolute(100f, 50f); // Scale image

                                PdfPCell logoCell = new PdfPCell(logo);
                                logoCell.Border = global::iTextSharp.text.pdf.PdfPCell.NO_BORDER;
                                logoCell.HorizontalAlignment = Element.ALIGN_LEFT;
                                logoCell.VerticalAlignment = Element.ALIGN_TOP;
                                headerTable.AddCell(logoCell);
                            }
                        }
                        else
                        {
                            PdfPCell emptyLogoCell = new PdfPCell(new Phrase(" "));
                            emptyLogoCell.Border = global::iTextSharp.text.pdf.PdfPCell.NO_BORDER;
                            headerTable.AddCell(emptyLogoCell);
                        }

                        PdfPCell companyInfoCell = new PdfPCell();
                        companyInfoCell.Border = global::iTextSharp.text.pdf.PdfPCell.NO_BORDER;
                        companyInfoCell.AddElement(new Paragraph("E-Shift Company", subTitleFont));
                        companyInfoCell.AddElement(new Paragraph("123 Business St,\nColombo, Sri Lanka\nPhone: +94 712345678", bodyFont));
                        companyInfoCell.AddElement(new Paragraph("Invoice Date: " + DateTime.Now.ToString("yyyy-MM-dd"), bodyFont));
                        companyInfoCell.AddElement(new Paragraph("Invoice #: INV-" + DateTime.Now.ToString("yyyyMMddHHmmss"), bodyFont));
                        headerTable.AddCell(companyInfoCell);

                        doc.Add(headerTable);
                        doc.Add(new Paragraph("\n", bodyFont)); 

                        PdfPTable pdfTable = new PdfPTable(dataGridViewMainGrid_MyDashboard_CustomerDashboard.Columns.Count);
                        pdfTable.WidthPercentage = 100;
                        pdfTable.SpacingBefore = 10f;
                        pdfTable.SpacingAfter = 10f;

                        float[] columnWidths = new float[dataGridViewMainGrid_MyDashboard_CustomerDashboard.Columns.Count];
                        for (int i = 0; i < dataGridViewMainGrid_MyDashboard_CustomerDashboard.Columns.Count; i++)
                            columnWidths[i] = 4f;
                        pdfTable.SetWidths(columnWidths);

                        foreach (DataGridViewColumn column in dataGridViewMainGrid_MyDashboard_CustomerDashboard.Columns)
                        {
                            PdfPCell headerCell = new PdfPCell(new Phrase(column.HeaderText, boldTableFont));
                            headerCell.BackgroundColor = BaseColor.LIGHT_GRAY;
                            headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
                            headerCell.Padding = 5;
                            pdfTable.AddCell(headerCell);
                        }

                        foreach (DataGridViewRow row in dataGridViewMainGrid_MyDashboard_CustomerDashboard.Rows)
                        {
                            if (!row.IsNewRow)
                            {
                                foreach (DataGridViewCell cell in row.Cells)
                                {
                                    PdfPCell dataCell = new PdfPCell(new Phrase(cell.Value?.ToString() ?? "", bodyFont));
                                    dataCell.Padding = 4;
                                    dataCell.HorizontalAlignment = Element.ALIGN_CENTER;
                                    pdfTable.AddCell(dataCell);
                                }
                            }
                        }

                        doc.Add(pdfTable);

                        doc.Add(new Paragraph("Terms and Conditions", boldTableFont));
                        doc.Add(new Paragraph("1. Payment is due within 7 days.\n2. Goods once sold will not be taken back.\n3. For support, contact our office.", bodyFont));
                        doc.Add(new Paragraph("\nThank you for your business!", endingFont));

                        doc.Close();
                        writer.Close();

                        MessageBox.Show("Invoice PDF exported successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + sfd.FileName + "\"");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failed to export PDF: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnUpdate_Settings_CustomerDashboard_Click(object sender, EventArgs e)
        {
            string oldPassword = txtOldPassword_Settings_CustomerDashboard.Text;
            string newPassword = txtNewPassword_Settings_CustomerDashboard.Text;
            string confirmPassword = txtConfirmPassword_Settings_CustomerDashboard.Text;
            bool changePassword = !string.IsNullOrWhiteSpace(oldPassword) ||
                                  !string.IsNullOrWhiteSpace(newPassword) ||
                                  !string.IsNullOrWhiteSpace(confirmPassword);

            if (changePassword)
            {
                if (string.IsNullOrWhiteSpace(oldPassword) || string.IsNullOrWhiteSpace(newPassword) || string.IsNullOrWhiteSpace(confirmPassword))
                {
                    MessageBox.Show("To change your password, fill all password fields.", "Missing Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (newPassword != confirmPassword)
                {
                    MessageBox.Show("New password and confirm password do not match.", "Mismatch", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            string GetSha256Hash(string input)
            {
                using (var sha256 = System.Security.Cryptography.SHA256.Create())
                {
                    byte[] bytes = Encoding.UTF8.GetBytes(input);
                    byte[] hash = sha256.ComputeHash(bytes);
                    return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                }
            }

            try
            {
                con.Open();

                int customerId = -1;
                string currentPasswordHash = null;

                using (SqlCommand getIdCmd = new SqlCommand("SELECT CustomerID, PasswordHash FROM Customers WHERE Username=@username", con))
                {
                    getIdCmd.Parameters.AddWithValue("@username", _username);
                    using (var reader = getIdCmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            customerId = reader.GetInt32(0);
                            currentPasswordHash = reader.GetString(1);
                        }
                        else
                        {
                            MessageBox.Show("Customer not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }

                string passwordHashToStore = currentPasswordHash;

                if (changePassword)
                {
                    string oldPasswordHash = GetSha256Hash(oldPassword);
                    if (oldPasswordHash != currentPasswordHash)
                    {
                        MessageBox.Show("Old password is incorrect.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    passwordHashToStore = GetSha256Hash(newPassword);
                }

                byte[] profilePhotoBytes = null;
                if (pictureBoxChangeProfilePhoto_Settings_CustomerDashboard.Image != null)
                {
                    using (var ms = new System.IO.MemoryStream())
                    {
                        pictureBoxChangeProfilePhoto_Settings_CustomerDashboard.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                        profilePhotoBytes = ms.ToArray();
                    }
                }

                var updateFields = new List<string>();
                var parameters = new List<SqlParameter>();

                void AddIfFilled(string field, string value)
                {
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        updateFields.Add($"{field} = @{field}");
                        parameters.Add(new SqlParameter($"@{field}", value.Trim()));
                    }
                }

                AddIfFilled("Title", ComboBoxTitle_Settings_CustomerDashboard.Text);
                AddIfFilled("FirstName", txtFName_Settings_CustomerDashboard.Text);
                AddIfFilled("LastName", txtLName_Settings_CustomerDashboard.Text);
                AddIfFilled("Address", txtAddress_Settings_CustomerDashboard.Text);
                AddIfFilled("PhoneNumber", txtPNumber_Settings_CustomerDashboard.Text);
                AddIfFilled("Email", txtNewEmail_Settings_CustomerDashboard.Text);
                AddIfFilled("NIC", txtNIC_Settings_CustomerDashboard.Text);
                AddIfFilled("Username", txtNewUsername_Settings_CustomerDashboard.Text);

                if (changePassword)
                {
                    updateFields.Add("PasswordHash = @PasswordHash");
                    parameters.Add(new SqlParameter("@PasswordHash", passwordHashToStore));
                }

                if (profilePhotoBytes != null)
                {
                    updateFields.Add("ProfilePhoto = @ProfilePhoto");
                    parameters.Add(new SqlParameter("@ProfilePhoto", profilePhotoBytes));
                }

                if (updateFields.Count == 0)
                {
                    MessageBox.Show("Please enter at least one field to update.", "No Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string updateQuery = $"UPDATE Customers SET {string.Join(", ", updateFields)} WHERE CustomerID = @CustomerID";
                using (SqlCommand updateCmd = new SqlCommand(updateQuery, con))
                {
                    foreach (var param in parameters)
                        updateCmd.Parameters.Add(param);

                    updateCmd.Parameters.AddWithValue("@CustomerID", customerId);

                    int rows = updateCmd.ExecuteNonQuery();
                    if (rows > 0)
                    {
                        MessageBox.Show("Profile updated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        if (!string.IsNullOrWhiteSpace(txtNewUsername_Settings_CustomerDashboard.Text))
                        {
                            _username = txtNewUsername_Settings_CustomerDashboard.Text.Trim();
                            lblCustomerUsername_CustomerDashboard.Text = _username;
                        }
                    }
                    else
                    {
                        MessageBox.Show("No changes were made.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to update profile: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
            }
        }

        private void pictureBoxChangeProfilePhoto_Settings_CustomerDashboard_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Select Profile Photo";
                ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var img = System.Drawing.Image.FromFile(ofd.FileName);
                        pictureBoxChangeProfilePhoto_Settings_CustomerDashboard.Image = img;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failed to load image: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnClear_Settings_CustomerDashboard_Click(object sender, EventArgs e)
        {
            ComboBoxTitle_Settings_CustomerDashboard.SelectedIndex = -1;
            txtFName_Settings_CustomerDashboard.Text = string.Empty;
            txtLName_Settings_CustomerDashboard.Text = string.Empty;
            txtNIC_Settings_CustomerDashboard.Text = string.Empty;
            txtPNumber_Settings_CustomerDashboard.Text = string.Empty;
            txtAddress_Settings_CustomerDashboard.Text = string.Empty;
            txtNewUsername_Settings_CustomerDashboard.Text = string.Empty;
            txtNewEmail_Settings_CustomerDashboard.Text = string.Empty;
            txtOldPassword_Settings_CustomerDashboard.Text = string.Empty;
            txtNewPassword_Settings_CustomerDashboard.Text = string.Empty;
            txtConfirmPassword_Settings_CustomerDashboard.Text = string.Empty;
        }

        private void checkBoxShowThePassword_Settings_CustomerDashboard_CheckedChanged(object sender, EventArgs e)
        {
            bool show = checkBoxShowThePassword_Settings_CustomerDashboard.Checked;
            if (show)
            {
                txtOldPassword_Settings_CustomerDashboard.UseSystemPasswordChar = false;
                txtOldPassword_Settings_CustomerDashboard.PasswordChar = '\0';
                txtNewPassword_Settings_CustomerDashboard.UseSystemPasswordChar = false;
                txtNewPassword_Settings_CustomerDashboard.PasswordChar = '\0';
                txtConfirmPassword_Settings_CustomerDashboard.UseSystemPasswordChar = false;
                txtConfirmPassword_Settings_CustomerDashboard.PasswordChar = '\0';
            }
            else
            {
                txtOldPassword_Settings_CustomerDashboard.UseSystemPasswordChar = true;
                txtOldPassword_Settings_CustomerDashboard.PasswordChar = '●';
                txtNewPassword_Settings_CustomerDashboard.UseSystemPasswordChar = true;
                txtNewPassword_Settings_CustomerDashboard.PasswordChar = '●';
                txtConfirmPassword_Settings_CustomerDashboard.UseSystemPasswordChar = true;
                txtConfirmPassword_Settings_CustomerDashboard.PasswordChar = '●';
            }
        }

        private void btnViewJob_MyDashboard_CustomerDashboard_Click(object sender, EventArgs e)
        {
            if (dataGridViewMainGrid_MyDashboard_CustomerDashboard.SelectedRows.Count > 0)
            {
                var row = dataGridViewMainGrid_MyDashboard_CustomerDashboard.SelectedRows[0];

                TabControlCustomerDashboard.SelectedTab = tabPagRequestAJob_CustomerDashboard;
                tabPagRequestAJob_CustomerDashboard.Select();

                txtFname_RequestAJob_CustomerDashboard.Text = row.Cells["FirstName"].Value?.ToString() ?? "";
                txtLName_RequestAJob_CustomerDashboard.Text = row.Cells["LastName"].Value?.ToString() ?? "";
                txtPNumber_RequestAJob_CustomerDashboard.Text = row.Cells["PhoneNumber"].Value?.ToString() ?? "";
                txtEmail_RequestAJob_CustomerDashboard.Text = row.Cells["Email"].Value?.ToString() ?? "";
                txtPuckupAddress_RequestAJob_CustomerDashboard.Text = row.Cells["PickupAddress"].Value?.ToString() ?? "";
                ComboBoxPickupCity_RequestAJob_CustomerDashboard.SelectedItem = row.Cells["PickupCity"].Value?.ToString() ?? null;
                txtDeliveryAddress_RequestAJob_CustomerDashboard.Text = row.Cells["DeliveryAddress"].Value?.ToString() ?? "";
                ComboBoxDeliveryCity_RequestAJob_CustomerDashboard.SelectedItem = row.Cells["DeliveryCity"].Value?.ToString() ?? null;
                ComboBoxLocationType_RequestAJob_CustomerDashboard.SelectedItem = row.Cells["LocationType"].Value?.ToString() ?? null;
                if (row.Cells["PickupDate"].Value != null && DateTime.TryParse(row.Cells["PickupDate"].Value.ToString(), out DateTime pickupDate))
                    dateTimePickerPickUpDate_RequestAJob_CustomerDashboard.Value = pickupDate;
                else
                    dateTimePickerPickUpDate_RequestAJob_CustomerDashboard.Value = DateTime.Now;
                ComboBoxFirstProduct_RequestAJob_CustomerDashboard.SelectedItem = row.Cells["FirstProduct"].Value?.ToString() ?? null;
                txtFirstProductMaxAmount_RequestAJob_CustomerDashboard.Text = row.Cells["FirstProductQuantity"].Value?.ToString() ?? "";
                ComboBoxSecondProduct_RequestAJob_CustomerDashboard.SelectedItem = row.Cells["SecondProduct"].Value?.ToString() ?? null;
                txtSecondProductMaxAmount_RequestAJob_CustomerDashboard.Text = row.Cells["SecondProductQuantity"].Value?.ToString() ?? "";
                btnDelete_RequestAJob_CustomerDashboard.Enabled = false;
                btnRequest_RequestAJob_CustomerDashboard.Enabled = false;

            }
        }

        private void btnLogout_Settings_CustomerDashboard_Click(object sender, EventArgs e)
        {
            Role_Selection signUpForm = new Role_Selection();
            signUpForm.Show();
            this.Hide();
        }

        private void dataGridViewMainGrid_MyDashboard_CustomerDashboard_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }

}
