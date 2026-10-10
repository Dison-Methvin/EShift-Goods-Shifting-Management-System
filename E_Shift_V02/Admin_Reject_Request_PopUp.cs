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

namespace E_Shift_V02
{
    public partial class Admin_Reject_Request_PopUp : Form
    {
        public Admin_Reject_Request_PopUp()
        {
            InitializeComponent();
        }

        SqlConnection con = new SqlConnection("Data Source=DISON_PC\\SQLEXPRESS;Initial Catalog=ESHIFT;Integrated Security=True;Encrypt=False");


        private void Admin_Reject_Request_PopUp_Load(object sender, EventArgs e)
        {

        }

        private void groupBoxRejecting_AdminRejectRequestPopUp_Enter(object sender, EventArgs e)
        {
        }

        public Admin_Reject_Request_PopUp(string jobId, string customerId)
        {
            InitializeComponent();
            ComboBoxJobID_AdminRejectRequestPopUp.Items.Clear();
            ComboBoxJobID_AdminRejectRequestPopUp.Items.Add(jobId);
            ComboBoxJobID_AdminRejectRequestPopUp.SelectedItem = jobId;

            ComboBoxCustomerID_AdminRejectRequestPopUp.Items.Clear();
            ComboBoxCustomerID_AdminRejectRequestPopUp.Items.Add(customerId);
            ComboBoxCustomerID_AdminRejectRequestPopUp.SelectedItem = customerId;
        }

        private void ComboBoxJobID_AdminRejectRequestPopUp_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void btnContinueRejecting_AdminRejectRequestPopUp_Click(object sender, EventArgs e)
        {
            string jobId = ComboBoxJobID_AdminRejectRequestPopUp.SelectedItem?.ToString();
            string customerId = ComboBoxCustomerID_AdminRejectRequestPopUp.SelectedItem?.ToString();
            string reason = txtReasonForRejecting_AdminRejectRequestPopUp.Text.Trim();
            if (string.IsNullOrEmpty(jobId) || string.IsNullOrEmpty(customerId) || string.IsNullOrEmpty(reason))
            {
                MessageBox.Show("Please ensure Job ID, Customer ID, and Reason are filled.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string customerEmail = "";
            string customerUsername = "";
            string senderEmail = "";
            try
            {
                con.Open();
                using (SqlCommand getJobCmd = new SqlCommand("SELECT Email, FirstName, CustomerID FROM Jobs WHERE JobID=@jobId", con))
                {
                    getJobCmd.Parameters.AddWithValue("@jobId", jobId);
                    using (SqlDataReader reader = getJobCmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            customerEmail = reader["Email"].ToString();
                            customerUsername = reader["FirstName"].ToString();
                        }
                    }
                }
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
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to fetch email information: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                con.Close();
            }

            int adminId = -1;
            try
            {
                con.Open();
                using (SqlCommand getAdminIdCmd = new SqlCommand("SELECT AdminID FROM Admins WHERE Email=@adminEmail", con))
                {
                    getAdminIdCmd.Parameters.AddWithValue("@adminEmail", senderEmail);
                    object adminIdObj = getAdminIdCmd.ExecuteScalar();
                    if (adminIdObj != null && adminIdObj != DBNull.Value)
                        adminId = Convert.ToInt32(adminIdObj);
                }
                using (SqlCommand insertCmd = new SqlCommand("INSERT INTO JobRejections (JobID, AdminID, Reason) VALUES (@jobId, @adminId, @reason)", con))
                {
                    insertCmd.Parameters.AddWithValue("@jobId", jobId);
                    insertCmd.Parameters.AddWithValue("@adminId", adminId);
                    insertCmd.Parameters.AddWithValue("@reason", reason);
                    insertCmd.ExecuteNonQuery();
                }
                using (SqlCommand updateCmd = new SqlCommand("UPDATE Jobs SET Status = @status WHERE JobID = @jobId", con))
                {
                    updateCmd.Parameters.AddWithValue("@status", "Rejected");
                    updateCmd.Parameters.AddWithValue("@jobId", jobId);
                    updateCmd.ExecuteNonQuery();
                }
            }
            catch (Exception dbEx)
            {
                MessageBox.Show("Failed to record rejection or delete job in database: " + dbEx.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                con.Close();
            }

            try
            {
                if (!string.IsNullOrEmpty(customerEmail) && !string.IsNullOrEmpty(senderEmail))
                {
                    System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
                    mail.From = new System.Net.Mail.MailAddress(senderEmail);
                    mail.To.Add(customerEmail);
                    mail.Subject = "E-Shift - Job Request Rejected";
                    mail.Body = $"Dear {customerUsername},\n\nYour request (Job ID: {jobId}) was cancelled by E-Shift Company.\n\nReason for rejection: {reason}\n\nIf you have any questions, please contact support.\n\nBest regards,\nE-Shift Team";

                    System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("smtp.gmail.com", 587);
                    smtp.EnableSsl = true;
                    smtp.Credentials = new System.Net.NetworkCredential(senderEmail, Environment.GetEnvironmentVariable("ESHIFT_SMTP_APP_PASSWORD"));
                    smtp.Send(mail);
                    MessageBox.Show("Rejection email sent successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Missing customer or sender email.", "Email Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception emailEx)
            {
                MessageBox.Show("Failed to send rejection email: " + emailEx.Message, "Email Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tnCancelRejecting_AdminRejectRequestPopUp_Click(object sender, EventArgs e)
        {

            this.Hide();
        }
    }
}
