using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace E_Shift_V02
{
    public partial class Customer_Login : Form
    {
        private readonly SqlConnection con = new SqlConnection("Data Source=DISON_PC\\SQLEXPRESS;Initial Catalog=ESHIFT;Integrated Security=True;Encrypt=False");

        public Customer_Login()
        {
            InitializeComponent();
            txtpassword_CustomerLoginForm.PasswordChar = '●'; // Hide password by default
        }

        // Hashes a string using SHA256 (for password security)
        private string GetSha256Hash(string input)
        {
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] bytes = Encoding.UTF8.GetBytes(input);
                byte[] hash = sha256.ComputeHash(bytes);
                StringBuilder builder = new StringBuilder();
                foreach (byte b in hash)
                {
                    builder.Append(b.ToString("x2"));
                }
                return builder.ToString();
            }
        }

        // Toggle password visibility
        private void checkBoxShowPassword_CustomerLoginForm_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxShowPassword_CustomerLoginForm.Checked)
            {
                txtpassword_CustomerLoginForm.UseSystemPasswordChar = false;
                txtpassword_CustomerLoginForm.PasswordChar = '\0';
            }
            else
            {
                txtpassword_CustomerLoginForm.UseSystemPasswordChar = true;
                txtpassword_CustomerLoginForm.PasswordChar = '●';
            }
        }

        // Handles customer login button click
        private void btnSiginIn_CustomerLoginForm_Click(object sender, EventArgs e)
        {
            string username = txtUsername_CustomerLoginForm.Text.Trim();
            string password = txtpassword_CustomerLoginForm.Text;

            if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
            {
                MessageBox.Show("Please enter both username and password.", "Missing Credentials",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string hashedPassword = GetSha256Hash(password);

            try
            {
                con.Open();
                using (SqlCommand cmd = new SqlCommand("SELECT COUNT(*) FROM Customers WHERE Username=@username AND PasswordHash=@password", con))
                {
                    cmd.Parameters.AddWithValue("@username", username);
                    cmd.Parameters.AddWithValue("@password", hashedPassword);
                    int count = (int)cmd.ExecuteScalar();
                    if (count == 1)
                    {
                        // Fetch customer email and name for notification
                        string customerEmail = "";
                        string customerName = "Customer";
                        using (SqlCommand getEmailCmd = new SqlCommand("SELECT Email, FirstName FROM Customers WHERE Username=@username", con))
                        {
                            getEmailCmd.Parameters.AddWithValue("@username", username);
                            using (SqlDataReader reader = getEmailCmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    customerEmail = reader["Email"].ToString();
                                    customerName = reader["FirstName"].ToString();
                                }
                            }
                        }

                        // Fetch sender (admin) email
                        string senderEmail = "";
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

                        // Send login notification email
                        try
                        {
                            if (!string.IsNullOrEmpty(customerEmail) && !string.IsNullOrEmpty(senderEmail))
                            {
                                System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
                                mail.From = new System.Net.Mail.MailAddress(senderEmail);
                                mail.To.Add(customerEmail);
                                mail.Subject = "Welcome to E-Shift - Login Detected";
                                mail.Body = $"Dear {customerName},\n\nWelcome to the E-Shift system!\n\nA login to your account was detected. If this was not you, please change your password immediately or contact support.\n\nThank you for using E-Shift!\n\nBest regards,\nE-Shift Team";

                                System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("smtp.gmail.com", 587);
                                smtp.EnableSsl = true;
                                smtp.Credentials = new System.Net.NetworkCredential(senderEmail, Environment.GetEnvironmentVariable("ESHIFT_SMTP_APP_PASSWORD"));
                                smtp.Send(mail);
                            }
                        }
                        catch (Exception emailEx)
                        {
                            MessageBox.Show("Login successful, but failed to send email notification: " + emailEx.Message, "Email Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }

                        DialogResult result = MessageBox.Show("Login successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        if (result == DialogResult.OK)
                        {
                            Customer_Dashboard dashboard = new Customer_Dashboard(username);
                            dashboard.Show();
                            this.Hide();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid username or password.", "Login Failed",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error connecting to the database:\n" + ex.Message, "Database Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        private void linkLabelSiginUp_CustomerLoginForm_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Customer_SignUp signUpForm = new Customer_SignUp();
            signUpForm.Show();
            this.Hide();
        }

        private void btnCancel_CustomerLoginForm_Click(object sender, EventArgs e)
        {
            Role_Selection signUpForm = new Role_Selection();
            signUpForm.Show();
            this.Hide();
        }

        private void btnClear_CustomerLoginForm_Click(object sender, EventArgs e)
        {
            txtUsername_CustomerLoginForm.Text = string.Empty;
            txtpassword_CustomerLoginForm.Text = string.Empty;
        }
    }
}
