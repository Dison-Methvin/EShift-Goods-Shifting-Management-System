using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace E_Shift_V02
{
    public partial class Customer_SignUp : Form
    {
        private readonly SqlConnection con = new SqlConnection("Data Source=DISON_PC\\SQLEXPRESS;Initial Catalog=ESHIFT;Integrated Security=True;Encrypt=False");

        public Customer_SignUp()
        {
            InitializeComponent();
            txtStrngPassword_CustomerSiginUpForm.PasswordChar = '●';
            txtConfirmPassword_CustomerSiginUpForm.PasswordChar = '●';
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            // Populate title combobox with common titles and those from DB
            ComboBoxTitle_CustomerSiginUpForm.Items.Clear();
            ComboBoxTitle_CustomerSiginUpForm.Items.AddRange(new object[] { "Mr", "Mrs", "Miss", "Ms", "Dr", "Prof" });
            txtStrngPassword_CustomerSiginUpForm.PasswordChar = '●';
            txtConfirmPassword_CustomerSiginUpForm.PasswordChar = '●';
            checkBoxShowPassword_CustomerSiginUpForm.CheckedChanged += checkBoxShowPassword_CustomerSiginUpForm_CheckedChanged;
            try
            {
                con.Open();
                SqlCommand cmd = new SqlCommand("SELECT DISTINCT Title FROM Customers WHERE Title IS NOT NULL AND Title <> ''", con);
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string dbTitle = reader["Title"].ToString();
                    if (!ComboBoxTitle_CustomerSiginUpForm.Items.Contains(dbTitle))
                        ComboBoxTitle_CustomerSiginUpForm.Items.Add(dbTitle);
                }
                reader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading titles:\n" + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
            }
        }

        private void linkLabelLogin_CustomerSiginUpForm_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Customer_Login signUpForm = new Customer_Login();
            signUpForm.Show();
            this.Hide();
        }

        // Handles customer sign up button click
        private void btnSiginUp_CustomerSiginUpForm_Click(object sender, EventArgs e)
        {
            string title = ComboBoxTitle_CustomerSiginUpForm.SelectedItem != null ? ComboBoxTitle_CustomerSiginUpForm.SelectedItem.ToString() : string.Empty;
            string firstName = txtFName_CustomerSiginUpForm.Text.Trim();
            string lastName = txtLName_CustomerSiginUpForm.Text.Trim();
            string gender = RadioButtonMale_CustomerSiginUpForm.Checked ? "Male" : (RadioButtonFemale_CustomerSiginUpForm.Checked ? "Female" : "");
            string address = txtAddress_CustomerSiginUpForm.Text.Trim();
            string phoneNumber = txtPNumber_CustomerSiginUpForm.Text.Trim();
            string email = txtPreferredEmail_CustomerSiginUpForm.Text.Trim();
            string nic = txtNIC_CustomerSiginUpForm.Text.Trim();
            string username = txtPreferredUsername_CustomerSiginUpForm.Text.Trim();
            string password = txtStrngPassword_CustomerSiginUpForm.Text;
            string confirmPassword = txtConfirmPassword_CustomerSiginUpForm.Text;
            byte[] profilePhoto = null;

            // Validate required fields
            if (string.IsNullOrEmpty(title) || string.IsNullOrEmpty(firstName) || string.IsNullOrEmpty(lastName) || string.IsNullOrEmpty(gender) ||
                string.IsNullOrEmpty(address) || string.IsNullOrEmpty(phoneNumber) || string.IsNullOrEmpty(email) ||
                string.IsNullOrEmpty(nic) || string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
            {
                MessageBox.Show("Please fill in all required fields.", "Missing Data", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (password != confirmPassword)
            {
                MessageBox.Show("Passwords do not match.", "Password Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string hashedPassword = GetSha256Hash(password);

            // Convert image to byte array if provided
            if (pictureBoxAddCustomerPic_CustomerSiginUpForm.Image != null)
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    pictureBoxAddCustomerPic_CustomerSiginUpForm.Image.Save(ms, pictureBoxAddCustomerPic_CustomerSiginUpForm.Image.RawFormat);
                    profilePhoto = ms.ToArray();
                }
            }

            try
            {
                con.Open();
                // Check for duplicate values
                if (IsDuplicate("PhoneNumber", phoneNumber) || IsDuplicate("NIC", nic) ||
                    IsDuplicate("Username", username) || IsDuplicate("Email", email))
                {
                    MessageBox.Show("Duplicate entry detected. Please use unique values for phone, NIC, username, and email.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Insert new customer
                SqlCommand cmd = new SqlCommand("INSERT INTO Customers (Title, FirstName, LastName, Gender, Address, PhoneNumber, Email, NIC, Username, PasswordHash, ProfilePhoto, RegisterDate, IsActive) VALUES (@Title, @FirstName, @LastName, @Gender, @Address, @PhoneNumber, @Email, @NIC, @Username, @PasswordHash, @ProfilePhoto, GETDATE(), 1)", con);
                cmd.Parameters.AddWithValue("@Title", title);
                cmd.Parameters.AddWithValue("@FirstName", firstName);
                cmd.Parameters.AddWithValue("@LastName", lastName);
                cmd.Parameters.AddWithValue("@Gender", gender);
                cmd.Parameters.AddWithValue("@Address", address);
                cmd.Parameters.AddWithValue("@PhoneNumber", phoneNumber);
                cmd.Parameters.AddWithValue("@Email", email);
                cmd.Parameters.AddWithValue("@NIC", nic);
                cmd.Parameters.AddWithValue("@Username", username);
                cmd.Parameters.AddWithValue("@PasswordHash", hashedPassword);
                cmd.Parameters.AddWithValue("@ProfilePhoto", (object)profilePhoto ?? DBNull.Value);
                int rows = cmd.ExecuteNonQuery();
                if (rows > 0)
                {
                    SendWelcomeEmail(email, firstName);
                    MessageBox.Show("Registration successful! You can now log in.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Customer_Login loginForm = new Customer_Login();
                    loginForm.Show();
                    this.Hide();
                }
                else
                {
                    MessageBox.Show("Registration failed. Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        // Checks for duplicate values in the Customers table
        private bool IsDuplicate(string column, string value)
        {
            using (SqlCommand cmd = new SqlCommand($"SELECT COUNT(*) FROM Customers WHERE {column}=@value", con))
            {
                cmd.Parameters.AddWithValue("@value", value);
                return (int)cmd.ExecuteScalar() > 0;
            }
        }

        // Sends a welcome email to the new customer
        private void SendWelcomeEmail(string email, string firstName)
        {
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
            try
            {
                if (!string.IsNullOrEmpty(email) && !string.IsNullOrEmpty(senderEmail))
                {
                    System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
                    mail.From = new System.Net.Mail.MailAddress(senderEmail);
                    mail.To.Add(email);
                    mail.Subject = "Welcome to E-Shift!";
                    mail.Body = $"Dear {firstName},\n\nWelcome to E-Shift! Thank you for registering. You can now log in and enjoy our services. If you have any questions, feel free to contact us.\n\nBest regards,\nE-Shift Team";

                    System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("smtp.gmail.com", 587);
                    smtp.EnableSsl = true;
                    smtp.Credentials = new System.Net.NetworkCredential(senderEmail, Environment.GetEnvironmentVariable("ESHIFT_SMTP_APP_PASSWORD"));
                    smtp.Send(mail);
                }
            }
            catch (Exception emailEx)
            {
                MessageBox.Show("Registration successful, but failed to send welcome email: " + emailEx.Message, "Email Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
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

        // Handles profile photo selection
        private void pictureBoxAddCustomerPic_CustomerSiginUpForm_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Select Profile Photo";
            ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                pictureBoxAddCustomerPic_CustomerSiginUpForm.Image = Image.FromFile(ofd.FileName);
            }
        }

        // Toggle password visibility
        private void checkBoxShowPassword_CustomerSiginUpForm_CheckedChanged(object sender, EventArgs e)
        {
            bool show = checkBoxShowPassword_CustomerSiginUpForm.Checked;
            if (show)
            {
                txtStrngPassword_CustomerSiginUpForm.UseSystemPasswordChar = false;
                txtStrngPassword_CustomerSiginUpForm.PasswordChar = '\0';
                txtConfirmPassword_CustomerSiginUpForm.UseSystemPasswordChar = false;
                txtConfirmPassword_CustomerSiginUpForm.PasswordChar = '\0';
            }
            else
            {
                txtStrngPassword_CustomerSiginUpForm.UseSystemPasswordChar = true;
                txtStrngPassword_CustomerSiginUpForm.PasswordChar = '●';
                txtConfirmPassword_CustomerSiginUpForm.UseSystemPasswordChar = true;
                txtConfirmPassword_CustomerSiginUpForm.PasswordChar = '●';
            }
        }

        // Cancels sign up and returns to role selection
        private void btnancel_CustomerSiginUpForm_Click(object sender, EventArgs e)
        {
            Role_Selection signUpForm = new Role_Selection();
            signUpForm.Show();
            this.Hide();
        }

        private void btnClear_CustomerSiginUpForm_Click(object sender, EventArgs e)
        {
            txtFName_CustomerSiginUpForm.Text = string.Empty;
            txtLName_CustomerSiginUpForm.Text = string.Empty;
            txtAddress_CustomerSiginUpForm.Text = string.Empty;
            txtPNumber_CustomerSiginUpForm.Text = string.Empty;
            txtPreferredEmail_CustomerSiginUpForm.Text = string.Empty;
            txtNIC_CustomerSiginUpForm.Text = string.Empty;
            txtPreferredUsername_CustomerSiginUpForm.Text = string.Empty;
            txtStrngPassword_CustomerSiginUpForm.Text = string.Empty;
            txtConfirmPassword_CustomerSiginUpForm.Text = string.Empty;
            RadioButtonMale_CustomerSiginUpForm.Checked = false;
            RadioButtonFemale_CustomerSiginUpForm.Checked = false;
            ComboBoxTitle_CustomerSiginUpForm.SelectedIndex = -1;
            pictureBoxAddCustomerPic_CustomerSiginUpForm.Image = null;
            checkBoxShowPassword_CustomerSiginUpForm.Checked = false;
        }
    }
}
