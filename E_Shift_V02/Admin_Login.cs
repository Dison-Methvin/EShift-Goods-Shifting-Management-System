using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Security.Cryptography;

namespace E_Shift_V02
{
    public partial class Admin_Login : Form
    {
        private readonly SqlConnection con = new SqlConnection("Data Source=DISON_PC\\SQLEXPRESS;Initial Catalog=ESHIFT;Integrated Security=True;Encrypt=False");

        public Admin_Login()
        {
            InitializeComponent();
        }

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

        private void Admin_Login_Load(object sender, EventArgs e)
        {
        }

        private void btnLogin_AdminLoginForm_Click(object sender, EventArgs e)
        {
            string username = txtUsername_AdminLoginForm.Text.Trim();
            string password = txtPassword_AdminLoginForm.Text;

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
                using (SqlCommand cmd = new SqlCommand("SELECT COUNT(*) FROM Admins WHERE Username=@username AND PasswordHash=@password", con))
                {
                    cmd.Parameters.AddWithValue("@username", username);
                    cmd.Parameters.AddWithValue("@password", hashedPassword);
                    int count = (int)cmd.ExecuteScalar();
                    if (count == 1)
                    {
                        DialogResult result = MessageBox.Show("Login successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        if (result == DialogResult.OK)
                        {
                            Admin_Dashboard dashboard = new Admin_Dashboard(username);
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

        // Clears login fields
        private void btnClear_AdminLoginForm_Click_1(object sender, EventArgs e)
        {
            txtUsername_AdminLoginForm.Text = string.Empty;
            txtPassword_AdminLoginForm.Text = string.Empty;
            MessageBox.Show("Fields cleared.", "Cleared", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Cancels login and returns to role selection
        private void btnCancel_AdminLoginForm_Click(object sender, EventArgs e)
        {
            Role_Selection signUpForm = new Role_Selection();
            signUpForm.Show();
            this.Hide();
        }

        // Opens the default dialer with the contact number
        private void linkLabelContactNumber_AdminLoginForm_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string phoneNumber = linkLabelContactNumber_AdminLoginForm.Text.Trim();
            try
            {
                System.Diagnostics.Process.Start($"tel:{phoneNumber}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to open the phone app. Please call manually.\n{phoneNumber}\nError: {ex.Message}", "Call Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
