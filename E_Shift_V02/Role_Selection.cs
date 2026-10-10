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
    public partial class Role_Selection : Form
    {
        public Role_Selection()
        {
            InitializeComponent();
        }
        SqlConnection con = new SqlConnection("Data Source=DISON_PC\\SQLEXPRESS;Initial Catalog=ESHIFT;Integrated Security=True;Encrypt=False");

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            try
            {
                con.Open();
                SqlCommand cmd = new SqlCommand("SELECT RoleName FROM RoleSelection", con);
                SqlDataReader reader = cmd.ExecuteReader();
                ComboBoxSelectYourRole_RoleSelectionForm.Items.Clear();
                while (reader.Read())
                {
                    ComboBoxSelectYourRole_RoleSelectionForm.Items.Add(reader["RoleName"].ToString());
                }
                reader.Close();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading roles:\n" + ex.Message, "Database Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnContinue_RoleSelectionForm_Click(object sender, EventArgs e)
        {
            if (ComboBoxSelectYourRole_RoleSelectionForm.SelectedItem == null)
            {
                MessageBox.Show("Please select a role to continue.", "Missing Selection",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string selectedRole = ComboBoxSelectYourRole_RoleSelectionForm.SelectedItem.ToString();

            if (selectedRole == "Admin")
            {
                Admin_Login adminLogin = new Admin_Login();
                adminLogin.Show();
                this.Hide();
            }
            else if (selectedRole == "Customer")
            {
                Customer_Login customerLogin = new Customer_Login();
                customerLogin.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Unknown role selected. Please contact support.", "Role Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
    }
}
