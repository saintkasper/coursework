using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace courseworkwarehouse.Forms
{
    /// <summary>
    /// Логика взаимодействия для FormEditSupplier.xaml
    /// </summary>
    public partial class FormEditSupplier : Window
    {
        string ID = "";
        public FormEditSupplier(string ID = "")
        {
            this.ID = ID;
            InitializeComponent();
            DataTable Table = new DataTable();
            MySqlDataAdapter Adapter;
            if (ID != null)
            {
                Adapter = new MySqlDataAdapter("SELECT Title, Description, Address, Phone_number FROM Supplier WHERE Supplier_num = " + ID, MySql.connection);
                Adapter.Fill(Table);
                tbTitleSupplier.Text = Table.Rows[0][0].ToString();
                tbDescription.Text = Table.Rows[0][1].ToString();
                tbAddress.Text = Table.Rows[0][2].ToString();
                tbPhoneSupplier.Text = Table.Rows[0][3].ToString();
            }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (ID != null)
                {
                    MySql.connection.Open();
                    MySqlCommand commandabonent = new MySqlCommand("UPDATE Supplier SET Title = '" + tbTitleSupplier.Text + "', Address = '" + tbAddress.Text + "', Phone_number = '" + tbPhoneSupplier.Text + "', Description = '" + tbDescription.Text + "' WHERE Supplier_num= " + ID, MySql.connection);
                    commandabonent.ExecuteNonQuery();

                    MessageBox.Show("Данные изменены!");
                    MySql.connection.Close();

                }
                else
                {

                    MySql.connection.Open();
                    MySqlCommand commandabonent;
                    if (tbDescription.Text == "")
                    {
                        commandabonent = new MySqlCommand($"Insert Into Supplier (Title, Description, Address, Phone_number) Values ('{tbTitleSupplier.Text}', NULL, '{tbAddress.Text}', '{tbPhoneSupplier.Text}')", MySql.connection);
                    }
                    else
                    {
                        commandabonent = new MySqlCommand($"Insert Into Supplier (Title, Description, Address, Phone_number) Values ('{tbTitleSupplier.Text}', '{tbDescription.Text}', '{tbAddress.Text}', '{tbPhoneSupplier.Text}')", MySql.connection);
                    }
                    commandabonent.ExecuteNonQuery();
                    MessageBox.Show("Данные добавлены!");
                    MySql.connection.Close();
                }
            }
            catch { MessageBox.Show("Данные введены неверно!", "Ошибка"); }
        }

        private void btnBackSupplier_Click(object sender, RoutedEventArgs e)
        {
            if (DataBank.user == "admin")
            {
                MySql.connection.Open();
                DataBank.formAdmin.LoadDataReport();
                DataBank.formAdmin.LoadDataProduct();
                DataBank.formAdmin.LoadDataSupplier();
                DataBank.countGroups = 1;
                MySql.connection.Close();
            }
            else if (DataBank.user == "manager")
            {
                MySql.connection.Open();
                DataBank.formManager.LoadDataInvoice();
                DataBank.formManager.LoadDataProduct();
                DataBank.formManager.LoadDataSupplier();
                DataBank.countGroups = 1;
                MySql.connection.Close();
            }
            Close();
        }
    }
}
