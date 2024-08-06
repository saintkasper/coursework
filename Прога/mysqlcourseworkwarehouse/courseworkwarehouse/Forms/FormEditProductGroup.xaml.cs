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
    /// Логика взаимодействия для FormEditProductGroup.xaml
    /// </summary>
    public partial class FormEditProductGroup : Window
    {
        string ID = "";
        public FormEditProductGroup(string ID = "")
        {
            this.ID = ID;
            InitializeComponent();
            DataTable Table = new DataTable();
            MySqlDataAdapter Adapter;
            if (ID != null)
            {
                Adapter = new MySqlDataAdapter("Select `product group`.Product_group_num as 'Код товарной группы', `product group`.Title as 'Название' From `product group` WHERE Product_group_num = " + ID, MySql.connection);
                Adapter.Fill(Table);
                tbTitleProductGroup.Text = Table.Rows[0][1].ToString();
            }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (ID != null)
                {
                    MySql.connection.Open();
                    MySqlCommand commandabonent = new MySqlCommand("UPDATE `product group` SET Title = '" + tbTitleProductGroup.Text + "' WHERE Product_group_num = " + ID, MySql.connection);
                    commandabonent.ExecuteNonQuery();

                    MessageBox.Show("Данные изменены!");
                    MySql.connection.Close();
                }
                else
                {

                    MySql.connection.Open();
                    MySqlCommand commandabonent;
                    commandabonent = new MySqlCommand($"Insert Into `product group` (Title) Values ('{tbTitleProductGroup.Text}')", MySql.connection);

                    commandabonent.ExecuteNonQuery();
                    MessageBox.Show("Данные добавлены!");
                    MySql.connection.Close();
                }
            }
            catch { MessageBox.Show("Данные введены неверно!", "Ошибка"); }
        }

        private void btnBack_Click(object sender, RoutedEventArgs e)
        {
            if (DataBank.user == "admin")
            {
                DataBank.countGroups = 1;
                DataBank.formAdmin.LoadDataProductGroup();
                DataBank.formAdmin.LoadDataReport();
            }
            else if (DataBank.user == "manager")
            {
                DataBank.countGroups = 1;
                DataBank.formManager.LoadDataProductGroup();
                DataBank.formManager.LoadDataInvoice();
            }
            Close();
        }
    }
}
