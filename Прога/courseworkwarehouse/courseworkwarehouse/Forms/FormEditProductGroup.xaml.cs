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
            SqlDataAdapter Adapter;
            if (ID != null)
            {
                Adapter = new SqlDataAdapter("Select [Product group].Product_group_num as [Код товарной группы], [Product group].Title as [Название] From [Product group] WHERE Product_group_num = " + ID, DataBank.sqlConnection);
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
                    DataBank.sqlConnection.Open();
                    SqlCommand commandabonent = new SqlCommand("UPDATE [Product group] SET Title = '" + tbTitleProductGroup.Text + "' WHERE Product_group_num = " + ID, DataBank.sqlConnection);
                    commandabonent.ExecuteNonQuery();

                    MessageBox.Show("Данные изменены!");
                    DataBank.sqlConnection.Close();
                }
                else
                {

                    DataBank.sqlConnection.Open();
                    SqlCommand commandabonent;
                    commandabonent = new SqlCommand($"Insert Into [Product group] (Title) Values ('{tbTitleProductGroup.Text}')", DataBank.sqlConnection);

                    commandabonent.ExecuteNonQuery();
                    MessageBox.Show("Данные добавлены!");
                    DataBank.sqlConnection.Close();
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
