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
    /// Логика взаимодействия для FormEditProduct.xaml
    /// </summary>
    public partial class FormEditProduct : Window
    {
        string ID = "";
        List<string> SupplierList = new List<string>();
        List<string> GroupsList = new List<string>();


        public FormEditProduct(string ID = "")
        {
            this.ID = ID;
            InitializeComponent();
            DataTable Table = new DataTable();
            MySqlDataAdapter Adapter;
            MySql.connection.Close();
            MySql.connection.Open();
            DataTable TableSuppliers = new DataTable();
            MySqlDataAdapter SqlDataSupplier = new MySqlDataAdapter("Select Supplier_num, Title From Supplier", MySql.connection);
            SqlDataSupplier.Fill(TableSuppliers);
            DataTable TableGroups = new DataTable();
            MySqlDataAdapter SqlDataGroups = new MySqlDataAdapter("Select Product_group_num, Title From `product group`", MySql.connection);
            SqlDataGroups.Fill(TableGroups);
            for (int i = 0; i < TableGroups.Rows.Count; i++)
            {
                cbProductGroups.Items.Add(TableGroups.Rows[i][1].ToString());
                GroupsList.Add(TableGroups.Rows[i][0].ToString());
            }
            for (int i = 0; i < TableSuppliers.Rows.Count; i++)
            {
                cbSuppliers.Items.Add(TableSuppliers.Rows[i][1].ToString());
                SupplierList.Add(TableSuppliers.Rows[i][0].ToString());
            }
            if (ID != null)
            {
                Adapter = new MySqlDataAdapter("SELECT Title, Cost, Quantity, Photo, Supplier_num, Product_group_num FROM Product WHERE Product_article = " + ID, MySql.connection);
                Adapter.Fill(Table);
                tbTitleProduct.Text = Table.Rows[0][0].ToString();
                tbCostProduct.Text = Table.Rows[0][1].ToString();
                tbQuantityProduct.Text = Table.Rows[0][2].ToString();
                tbPhoto.Text = Table.Rows[0][3].ToString();
                for (int i = 0; i < cbProductGroups.Items.Count; i++)
                {
                    if (Table.Rows[0][5].ToString() == GroupsList[i])
                    {
                        cbProductGroups.SelectedIndex = i;
                        //cbProductGroups.Text = GroupsList[i];
                    }
                }
                for (int i = 0; i < cbSuppliers.Items.Count; i++)
                {
                    if (Table.Rows[0][4].ToString() == SupplierList[i])
                    {
                        cbSuppliers.SelectedIndex = i;
                        //cbSuppliers.Text = SupplierList[i];
                    }
                }
                if (tbPhoto.Text != "")
                    imagePhoto.Source = new BitmapImage(new Uri(Table.Rows[0][3].ToString(), UriKind.Absolute));

                MySql.connection.Close();
            }

        }

        private void btnSelectPhoto_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dialog = new Microsoft.Win32.OpenFileDialog();
                dialog.Filter = "JPG Files (*.jpg)|*.jpg|PNG Files (*.png)|*.png|All Files (*.*)|*.*"; // Filter files by extension

                bool? result = dialog.ShowDialog();

                if (result == true)
                {
                    DataBank.photopath = dialog.FileName;
                    string filename = dialog.FileName;
                    imagePhoto.Source = new BitmapImage(new Uri(filename, UriKind.Absolute));
                    tbPhoto.Text = filename;
                }
            }
            catch (Exception)
            {

            }

            /* Подсчет слешей в пути
            char bslash = '\\';
            int freq = DataBank.photopath.Count(f => (f == bslash));
            MessageBox.Show(freq.ToString() + "\n" + DataBank.photopath);
            */
        }

        private void btnSaveProduct_Click(object sender, RoutedEventArgs e)
        {
            string newphotopath = DataBank.photopath.Replace("\\", "\\\\");
            try
            {
                string groupid = "";
                string supplierid = "";
                for (int i = 0; i < cbProductGroups.Items.Count; i++)
                {
                    if (cbProductGroups.Text == cbProductGroups.Items[i].ToString())
                    {
                        groupid = GroupsList[i];
                    }
                }
                for (int i = 0; i < cbSuppliers.Items.Count; i++)
                {
                    if (cbSuppliers.Text == cbSuppliers.Items[i].ToString())
                    {
                        supplierid = SupplierList[i];
                    }
                }
                if (ID != null)
                {
                    if (tbPhoto.Text != "")
                    {
                        MySql.connection.Open();
                        MySqlCommand commandabonent = new MySqlCommand("UPDATE Product SET Title = '" + tbTitleProduct.Text + "', Cost = '" + tbCostProduct.Text + "', Quantity = '" + tbQuantityProduct.Text + "', Photo = '" + newphotopath + "', Supplier_num = '" + supplierid + "', Product_group_num = '" + groupid + "' WHERE Product_article = " + ID, MySql.connection);
                        commandabonent.ExecuteNonQuery();

                        MessageBox.Show("Данные изменены!");
                        MySql.connection.Close();
                    }
                    else
                    {
                        if (tbTitleProduct.Text == "" || tbCostProduct.Text == "" || tbQuantityProduct.Text == "" || cbSuppliers.Text == "" || cbProductGroups.Text == "")
                        {
                            MessageBox.Show("Данные не введены", "Ошибка");
                        }
                        MySql.connection.Open();
                        MySqlCommand commandabonent = new MySqlCommand("UPDATE Product SET Title = '" + tbTitleProduct.Text + "', Cost = '" + tbCostProduct.Text + "', Quantity = '" + tbQuantityProduct.Text + "', Photo = NULL, Supplier_num = '" + supplierid + "', Product_group_num = '" + groupid + "' WHERE Product_article = " + ID, MySql.connection);
                        commandabonent.ExecuteNonQuery();

                        MessageBox.Show("Данные изменены!");
                        MySql.connection.Close();
                    }
                }
                else
                {
                    MySqlCommand commandabonent = new MySqlCommand($"Insert Into Product (Title, Cost, Quantity, Photo, Supplier_num, Product_group_num) Values ('{tbTitleProduct.Text}', '{tbCostProduct.Text}', '{tbQuantityProduct.Text}', '{tbPhoto.Text}', '{supplierid}', '{groupid}')", MySql.connection);
                    commandabonent.ExecuteNonQuery();

                    MessageBox.Show("Данные добавлены!");
                    MySql.connection.Close();
                }
        }
            catch { MessageBox.Show("Данные введены неверно!", "Ошибка"); }
}

        private void btnBackProduct_Click(object sender, RoutedEventArgs e)
        {
            if (DataBank.user == "admin")
            {
                DataBank.formAdmin.LoadDataProduct();
                DataBank.countGroups = 1;
                DataBank.formAdmin.LoadDataReport();
            }
            else if (DataBank.user == "manager")
            {
                DataBank.formManager.LoadDataProduct();
                DataBank.countGroups = 1;
                DataBank.formManager.LoadDataInvoice();
            }
            Close();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            tbPhoto.Text = "";
        }

        private void tbCostProduct_PreviewTextInput(object sender, TextCompositionEventArgs e) // Ввод только цифв
        {
            e.Handled = !(Char.IsDigit(e.Text, 0));
        }
    }
}
