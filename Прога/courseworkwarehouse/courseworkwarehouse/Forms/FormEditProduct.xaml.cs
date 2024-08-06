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
            SqlDataAdapter Adapter;
            DataBank.sqlConnection.Close();
            DataBank.sqlConnection.Open();
            DataTable TableSuppliers = new DataTable();
            SqlDataAdapter SqlDataSupplier = new SqlDataAdapter("Select Supplier_num, Title From Supplier", DataBank.sqlConnection);
            SqlDataSupplier.Fill(TableSuppliers);
            DataTable TableGroups = new DataTable();
            SqlDataAdapter SqlDataGroups = new SqlDataAdapter("Select Product_group_num, Title From [Product group]", DataBank.sqlConnection);
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
                Adapter = new SqlDataAdapter("SELECT Title, Cost, Quantity, Photo, Supplier_num, Product_group_num FROM Product WHERE Product_article = " + ID, DataBank.sqlConnection);
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

                DataBank.sqlConnection.Close();
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
                    string filename = dialog.FileName;
                    imagePhoto.Source = new BitmapImage(new Uri(filename, UriKind.Absolute));
                    tbPhoto.Text = filename;
                }
            }
            catch (Exception)
            {

            }
        }

        private void btnSaveProduct_Click(object sender, RoutedEventArgs e)
        {
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
                        DataBank.sqlConnection.Open();
                        SqlCommand commandabonent = new SqlCommand("UPDATE Product SET Title = '" + tbTitleProduct.Text + "', Cost = '" + tbCostProduct.Text + "', Quantity = '" + tbQuantityProduct.Text + "', Photo = '" + tbPhoto.Text + "', Supplier_num = '" + supplierid + "', Product_group_num = '" + groupid + "' WHERE Product_article = " + ID, DataBank.sqlConnection);
                        commandabonent.ExecuteNonQuery();

                        MessageBox.Show("Данные изменены!");
                        DataBank.sqlConnection.Close();
                    }
                    else
                    {
                        if (tbTitleProduct.Text == "" || tbCostProduct.Text == "" || tbQuantityProduct.Text == "" || cbSuppliers.Text == "" || cbProductGroups.Text == "")
                        {
                            MessageBox.Show("Данные не введены", "Ошибка");
                        }
                        DataBank.sqlConnection.Open();
                        SqlCommand commandabonent = new SqlCommand("UPDATE Product SET Title = '" + tbTitleProduct.Text + "', Cost = '" + tbCostProduct.Text + "', Quantity = '" + tbQuantityProduct.Text + "', Photo = NULL, Supplier_num = '" + supplierid + "', Product_group_num = '" + groupid + "' WHERE Product_article = " + ID, DataBank.sqlConnection);
                        commandabonent.ExecuteNonQuery();

                        MessageBox.Show("Данные изменены!");
                        DataBank.sqlConnection.Close();
                    }
                }
                else
                {
                    SqlCommand commandabonent = new SqlCommand($"Insert Into Product (Title, Cost, Quantity, Photo, Supplier_num, Product_group_num) Values ('{tbTitleProduct.Text}', '{tbCostProduct.Text}', '{tbQuantityProduct.Text}', '{tbPhoto.Text}', '{supplierid}', '{groupid}')", DataBank.sqlConnection);
                    commandabonent.ExecuteNonQuery();

                    MessageBox.Show("Данные добавлены!");
                    DataBank.sqlConnection.Close();
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
