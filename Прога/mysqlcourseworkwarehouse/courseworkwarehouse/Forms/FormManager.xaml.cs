using Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using Window = Microsoft.Office.Interop.Excel.Window;
using MySql.Data.MySqlClient;

namespace courseworkwarehouse.Forms
{
    /// <summary>
    /// Логика взаимодействия для FormManager.xaml
    /// </summary>
    public partial class FormManager : Window
    {
        string Filter = "";
        string AscDesc = "";
        int k = 0;

        public FormManager()
        {
            InitializeComponent();
            DataBank.formManager = this;
            SMassegeManager.Content = $"Добро пожаловать {DataBank.userFIO}!";
            System.Windows.Threading.DispatcherTimer dispatcherTimer = new System.Windows.Threading.DispatcherTimer();
            dispatcherTimer.Interval = TimeSpan.FromSeconds(1);
            dispatcherTimer.Tick += new EventHandler(dispatcherTimer_Tick);
            dispatcherTimer.Start();
            if (k == 3)
            {
                dispatcherTimer.Stop();
            }
        }

        private void dispatcherTimer_Tick(object sender, EventArgs e)
        {
            if (k == 3)
            {
                SnackManager.IsActive = false;
            }
            else
            {
                SnackManager.IsActive = true;
                k++;
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            DataBank.user = "";
            Authorization authorization = new Authorization();
            Hide();
            authorization.ShowDialog();
        }

        #region Product
        public void LoadDataProduct()
        {
            DataTable table = new DataTable();
            MySqlDataAdapter adapter = new MySqlDataAdapter($"select Product.Product_article as 'Артикул товара', Product.Title as 'Название', Product.Cost as 'Цена', Product.Quantity as 'Количество', Product.Supplier_num as 'Код поставщика', Product.Product_group_num as 'Код товарной группы', Supplier.Title as 'Поставщик', `product group`.Title as 'Товарная группа', Photo From Product inner join Supplier on Product.Supplier_num = Supplier.Supplier_num inner join `product group` on `product group`.Product_group_num = Product.Product_group_num WHERE Product.Title Like '%{tbSearchProduct.Text}%' {Filter} {AscDesc}", MySql.connection);
            adapter.Fill(table);
            dgProduct.ItemsSource = table.DefaultView;

            if (DataBank.countGroups != 2)
                cbRefresh();
        }

        private void btnExitProduc_Click(object sender, RoutedEventArgs e)
        {
            DataBank.user = "";
            Authorization authorization = new Authorization();
            Hide();
            authorization.ShowDialog();
        }


        private void cbRefresh()
        {
            int currentCount = cbProductGroup.Items.Count;

            if (cbProductGroup.Items.Count != 0)
            {
                if (cbProductGroup.Items.Count != currentCount || DataBank.countGroups == 1)
                {
                    int cbCount = cbProductGroup.Items.Count - 1;
                    for (int i = 0; i < cbCount; i++)
                    {
                        cbProductGroup.Items.RemoveAt(cbCount - i);
                    }
                }
                DataTable table = new DataTable();
                MySqlDataAdapter adapter = new MySqlDataAdapter($"select distinct Product.Product_group_num, `product group`.Title From Product join `product group` on `product group`.Product_group_num = Product.Product_group_num ", MySql.connection);
                adapter.Fill(table);

                for (int i = 0, c = 1; i < table.Rows.Count; i++, c++)
                {
                    cbProductGroup.Items.Insert(c, table.Rows[i][1].ToString());
                }
            }
        }

        private void tbSearchProduct_TextChanged(object sender, TextChangedEventArgs e)
        {
            LoadDataProduct();
        }

        private void rbWithoutProduct_Checked(object sender, RoutedEventArgs e)
        {
            if (rbWithoutProduct.IsChecked == true)
            {
                AscDesc = "";
                LoadDataProduct();
            }
        }

        private void rbAscProduct_Checked(object sender, RoutedEventArgs e)
        {
            if (rbAscProduct.IsChecked == true)
            {
                AscDesc = "ORDER BY Product.Cost ASC";
                LoadDataProduct();
            }
        }

        private void rbDescProduct_Checked(object sender, RoutedEventArgs e)
        {
            if (rbDescProduct.IsChecked == true)
            {
                AscDesc = "ORDER BY Product.Cost Desc";
                LoadDataProduct();
            }
        }

        private void btnAddProduct_Click(object sender, RoutedEventArgs e)
        {
            FormEditProduct editProduct = new FormEditProduct(null);
            editProduct.Show();
        }

        private void btnDeleteProduct_Click(object sender, RoutedEventArgs e)
        {
            if (dgProduct.SelectedItems.Count == 0)
            {
                MessageBox.Show("Выберите товар для удаления");
            }
            else
            {
                try
                {

                    DataBank.countGroups = 1;
                    MySql.connection.Open();

                    DataRowView row = (DataRowView)dgProduct.SelectedItems[0];

                    MySqlCommand commandcity = new MySqlCommand("DELETE FROM Product WHERE Product_article = " + row["Артикул товара"].ToString(), MySql.connection);
                    commandcity.ExecuteNonQuery();

                    MySql.connection.Close();

                    MessageBox.Show("Данные удалены!");
                    LoadDataProduct();
                }
                catch
                {
                    MessageBox.Show("В программе хранятся записи с этим товаром, при удалении, также удалятся и старые чеки");
                }
            }
        }

        private void btnEditProduct_Click(object sender, RoutedEventArgs e)
        {
            DataRowView row = (DataRowView)dgProduct.SelectedItems[0];
            FormEditProduct editProduct = new FormEditProduct(row[0].ToString());
            editProduct.Show();
        }

        private void tiProduct_Loaded(object sender, RoutedEventArgs e)
        {
            LoadDataProduct();
        }

        private void cbProductGroup_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (tbSearchProduct != null)
            {
                if (cbProductGroup.SelectedIndex != 0)
                {
                    DataBank.countGroups = 2;
                    Filter = $" And `product group`.Title = '{cbProductGroup.SelectedItem}'";
                    LoadDataProduct();
                }
                else
                {
                    DataBank.countGroups = 2;
                    Filter = "";
                    LoadDataProduct();
                }
            }
        }
        #endregion

        #region Supplier
        public void LoadDataSupplier()
        {
            DataTable table = new DataTable();
            MySqlDataAdapter adapter = new MySqlDataAdapter($"select Supplier_num as 'Код поставщика', Title as 'Название', Description as 'Описание', Address as 'Адрес', Phone_number as 'Номер телефона' From Supplier WHERE Title Like '%{tbSearchSupplier.Text}%' {AscDesc}", MySql.connection);
            adapter.Fill(table);
            dgSupplier.ItemsSource = table.DefaultView;
        }

        private void btnExitSupplier_Click(object sender, RoutedEventArgs e)
        {
            DataBank.user = "";
            Authorization authorization = new Authorization();
            Hide();
            authorization.ShowDialog();
        }

        private void EditSupplier_Click(object sender, RoutedEventArgs e)
        {
            DataRowView row = (DataRowView)dgSupplier.SelectedItems[0];
            FormEditSupplier editSupplier = new FormEditSupplier(row[0].ToString());
            editSupplier.Show();
        }

        private void btnAddSupplier_Click(object sender, RoutedEventArgs e)
        {
            FormEditSupplier editSupplier = new FormEditSupplier(null);
            editSupplier.Show();
        }

        private void btnDeleteSupplier_Click(object sender, RoutedEventArgs e)
        {
            if (dgSupplier.SelectedItems.Count == 0)
            {
                MessageBox.Show("Выберите товар для удаления");
            }
            else
            {
                try
                {

                    MySql.connection.Open();

                    DataRowView row = (DataRowView)dgSupplier.SelectedItems[0];

                    MySqlCommand commandcity = new MySqlCommand("DELETE FROM Supplier WHERE Supplier_num = " + row["Код поставщика"].ToString(), MySql.connection);
                    commandcity.ExecuteNonQuery();

                    MySql.connection.Close();

                    MessageBox.Show("Данные удалены!");
                    LoadDataSupplier();
                }
                catch
                {
                    MessageBox.Show("В программе хранятся записи с этой товарной группой!");
                }
            }
        }

        private void tbSearchSupplier_TextChanged(object sender, TextChangedEventArgs e)
        {
            LoadDataSupplier();
        }

        private void tiSupplier_Loaded(object sender, RoutedEventArgs e)
        {
            LoadDataSupplier();
        }

        private void rbWithoutSupplier_Checked(object sender, RoutedEventArgs e)
        {
            if (rbWithoutSupplier.IsChecked == true)
            {
                AscDesc = "";
                LoadDataSupplier();
            }
        }

        private void rbAscSupplier_Checked(object sender, RoutedEventArgs e)
        {
            if (rbAscSupplier.IsChecked == true)
            {
                AscDesc = "ORDER BY Title ASC";
                LoadDataSupplier();
            }
        }

        private void rbDescSupplier_Checked(object sender, RoutedEventArgs e)
        {
            if (rbDescSupplier.IsChecked == true)
            {
                AscDesc = "ORDER BY Title Desc";
                LoadDataSupplier();
            }
        }

        #endregion    

        #region ProductGroup
        public void LoadDataProductGroup()
        {
            DataTable table = new DataTable();
            MySqlDataAdapter adapter = new MySqlDataAdapter($"Select `product group`.Product_group_num as 'Код товарной группы', `product group`.Title as 'Название' From `product group` WHERE Title Like '%{tbSearchProductGroup.Text}%' {AscDesc}", MySql.connection);
            adapter.Fill(table);
            dgProductGroup.ItemsSource = table.DefaultView;
        }

        private void btnExitProductGroup_Click(object sender, RoutedEventArgs e)
        {
            DataBank.user = "";
            Authorization authorization = new Authorization();
            Hide();
            authorization.ShowDialog();
        }

        private void tiProductGroup_Loaded(object sender, RoutedEventArgs e)
        {
            LoadDataProductGroup();
        }

        private void tbSearchProductGroup_TextChanged(object sender, TextChangedEventArgs e)
        {
            LoadDataProductGroup();
        }

        private void rbWithoutProductGroup_Checked(object sender, RoutedEventArgs e)
        {
            if (rbWithoutProductGroup.IsChecked == true)
            {
                AscDesc = "";
                LoadDataProductGroup();
            }
        }

        private void rbAscProductGroup_Checked(object sender, RoutedEventArgs e)
        {
            if (rbAscProductGroup.IsChecked == true)
            {
                AscDesc = "Order By Title Asc";
                LoadDataProductGroup();
            }
        }

        private void rbDescProductGroup_Checked(object sender, RoutedEventArgs e)
        {
            if (rbDescProductGroup.IsChecked == true)
            {
                AscDesc = "Order By Title Desc";
                LoadDataProductGroup();
            }
        }

        private void btnAddProduct_ClickGroup(object sender, RoutedEventArgs e)
        {
            FormEditProductGroup editGroup = new FormEditProductGroup(null);
            editGroup.Show();
        }

        private void btnDeleteProductGroup_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dgProductGroup.SelectedItems.Count == 0)
                {
                    MessageBox.Show("Выберите товарную группу для удаления");
                }
                else
                {
                    MySql.connection.Open();

                    DataRowView row = (DataRowView)dgProductGroup.SelectedItems[0];

                    MySqlCommand commandcity = new MySqlCommand("DELETE FROM `product group` WHERE Product_group_num = " + row["Код товарной группы"].ToString(), MySql.connection);
                    commandcity.ExecuteNonQuery();

                    MySql.connection.Close();

                    MessageBox.Show("Данные удалены!");
                    LoadDataProductGroup();
                }
            }
            catch
            {
                MessageBox.Show("Что-то пошло не так!\n Возможно вы пытаетесь удалить товарную группу, в которой находятся товары", "Ошибка");
            }
        }

        private void btnEditProductGroup_Click(object sender, RoutedEventArgs e)
        {
            DataRowView row = (DataRowView)dgProductGroup.SelectedItems[0];
            FormEditProductGroup editGroup = new FormEditProductGroup(row[0].ToString());
            editGroup.Show();
        }
        #endregion

        #region Receipt invoice
        private void tiInvoice_Loaded(object sender, RoutedEventArgs e)
        {
            LoadDataInvoice();
        }

        private void btnExitReport_Click(object sender, RoutedEventArgs e)
        {
            DataBank.user = "";
            Authorization authorization = new Authorization();
            Hide();
            authorization.ShowDialog();
        }

        public void LoadDataInvoice()
        {
            //MySql.connection.Open();
            DataTable table = new DataTable();
            MySqlDataAdapter adapter = new MySqlDataAdapter($"Select Product_article as 'Артикул товара', Title as 'Название', Cost as 'Цена', Quantity as 'Количество', Photo as 'Фото' From Product WHERE Title Like '%{tbSearchReport.Text}%' {AscDesc}", MySql.connection);
            adapter.Fill(table);
            dgReportProducts.ItemsSource = table.DefaultView;
            MySql.connection.Close();
        }

        private void btnAddReportProduct1_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dgReportKorzina.ItemsSource == null)
                {
                    DataRowView row = (DataRowView)dgReportProducts.SelectedItems[0];
                    int quantityInSclad = int.Parse(row[3].ToString());
                    if (quantityInSclad > 0)
                    {
                        int pQuantity = int.Parse(row[3].ToString());
                        string pArticle = row[0].ToString();
                        int pCost = int.Parse(row[2].ToString());

                        DataTable table = new DataTable();
                        MySqlDataAdapter adapter = new MySqlDataAdapter($"Select Product_article as 'Артикул товара', Title as 'Название', Cost as 'Цена', Quantity as 'Количество', Photo as 'Фото' From Product WHERE Product_article = '{pArticle}'", MySql.connection);
                        adapter.Fill(table);

                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            table.Rows[i][3] = 1;
                        }
                        dgReportKorzina.ItemsSource = table.DefaultView;
                    }
                }
                else
                {
                    bool have = false;
                    bool kolvo = false;
                    DataTable dtAll = new DataTable();
                    DataTable Korzina = ((DataView)dgReportKorzina.ItemsSource).ToTable();

                    DataRowView row = (DataRowView)dgReportProducts.SelectedItems[0];

                    string pArticle = row[0].ToString();
                    int pCost = int.Parse(row[2].ToString());
                    int quantityInSclad = int.Parse(row[3].ToString()); // Количество на складе

                    DataTable table = new DataTable();
                    MySqlDataAdapter adapter = new MySqlDataAdapter($"Select Product_article as 'Артикул товара', Title as 'Название', Cost as 'Цена', Quantity as 'Количество', Photo as 'Фото' From Product WHERE Product_article = '{pArticle}'", MySql.connection);
                    adapter.Fill(table);
                    for (int i = 0; i < Korzina.Rows.Count; i++)
                    {
                        if (Korzina.Rows[i][3].ToString() == quantityInSclad.ToString())
                        {
                            kolvo = true;
                        }
                    }
                    if (kolvo == false)
                    {
                        for (int i = 0; i < Korzina.Rows.Count; i++)
                        {
                            if (Korzina.Rows[i][0].ToString() == pArticle)
                            {
                                have = true;
                                Korzina.Rows[i][3] = int.Parse(Korzina.Rows[i][3].ToString()) + 1;
                                break;
                            }
                        }

                        if (have == true)
                        {
                            dgReportKorzina.ItemsSource = Korzina.DefaultView;
                        }
                        if (have == false)
                        {
                            for (int i = 0; i < table.Rows.Count; i++)
                            {
                                table.Rows[i][3] = 1;
                            }
                            dtAll = Korzina.Copy();
                            dtAll.Merge(table);
                            dgReportKorzina.ItemsSource = dtAll.DefaultView;
                        }
                    }
                }
            }
            catch
            {

            }
        }

        private void btnAddReportProduct10_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dgReportKorzina.ItemsSource == null)
                {
                    DataRowView row = (DataRowView)dgReportProducts.SelectedItems[0];
                    int quantityInSclad = int.Parse(row[3].ToString());

                    if (quantityInSclad >= 10)
                    {
                        string pArticle = row[0].ToString();
                        int pCost = int.Parse(row[2].ToString());

                        DataTable table = new DataTable();
                        MySqlDataAdapter adapter = new MySqlDataAdapter($"Select Product_article as 'Артикул товара', Title as 'Название', Cost as 'Цена', Quantity as 'Количество', Photo as 'Фото' From Product WHERE Product_article = '{pArticle}'", MySql.connection);
                        adapter.Fill(table);

                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            table.Rows[i][3] = 10;
                        }
                        dgReportKorzina.ItemsSource = table.DefaultView;
                    }
                }
                else
                {
                    bool kolvo = false;
                    bool have = false;
                    DataTable dtAll = new DataTable();
                    DataTable Korzina = ((DataView)dgReportKorzina.ItemsSource).ToTable();

                    DataRowView row = (DataRowView)dgReportProducts.SelectedItems[0];

                    string pArticle = row[0].ToString();
                    int pCost = int.Parse(row[2].ToString());
                    int quantityInSclad = int.Parse(row[3].ToString());

                    DataTable table = new DataTable();
                    MySqlDataAdapter adapter = new MySqlDataAdapter($"Select Product_article as 'Артикул товара', Title as 'Название', Cost as 'Цена', Quantity as 'Количество', Photo as 'Фото' From Product WHERE Product_article = '{pArticle}'", MySql.connection);
                    adapter.Fill(table);

                    for (int i = 0; i < Korzina.Rows.Count; i++)
                    {
                        if (Korzina.Rows[i][0].ToString() == pArticle)
                        {
                            have = true;
                            break;
                        }
                    }
                    if (have == false)
                    {
                        if (quantityInSclad >= 10)
                        {
                            for (int i = 0; i < table.Rows.Count; i++)
                            {
                                table.Rows[i][3] = 10;
                            }
                            dtAll = Korzina.Copy();
                            dtAll.Merge(table);
                            dgReportKorzina.ItemsSource = dtAll.DefaultView;
                        }

                    }
                    else
                    {
                        for (int i = 0; i < Korzina.Rows.Count; i++)
                        {
                            if (Korzina.Rows[i][0].ToString() == pArticle)
                            {
                                if ((int.Parse(Korzina.Rows[i][3].ToString()) + 10) > quantityInSclad)
                                {
                                    kolvo = true;
                                }
                            }

                        }
                        if (kolvo == false)
                        {
                            for (int i = 0; i < Korzina.Rows.Count; i++)
                            {
                                if (Korzina.Rows[i][0].ToString() == pArticle)
                                {
                                    have = true;
                                    Korzina.Rows[i][3] = int.Parse(Korzina.Rows[i][3].ToString()) + 10;
                                    break;
                                }
                            }
                            if (have == true)
                            {
                                dgReportKorzina.ItemsSource = Korzina.DefaultView;
                            }
                            if (have == false)
                            {
                                for (int i = 0; i < table.Rows.Count; i++)
                                {
                                    table.Rows[i][3] = 10;
                                }
                                dtAll = Korzina.Copy();
                                dtAll.Merge(table);
                                dgReportKorzina.ItemsSource = dtAll.DefaultView;
                            }
                        }
                    }

                    //for (int i = 0; i < Korzina.Rows.Count; i++)
                    //{
                    //    if (Korzina.Rows[i][0].ToString() == pArticle)
                    //    {
                    //        if ((int.Parse(Korzina.Rows[i][3].ToString()) + 10) >= quantityInSclad)
                    //        {
                    //            kolvo = true;
                    //        }
                    //    }

                    //}
                    //if (kolvo == false)
                    //{
                    //    for (int i = 0; i < Korzina.Rows.Count; i++)
                    //    {
                    //        if (Korzina.Rows[i][0].ToString() == pArticle)
                    //        {
                    //            have = true;
                    //            Korzina.Rows[i][3] = int.Parse(Korzina.Rows[i][3].ToString()) + 10;
                    //            break;
                    //        }
                    //    }

                    //    if (have == true)
                    //    {
                    //        dgReportKorzina.ItemsSource = Korzina.DefaultView;
                    //    }
                    //    if (have == false)
                    //    {
                    //        for (int i = 0; i < table.Rows.Count; i++)
                    //        {
                    //            table.Rows[i][3] = 10;
                    //        }
                    //        dtAll = Korzina.Copy();
                    //        dtAll.Merge(table);
                    //        dgReportKorzina.ItemsSource = dtAll.DefaultView;
                    //    }
                    //}
                }
            }
            catch
            {

            }

        }

        private void btnAddReportProduct100_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dgReportKorzina.ItemsSource == null)
                {
                    DataRowView row = (DataRowView)dgReportProducts.SelectedItems[0];
                    int quantityInSclad = int.Parse(row[3].ToString());
                    if (quantityInSclad >= 100)
                    {
                        string pArticle = row[0].ToString();
                        int pCost = int.Parse(row[2].ToString());

                        DataTable table = new DataTable();
                        MySqlDataAdapter adapter = new MySqlDataAdapter($"Select Product_article as 'Артикул товара', Title as 'Название', Cost as 'Цена', Quantity as 'Количество', Photo as 'Фото' From Product WHERE Product_article = '{pArticle}'", MySql.connection);
                        adapter.Fill(table);

                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            table.Rows[i][3] = 100;
                        }
                        dgReportKorzina.ItemsSource = table.DefaultView;
                    }
                }
                else
                {
                    bool kolvo = false;
                    bool have = false;
                    DataTable dtAll = new DataTable();
                    DataTable Korzina = ((DataView)dgReportKorzina.ItemsSource).ToTable();

                    DataRowView row = (DataRowView)dgReportProducts.SelectedItems[0];

                    string pArticle = row[0].ToString();
                    int pCost = int.Parse(row[2].ToString());
                    int quantityInSclad = int.Parse(row[3].ToString());

                    DataTable table = new DataTable();
                    MySqlDataAdapter adapter = new MySqlDataAdapter($"Select Product_article as 'Артикул товара', Title as 'Название', Cost as 'Цена', Quantity as 'Количество', Photo as 'Фото' From Product WHERE Product_article = '{pArticle}'", MySql.connection);
                    adapter.Fill(table);

                    for (int i = 0; i < Korzina.Rows.Count; i++)
                    {
                        if (Korzina.Rows[i][0].ToString() == pArticle)
                        {
                            have = true;
                            break;
                        }
                    }
                    if (have == false)
                    {
                        if (quantityInSclad >= 100)
                        {
                            for (int i = 0; i < table.Rows.Count; i++)
                            {
                                table.Rows[i][3] = 100;
                            }
                            dtAll = Korzina.Copy();
                            dtAll.Merge(table);
                            dgReportKorzina.ItemsSource = dtAll.DefaultView;
                        }

                    }
                    else
                    {
                        for (int i = 0; i < Korzina.Rows.Count; i++)
                        {
                            if (Korzina.Rows[i][0].ToString() == pArticle)
                            {
                                if ((int.Parse(Korzina.Rows[i][3].ToString()) + 100) > quantityInSclad)
                                {
                                    kolvo = true;
                                }
                            }

                        }
                        if (kolvo == false)
                        {
                            for (int i = 0; i < Korzina.Rows.Count; i++)
                            {
                                if (Korzina.Rows[i][0].ToString() == pArticle)
                                {
                                    have = true;
                                    Korzina.Rows[i][3] = int.Parse(Korzina.Rows[i][3].ToString()) + 100;
                                    break;
                                }
                            }
                            if (have == true)
                            {
                                dgReportKorzina.ItemsSource = Korzina.DefaultView;
                            }
                            if (have == false)
                            {
                                for (int i = 0; i < table.Rows.Count; i++)
                                {
                                    table.Rows[i][3] = 100;
                                }
                                dtAll = Korzina.Copy();
                                dtAll.Merge(table);
                                dgReportKorzina.ItemsSource = dtAll.DefaultView;
                            }
                        }
                    }




                    //for (int i = 0; i < Korzina.Rows.Count; i++)
                    //{
                    //    if (Korzina.Rows[i][0].ToString() == pArticle)
                    //    {
                    //        if ((int.Parse(Korzina.Rows[i][3].ToString()) + 100) >= quantityInSclad)
                    //        {
                    //            kolvo = true;
                    //        }
                    //    }

                    //}
                    //if (kolvo == false)
                    //{
                    //    for (int i = 0; i < Korzina.Rows.Count; i++)
                    //    {
                    //        if (Korzina.Rows[i][0].ToString() == pArticle)
                    //        {
                    //            have = true;
                    //            Korzina.Rows[i][3] = int.Parse(Korzina.Rows[i][3].ToString()) + 100;
                    //            break;
                    //        }
                    //    }
                    //    if (have == true)
                    //    {
                    //        dgReportKorzina.ItemsSource = Korzina.DefaultView;
                    //    }
                    //    if (have == false)
                    //    {
                    //        for (int i = 0; i < table.Rows.Count; i++)
                    //        {
                    //            table.Rows[i][3] = 100;
                    //        }
                    //        dtAll = Korzina.Copy();
                    //        dtAll.Merge(table);
                    //        dgReportKorzina.ItemsSource = dtAll.DefaultView;
                    //    }
                    //}
                }
            }
            catch
            {

            }

        }

        private void btnDeleteReportProduct1_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                DataTable Korzina = ((DataView)dgReportKorzina.ItemsSource).ToTable();

                DataRowView row = (DataRowView)dgReportKorzina.SelectedItems[0];
                int CurrentRowID = dgReportKorzina.SelectedIndex;

                string pKorzinaQuantity = row[3].ToString();

                if (pKorzinaQuantity == 1.ToString())
                {
                    for (int i = 0; i < Korzina.Rows.Count; i++)
                    {
                        if (i == CurrentRowID)
                        {
                            Korzina.Rows.RemoveAt(i);
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < Korzina.Rows.Count; i++)
                    {
                        if (i == CurrentRowID)
                        {
                            Korzina.Rows[i][3] = int.Parse(Korzina.Rows[i][3].ToString()) - 1;
                        }
                    }
                }
                dgReportKorzina.ItemsSource = Korzina.DefaultView;
            }
            catch
            {

            }

        }

        private void btnDeleteReportProduct10_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                DataTable Korzina = ((DataView)dgReportKorzina.ItemsSource).ToTable();

                DataRowView row = (DataRowView)dgReportKorzina.SelectedItems[0];
                int CurrentRowID = dgReportKorzina.SelectedIndex;

                string pKorzinaQuantity = row[3].ToString();

                if (int.Parse(pKorzinaQuantity) <= 10)
                {
                    for (int i = 0; i < Korzina.Rows.Count; i++)
                    {
                        if (i == CurrentRowID)
                        {
                            Korzina.Rows.RemoveAt(i);
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < Korzina.Rows.Count; i++)
                    {
                        if (i == CurrentRowID)
                        {
                            Korzina.Rows[i][3] = int.Parse(Korzina.Rows[i][3].ToString()) - 10;
                        }
                    }
                }
                dgReportKorzina.ItemsSource = Korzina.DefaultView;
            }
            catch
            {

            }

        }

        private void btnDeleteReportProduct100_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                DataTable Korzina = ((DataView)dgReportKorzina.ItemsSource).ToTable();

                DataRowView row = (DataRowView)dgReportKorzina.SelectedItems[0];
                int CurrentRowID = dgReportKorzina.SelectedIndex;

                string pKorzinaQuantity = row[3].ToString();

                if (int.Parse(pKorzinaQuantity) <= 100)
                {
                    for (int i = 0; i < Korzina.Rows.Count; i++)
                    {
                        if (i == CurrentRowID)
                        {
                            Korzina.Rows.RemoveAt(i);
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < Korzina.Rows.Count; i++)
                    {
                        if (i == CurrentRowID)
                        {
                            Korzina.Rows[i][3] = int.Parse(Korzina.Rows[i][3].ToString()) - 100;
                        }
                    }
                }
                dgReportKorzina.ItemsSource = Korzina.DefaultView;
            }
            catch
            {

            }

        }

        private void btnClear_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                dgReportKorzina.ItemsSource = null;
            }
            catch { }

        }

        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            if (dgReportKorzina.ItemsSource != null)
            {
                bool receiptnum = false;
                Excel.Application excel = new Excel.Application();
                Workbook workbook = excel.Workbooks.Open(@"C:\Users\furuh\Documents\ReceiptInvoice.xlsx");
                Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
                DataTable max = new DataTable();
                MySqlDataAdapter adapterMax = new MySqlDataAdapter("SELECT MAX(`tabular part of the receipt`.Receipt_invoice_num) as 'Код' FROM`tabular part of the receipt`", MySql.connection);
                adapterMax.Fill(max);

                DataTable Korzina = ((DataView)dgReportKorzina.ItemsSource).ToTable();
                Korzina.Columns.RemoveAt(4);
                Korzina.Columns.Add("Цена со скидкой(руб.)", typeof(double));
                Korzina.Columns.Add("Итого(руб.)", typeof(double));

                for (int i = 0; i < Korzina.Rows.Count; i++)
                {
                    if (int.Parse(Korzina.Rows[i][2].ToString()) > 500000 || (int.Parse(Korzina.Rows[i][2].ToString()) * int.Parse(Korzina.Rows[i][3].ToString())) > 500000)
                    {
                        Korzina.Rows[i][4] = (double.Parse(Korzina.Rows[i][2].ToString()) - (double.Parse(Korzina.Rows[i][2].ToString()) * 0.17)).ToString();
                    }
                    else if (int.Parse(Korzina.Rows[i][2].ToString()) > 100000 || (int.Parse(Korzina.Rows[i][2].ToString()) * int.Parse(Korzina.Rows[i][3].ToString())) > 100000)
                    {
                        Korzina.Rows[i][4] = (double.Parse(Korzina.Rows[i][2].ToString()) - (double.Parse(Korzina.Rows[i][2].ToString()) * 0.15)).ToString();
                    }
                    else if (int.Parse(Korzina.Rows[i][2].ToString()) > 10000 || (int.Parse(Korzina.Rows[i][2].ToString()) * int.Parse(Korzina.Rows[i][3].ToString())) > 10000)
                    {
                        Korzina.Rows[i][4] = (double.Parse(Korzina.Rows[i][2].ToString()) - (double.Parse(Korzina.Rows[i][2].ToString()) * 0.1)).ToString();
                    }
                    else Korzina.Rows[i][4] = Korzina.Rows[i][2].ToString();

                    Korzina.Rows[i][5] = (double.Parse(Korzina.Rows[i][4].ToString()) * double.Parse(Korzina.Rows[i][3].ToString())).ToString();
                }
                string receiptInvoiceNum = "";
                Range myRange4 = (Range)sheet1.Cells[5, 1];
                try
                {
                    myRange4.Value2 += $"{int.Parse(max.Rows[0][0].ToString()) + 1}";
                    receiptnum = true;
                }
                catch
                {
                    myRange4.Value2 += 1;
                    receiptInvoiceNum = "1";
                }

                Range myRange3 = (Range)sheet1.Cells[6, 4];
                myRange3.Value2 += DateTime.Now.ToShortDateString();
                for (int j = 0; j < Korzina.Columns.Count; j++)
                {
                    if (j == 0)
                    {
                        Range myRange = (Range)sheet1.Cells[10, j + 1];
                        sheet1.Cells[10, j + 1].Font.Bold = true;
                        myRange.Value2 = "№ Товара";
                    }
                    else
                    {
                        Range myRange = (Range)sheet1.Cells[10, j + 1]; // Откуда начинаем
                        sheet1.Cells[10, j + 1].Font.Bold = true; // В этих ячейках ставим жирный шрифт
                        myRange.Value2 = Korzina.Columns[j].ColumnName; // Добавляем названия колонок ([j].ColumnName) в первую строку 
                    }
                }

                int nomer = 1;

                for (int i = 0; i < Korzina.Rows.Count; i++)
                {
                    for (int j = 1; j < Korzina.Columns.Count; j++)
                    {
                        string b = Korzina.Rows[i][j].ToString();              // Суем в строку b элементы по порядку (Сначала всю первую строку, вторую и т.д.)
                        Range myRange = (Range)sheet1.Cells[i + 11, j + 1];     // Указываем в какой позиции будем заполнять (в экселе индексы идут с 1, а не с 0), т.к. в 1-ой строке названия столбцов, то начинаем со второй
                        myRange.Value2 = b;
                        myRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                        myRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        myRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        myRange = (Range)sheet1.Cells[i + 11, 1];
                        myRange.Value2 = nomer;
                        myRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                        myRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        myRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    }
                    nomer++;
                }
                double EndCost = 0;
                for (int i = 0; i < Korzina.Rows.Count; i++)
                {
                    EndCost += double.Parse(Korzina.Rows[i][5].ToString());
                }

                Range myRange7 = (Range)sheet1.Cells[Korzina.Rows.Count + 11, 6];
                myRange7.Value2 = $"{EndCost}(руб.)";

                Range myRange2 = sheet1.get_Range((Range)sheet1.Cells[Korzina.Rows.Count + 12, 1], (Range)sheet1.Cells[Korzina.Rows.Count + 12, 5]);
                myRange2.Merge();
                myRange2.Value2 = $"Сдал: {DataBank.userFIO}";


                //Range myRange5 = (Range)sheet1.Cells[Korzina.Rows.Count + 11, 1];
                Range myRange5 = sheet1.get_Range((Range)sheet1.Cells[Korzina.Rows.Count + 13, 1], (Range)sheet1.Cells[Korzina.Rows.Count + 13, 3]);
                myRange5.Merge();
                myRange5.Value2 = "Подпись:_____________";
                Range myRange6 = sheet1.UsedRange;
                myRange6.Columns.AutoFit();
                excel.Visible = true; // Открываем эксель

                MySql.connection.Open();
                DataTable dataQuantityInSklad = new DataTable();
                MySqlDataAdapter MySqlDataAdapter;
                DataTable dtPA = new DataTable();
                MySqlCommand adapter;
                MySqlDataAdapter ProdArticle;


                for (int i = 0; i < Korzina.Rows.Count; i++)
                {
                    if (receiptnum == true)
                    {
                        adapter = new MySqlCommand($"Insert Into `tabular part of the receipt` (Receipt_invoice_num, Product_article, Quantity, Wholesale_price, Cost, Employee_num, Date_of_the_receipt) Values ('{int.Parse(max.Rows[0][0].ToString()) + 1}', '{Korzina.Rows[i][0].ToString()}','{Korzina.Rows[i][3].ToString()}', '{Korzina.Rows[i][4].ToString()}', '{Korzina.Rows[i][5].ToString()}', '{DataBank.employeenum}', '{DateTime.Today.ToShortDateString()}')", MySql.connection);
                        adapter.ExecuteNonQuery();
                        ProdArticle = new MySqlDataAdapter("SELECT `tabular part of the receipt`.Tabular_part_of_the_receipt_num, `tabular part of the receipt`.Product_article FROM `tabular part of the receipt`", MySql.connection);
                        ProdArticle.Fill(dtPA);
                        MySqlDataAdapter = new MySqlDataAdapter($"Select Quantity From Product Where Product_article = '{dtPA.Rows[dtPA.Rows.Count - 1][1].ToString()}'", MySql.connection);
                        MySqlDataAdapter.Fill(dataQuantityInSklad);
                        MySqlCommand Update = new MySqlCommand($"Update Product Set Quantity = '{(int.Parse(dataQuantityInSklad.Rows[i][0].ToString())) - (int.Parse(Korzina.Rows[i][3].ToString()))}' Where Product_article = '{Korzina.Rows[i][0].ToString()}'", MySql.connection);
                        Update.ExecuteNonQuery();
                    }
                    else
                    {
                        adapter = new MySqlCommand($"Insert Into `tabular part of the receipt` (Receipt_invoice_num, Product_article, Quantity, Wholesale_price, Cost, Employee_num, Date_of_the_receipt) Values ('{receiptInvoiceNum}', '{Korzina.Rows[i][0].ToString()}', '{Korzina.Rows[i][3].ToString()}', '{Korzina.Rows[i][4].ToString()}', '{Korzina.Rows[i][5].ToString()}', '{DataBank.employeenum}', '{DateTime.Today.ToShortDateString()}')", MySql.connection);
                        adapter.ExecuteNonQuery();
                        ProdArticle = new MySqlDataAdapter("SELECT `tabular part of the receipt`.Tabular_part_of_the_receipt_num, `tabular part of the receipt`.Product_article FROM `tabular part of the receipt`", MySql.connection);
                        ProdArticle.Fill(dtPA);
                        MySqlDataAdapter = new MySqlDataAdapter($"Select Quantity From Product Where Product_article = '{dtPA.Rows[dtPA.Rows.Count - 1][1].ToString()}'", MySql.connection);
                        MySqlDataAdapter.Fill(dataQuantityInSklad);
                        MySqlCommand Update = new MySqlCommand($"Update Product Set Quantity = '{(int.Parse(dataQuantityInSklad.Rows[i][0].ToString())) - (int.Parse(Korzina.Rows[i][3].ToString()))}' Where Product_article = '{Korzina.Rows[i][0].ToString()}'", MySql.connection);
                        Update.ExecuteNonQuery();
                    }
                }

                dgReportKorzina.ItemsSource = null;
                LoadDataInvoice();
                MySql.connection.Close();




                //DataTable dt = new DataTable();
                //MySqlDataAdapter common = new MySqlDataAdapter("Select TP_invoice_num From [Receipt invoice]", MySql.connection);
                //common.Fill(dt);
                //int TPreceipt = 1;
                //try
                //{
                //    TPreceipt = (int.Parse(dt.Rows[dt.Rows.Count - 1][0].ToString())) + 1;
                //}
                //catch {  }
                //MySql.connection.Open();
                //for (int i = 0; i < Korzina.Rows.Count; i++)
                //{
                //    MySqlCommand command = new MySqlCommand($"Insert into [Receipt invoice] (TP_invoice_num, Product_article, Date_of_the_receipt) Values ('{TPreceipt}', '{Korzina.Rows[i][0].ToString()}', '{DateTime.Today.ToShortDateString()}')", MySql.connection);
                //    command.ExecuteNonQuery();
                //}
                //DataTable dt2 = new DataTable();
                //MySqlDataAdapter common2 = new MySqlDataAdapter("Select Receipt_invoice_num From [Receipt invoice]", MySql.connection);
                //common2.Fill(dt2);

                //int receiptNum = 1;

                //for (int i = 0; i < Korzina.Rows.Count; i++)
                //{

                //    //MySqlCommand command = new MySqlCommand($"Insert into [Receipt invoice] (TP_invoice_num, Product_article, Date_of_the_receipt) Values ('{TPreceipt}', '{Korzina.Rows[i][0].ToString()}', '{DateTime.Today.ToShortDateString()}')", MySql.connection);
                //    //command.ExecuteNonQuery();
                //    try
                //    {
                //        receiptNum = int.Parse(dt2.Rows[dt2.Rows.Count - 1][0].ToString());
                //    }
                //    catch { receiptNum = int.Parse(dt2.Rows[dt2.Rows.Count][0].ToString()); }
                //    MySqlCommand command2 = new MySqlCommand($"Insert into `tabular part of the receipt` (Receipt_invoice_num, Quantity, Cost, Wholesale_price, Employee_num) Values ('{receiptNum}', '{Korzina.Rows[i][3].ToString()}', '{Korzina.Rows[i][2].ToString()}', '{Korzina.Rows[i][4].ToString()}', '{DataBank.employeenum}')", MySql.connection);
                //    command2.ExecuteNonQuery();
                //}
                //MySql.connection.Close();
            }
            else MessageBox.Show("Добавьте товары в корзину для составления расходной накладной");

        }

        private void btnDeleteReportRow_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataTable Korzina = ((DataView)dgReportKorzina.ItemsSource).ToTable();

                int CurrentRowID = dgReportKorzina.SelectedIndex;

                Korzina.Rows.RemoveAt(CurrentRowID);
                dgReportKorzina.ItemsSource = Korzina.DefaultView;
                if (Korzina.Rows.Count == 0)
                {
                    try
                    {
                        dgReportKorzina.ItemsSource = null;
                    }
                    catch { }
                }
            }
            catch { }
        }

        private void tbSearchReport_TextChanged(object sender, TextChangedEventArgs e)
        {
            LoadDataInvoice();
        }
        #endregion

        #region Excel
        dynamic Window.Activate()
        {
            throw new NotImplementedException();
        }
        dynamic Window.ActivateNext()
        {
            throw new NotImplementedException();
        }
        dynamic Window.ActivatePrevious()
        {
            throw new NotImplementedException();
        }
        bool Window.Close(object SaveChanges, object Filename, object RouteWorkbook)
        {
            throw new NotImplementedException();
        }
        dynamic Window.LargeScroll(object Down, object Up, object ToRight, object ToLeft)
        {
            throw new NotImplementedException();
        }
        Window Window.NewWindow()
        {
            throw new NotImplementedException();
        }
        dynamic Window._PrintOut(object From, object To, object Copies, object Preview, object ActivePrinter, object PrintToFile, object Collate, object PrToFileName)
        {
            throw new NotImplementedException();
        }
        dynamic Window.PrintPreview(object EnableChanges)
        {
            throw new NotImplementedException();
        }
        dynamic Window.ScrollWorkbookTabs(object Sheets, object Position)
        {
            throw new NotImplementedException();
        }
        dynamic Window.SmallScroll(object Down, object Up, object ToRight, object ToLeft)
        {
            throw new NotImplementedException();
        }
        int Window.PointsToScreenPixelsX(int Points)
        {
            throw new NotImplementedException();
        }
        int Window.PointsToScreenPixelsY(int Points)
        {
            throw new NotImplementedException();
        }
        dynamic Window.RangeFromPoint(int x, int y)
        {
            throw new NotImplementedException();
        }
        void Window.ScrollIntoView(int Left, int Top, int Width, int Height, object Start)
        {
            throw new NotImplementedException();
        }
        dynamic Window.PrintOut(object From, object To, object Copies, object Preview, object ActivePrinter, object PrintToFile, object Collate, object PrToFileName)
        {
            throw new NotImplementedException();
        }
        Excel.Application Window.Application => throw new NotImplementedException();
        XlCreator Window.Creator => throw new NotImplementedException();
        dynamic Window.Parent => throw new NotImplementedException();
        Range Window.ActiveCell => throw new NotImplementedException();
        Chart Window.ActiveChart => throw new NotImplementedException();
        Pane Window.ActivePane => throw new NotImplementedException();
        dynamic Window.ActiveSheet => throw new NotImplementedException();
        dynamic Window.Caption { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool Window.DisplayFormulas { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool Window.DisplayGridlines { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool Window.DisplayHeadings { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool Window.DisplayHorizontalScrollBar { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool Window.DisplayOutline { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool Window._DisplayRightToLeft { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool Window.DisplayVerticalScrollBar { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool Window.DisplayWorkbookTabs { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool Window.DisplayZeros { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool Window.EnableResize { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool Window.FreezePanes { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        int Window.GridlineColor { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        XlColorIndex Window.GridlineColorIndex { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        double Window.Height { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        int Window.Index => throw new NotImplementedException();
        double Window.Left { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        string Window.OnWindow { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        Panes Window.Panes => throw new NotImplementedException();
        Range Window.RangeSelection => throw new NotImplementedException();
        int Window.ScrollColumn { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        int Window.ScrollRow { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        Sheets Window.SelectedSheets => throw new NotImplementedException();
        dynamic Window.Selection => throw new NotImplementedException();
        bool Window.Split { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        int Window.SplitColumn { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        double Window.SplitHorizontal { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        int Window.SplitRow { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        double Window.SplitVertical { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        double Window.TabRatio { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        double Window.Top { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        XlWindowType Window.Type => throw new NotImplementedException();
        double Window.UsableHeight => throw new NotImplementedException();
        double Window.UsableWidth => throw new NotImplementedException();
        bool Window.Visible { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        Range Window.VisibleRange => throw new NotImplementedException();
        double Window.Width { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        int Window.WindowNumber => throw new NotImplementedException();
        XlWindowState Window.WindowState { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        dynamic Window.Zoom { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        XlWindowView Window.View { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool Window.DisplayRightToLeft { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        SheetViews Window.SheetViews => throw new NotImplementedException();
        dynamic Window.ActiveSheetView => throw new NotImplementedException();
        bool Window.DisplayRuler { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool Window.AutoFilterDateGrouping { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        bool Window.DisplayWhitespace { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        int Window.Hwnd => throw new NotImplementedException();
        #endregion

    }
}